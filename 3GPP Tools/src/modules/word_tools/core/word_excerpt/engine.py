"""Core isolated-process fragment extraction and multi-source Word merge engine."""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import pythoncom
import win32com.client
from docx.oxml.text.paragraph import CT_P

from core.utils.office_printer_guard import temporary_virtual_default_printer
from core.utils.paths import get_temp_root, to_extended_path
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import ensure_3gpp_package_styles
from .com_helpers import (
    apply_label_in_fresh_word,
    configure_word,
    copy_formatting_parts,
    is_rpc_disconnect,
    wait_for_file_unlocked,
)
from .scanner import HeadingInfo

logger = logging.getLogger(__name__)


def _resolve_facade_attr(name: str, default):
    facade = sys.modules.get("modules.word_tools.core.word_excerpt_extractor")
    return getattr(facade, name, default) if facade is not None else default


class WordExcerptEngine:
    """Extracts selected heading ranges into isolated fragments and merges them cleanly."""

    def __init__(
        self,
        source_paths: Sequence[str],
        headings: Sequence[HeadingInfo],
        progress_callback=None,
    ):
        self.source_paths = list(source_paths)
        self.headings = {h.id: h for h in headings}
        self.progress_callback = progress_callback

    def _progress(self, percent: int, message: str):
        if self.progress_callback:
            self.progress_callback(max(0, min(100, int(percent))), message)

    def _normalize(self, selected_ids: Iterable[str]) -> List[HeadingInfo]:
        chosen = [self.headings[x] for x in selected_ids if x in self.headings]
        chosen.sort(key=lambda h: (h.source_index, h.block_index))
        result = []
        for h in chosen:
            if any(
                p.source_index == h.source_index
                and p.block_index <= h.block_index < p.end_index
                for p in result
            ):
                continue
            result.append(h)
        return result

    @staticmethod
    def _safe_remove(block, body):
        if isinstance(block, CT_P):
            try:
                sect_pr = block.pPr.sectPr if block.pPr is not None else None
                if sect_pr is not None:
                    body.append(sect_pr)
            except Exception:
                pass
        parent = block.getparent()
        if parent is not None:
            parent.remove(block)

    @staticmethod
    def _merge_ranges(ranges: List[Tuple[int, int]], block_count: int) -> List[Tuple[int, int]]:
        normalized = sorted(
            (max(0, start), min(end, block_count))
            for start, end in ranges
            if start < end and start < block_count
        )
        merged: List[List[int]] = []
        for start, end in normalized:
            if not merged or start > merged[-1][1]:
                merged.append([start, end])
            else:
                merged[-1][1] = max(merged[-1][1], end)
        return [(start, end) for start, end in merged]

    def _make_fragment(self, source: Path, ranges: List[Tuple[int, int]], target: Path):
        """Build the expensive python-docx fragment in a separate process.

        Keeping this work outside the GUI process is intentional: very large 3GPP
        specifications can spend minutes serializing a DOCX package. A QThread
        still shares the application's Python interpreter/GIL, whereas this helper
        process cannot starve the Qt event loop. The helper uses the same
        formatting-safe python-docx body reconstruction as the in-process version.
        """
        helper = Path(__file__).resolve().parent.parent / "word_excerpt_fragment_worker.py"
        if not helper.is_file():
            raise RuntimeError(f"Excerpt fragment worker not found: {helper}")

        merged = self._merge_ranges(ranges, 2**31 - 1)
        logger.info(
            "   -> Starting isolated excerpt preparation for %s (%d selected range(s))",
            source.name,
            len(merged),
        )

        cmd = [
            sys.executable,
            str(helper),
            str(source),
            str(target),
            json.dumps(merged),
        ]
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=creationflags,
            check=False,
        )
        if result.stdout.strip():
            for line in result.stdout.splitlines():
                logger.info("   [excerpt worker] %s", line)
        if result.returncode != 0:
            detail = result.stderr.strip() or result.stdout.strip() or "Unknown worker failure"
            raise RuntimeError(
                f"Isolated excerpt preparation failed for {source.name}: {detail}"
            )
        if not target.exists() or target.stat().st_size == 0:
            raise RuntimeError(
                f"Isolated excerpt preparation produced no fragment for {source.name}."
            )

    def extract(self, selected_ids: Iterable[str], output_path: str) -> Path:
        selected = self._normalize(selected_ids)
        if not selected:
            raise ValueError("No headings selected.")

        grouped: Dict[str, List[Tuple[int, int]]] = {}
        source_order: List[str] = []
        for h in selected:
            if h.source_path not in grouped:
                grouped[h.source_path] = []
                source_order.append(h.source_path)
            grouped[h.source_path].append((h.block_index, h.end_index))

        output = Path(output_path).resolve()
        os.makedirs(to_extended_path(output.parent), exist_ok=True)
        work = Path(tempfile.mkdtemp(prefix="word_excerpt_", dir=str(get_temp_root())))
        fragments = []
        try:
            total = max(1, len(source_order))
            for idx, source_path in enumerate(source_order):
                self._progress(
                    5 + int(45 * idx / total),
                    f"Preparing excerpt from {Path(source_path).name}...",
                )
                fragment = work / f"fragment_{idx:03d}.docx"
                self._make_fragment(Path(source_path), grouped[source_path], fragment)
                fragments.append(fragment)
            if len(fragments) == 1:
                # A fragment produced from one source document is already a complete
                # formatting-preserving DOCX. Do not round-trip it through Word's
                # InsertFile/SaveAs2 path; that adds no value and has proven fragile
                # with managed Office installations.
                self._progress(55, "Finalizing single-document excerpt...")
                logger.info("   -> Single source fragment: skipping Microsoft Word merge.")
                working = fragments[0]

                # Apply the required corporate label in a fresh Word session and lock in 3GPP styles
                # inside the isolated temp working directory before copying to the final destination.
                self._progress(80, "Applying corporate sensitivity label and 3GPP styles...")
                pythoncom.CoInitialize()
                try:
                    self._apply_label_in_fresh_word(working)
                finally:
                    pythoncom.CoUninitialize()
                WordExcerptEngine._wait_for_file_unlocked(working)
                ensure_styles = _resolve_facade_attr(
                    "ensure_3gpp_package_styles", ensure_3gpp_package_styles
                )
                ensure_styles(working)

                # Only a completely finalized, closed file is copied to the user target.
                ext_output = to_extended_path(output)
                os.makedirs(to_extended_path(output.parent), exist_ok=True)
                if os.path.exists(ext_output):
                    try:
                        os.remove(ext_output)
                    except OSError:
                        pass
                shutil.copy2(to_extended_path(working), ext_output)
                if not os.path.exists(ext_output) or os.path.getsize(ext_output) == 0:
                    raise RuntimeError("Failed to create the single-document excerpt output.")
            else:
                self._progress(55, "Combining excerpts with Microsoft Word...")
                self._merge_with_word(fragments, output)

            self._progress(100, "Extraction complete.")
            return output
        finally:
            shutil.rmtree(work, ignore_errors=True)

    @staticmethod
    def _copy_formatting_parts(
        source: Path, target: Path, attempts: int = 12, delay: float = 0.5
    ):
        return copy_formatting_parts(source, target, attempts=attempts, delay=delay)

    @staticmethod
    def _wait_for_file_unlocked(
        file_path: Path, max_wait: float = 6.0, delay: float = 0.2
    ) -> bool:
        return wait_for_file_unlocked(file_path, max_wait=max_wait, delay=delay)

    @staticmethod
    def _configure_word(word):
        configure_word(word)

    @staticmethod
    def _is_rpc_disconnect(exc: BaseException) -> bool:
        return is_rpc_disconnect(exc)

    @staticmethod
    def _apply_label_in_fresh_word(output: Path):
        apply_label_in_fresh_word(output)

    @staticmethod
    def _merge_with_word(fragments: Sequence[Path], output: Path):
        """Merge multiple excerpts by editing a copy of the first fragment.

        The working document lives under the application temp root (the fragment
        job directory). This avoids Documents.Add()/SaveAs2() and keeps Word away
        from OneDrive/network/user-selected destinations until finalization is done.
        """
        if not fragments:
            raise ValueError("No document fragments were generated.")
        if len(fragments) < 2:
            raise ValueError("Word merge is only required for multiple fragments.")

        working = fragments[0].parent / "merged.docx"
        if working.exists():
            working.unlink()
        shutil.copy2(fragments[0], working)

        word = None
        dest = None
        com_initialized = False
        com_server_dead = False
        apply_label = _resolve_facade_attr("set_sensitivity_label", set_sensitivity_label)
        ensure_styles = _resolve_facade_attr(
            "ensure_3gpp_package_styles", ensure_3gpp_package_styles
        )
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            with temporary_virtual_default_printer(log=logger):
                word = win32com.client.DispatchEx("Word.Application")
                WordExcerptEngine._configure_word(word)

                logger.info(
                    "   -> Opening first excerpt fragment as merge base: %s",
                    fragments[0].name,
                )
                dest = word.Documents.Open(
                    os.path.normpath(str(working)),
                    ReadOnly=False,
                    AddToRecentFiles=False,
                )

                # Unlock and pre-apply corporate label on merge base before InsertFile()
                try:
                    if dest.ProtectionType != -1:
                        dest.Unprotect()
                except Exception:
                    pass

                try:
                    apply_label(dest)
                except Exception as label_exc:
                    logger.warning(
                        "   -> Pre-labeling merge base encountered an issue (will retry at finalization): %s",
                        label_exc,
                    )

                for index, fragment in enumerate(fragments[1:], start=2):
                    logger.info(
                        "   -> Appending excerpt fragment %d/%d: %s",
                        index,
                        len(fragments),
                        fragment.name,
                    )
                    rng = dest.Range(dest.Content.End - 1, dest.Content.End - 1)
                    rng.InsertFile(FileName=os.path.normpath(str(fragment)))

                dest.Save()
                dest.Close(SaveChanges=False)
                dest = None
                word.Quit(SaveChanges=0)
                word = None

            if not working.exists() or working.stat().st_size == 0:
                raise RuntimeError("Microsoft Word returned without producing the merged excerpt.")

            WordExcerptEngine._copy_formatting_parts(fragments[0], working)
            WordExcerptEngine._apply_label_in_fresh_word(working)
            ensure_styles(working)

            ext_output = to_extended_path(output)
            os.makedirs(to_extended_path(output.parent), exist_ok=True)
            if os.path.exists(ext_output):
                try:
                    os.remove(ext_output)
                except OSError:
                    pass
            shutil.copy2(to_extended_path(working), ext_output)
            if not os.path.exists(ext_output) or os.path.getsize(ext_output) == 0:
                raise RuntimeError(
                    "Failed to copy the finalized merged excerpt to its destination."
                )
            logger.info("   -> Finalized merged excerpt copied to: %s", output)

        except Exception as exc:
            com_server_dead = WordExcerptEngine._is_rpc_disconnect(exc)
            raise
        finally:
            if not com_server_dead:
                if dest is not None:
                    try:
                        dest.Close(SaveChanges=False)
                    except Exception as exc:
                        if WordExcerptEngine._is_rpc_disconnect(exc):
                            com_server_dead = True
                if word is not None and not com_server_dead:
                    try:
                        word.Quit(SaveChanges=0)
                    except Exception:
                        pass
            dest = None
            word = None
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
