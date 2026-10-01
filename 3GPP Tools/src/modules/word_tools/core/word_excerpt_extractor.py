import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import pythoncom
import win32com.client
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from PyQt5.QtCore import QThread, pyqtSignal

from modules.word_tools.core.sensitivity_label import set_sensitivity_label

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class HeadingInfo:
    id: str
    source_path: str
    source_index: int
    block_index: int
    end_index: int
    level: int
    text: str


def natural_key(path: str):
    name = Path(path).name.lower()
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", name)]


def order_sources(paths: Sequence[str], mode: str = "filename") -> List[str]:
    clean = [str(Path(p)) for p in paths if p]
    return sorted(clean, key=natural_key) if mode == "filename" else clean


class WordHeadingScanner:
    @staticmethod
    def _blocks(doc):
        return [x for x in doc.element.body if isinstance(x, (CT_P, CT_Tbl))]

    @staticmethod
    def _heading_level(para: Paragraph) -> int:
        style_name = ""
        try:
            style_name = para.style.name or ""
        except Exception:
            pass
        match = re.match(r"^Heading\s+(\d+)$", style_name, re.I)
        if match:
            return int(match.group(1))
        try:
            outline = para._p.pPr.outlineLvl
            if outline is not None:
                return int(outline.val) + 1
        except Exception:
            pass
        return 0

    def scan(self, paths: Sequence[str], order_mode: str = "filename", progress_callback=None) -> List[HeadingInfo]:
        ordered = order_sources(paths, order_mode)
        result: List[HeadingInfo] = []
        total_files = len(ordered)
        for source_index, path in enumerate(ordered):
            if progress_callback:
                progress_callback(source_index, total_files, 0, 1, path)
            doc = Document(path)
            blocks = self._blocks(doc)
            total_blocks = max(1, len(blocks))
            raw = []
            for block_index, block in enumerate(blocks):
                if progress_callback and (block_index % 100 == 0 or block_index == total_blocks - 1):
                    progress_callback(source_index, total_files, block_index + 1, total_blocks, path)
                if not isinstance(block, CT_P):
                    continue
                para = Paragraph(block, doc)
                text = para.text.strip()
                if not text:
                    continue
                level = self._heading_level(para)
                if level:
                    raw.append((block_index, level, text))
            for i, (block_index, level, text) in enumerate(raw):
                end_index = len(blocks)
                for next_block, next_level, _ in raw[i + 1:]:
                    if next_level <= level:
                        end_index = next_block
                        break
                result.append(HeadingInfo(
                    id=f"{source_index}:{block_index}", source_path=path,
                    source_index=source_index, block_index=block_index,
                    end_index=end_index, level=level, text=text,
                ))
        return result


class WordExcerptEngine:
    def __init__(self, source_paths: Sequence[str], headings: Sequence[HeadingInfo], progress_callback=None):
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
        specifications can spend minutes serializing a DOCX package.  A QThread
        still shares the application's Python interpreter/GIL, whereas this helper
        process cannot starve the Qt event loop.  The helper uses the same
        formatting-safe python-docx body reconstruction as the in-process version.
        """
        helper = Path(__file__).with_name("word_excerpt_fragment_worker.py")
        if not helper.is_file():
            raise RuntimeError(f"Excerpt fragment worker not found: {helper}")

        merged = self._merge_ranges(ranges, 2**31 - 1)
        logger.info(
            "   -> Starting isolated excerpt preparation for %s (%d selected range(s))",
            source.name, len(merged),
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
        output.parent.mkdir(parents=True, exist_ok=True)
        work = Path(tempfile.mkdtemp(prefix="3gpp_word_excerpt_"))
        fragments = []
        try:
            total = max(1, len(source_order))
            for idx, source_path in enumerate(source_order):
                self._progress(5 + int(45 * idx / total), f"Preparing excerpt from {Path(source_path).name}...")
                fragment = work / f"fragment_{idx:03d}.docx"
                self._make_fragment(Path(source_path), grouped[source_path], fragment)
                fragments.append(fragment)
            self._progress(55, "Finalizing excerpt with Microsoft Word...")
            self._merge_with_word(fragments, output)
            self._progress(100, "Extraction complete.")
            return output
        finally:
            shutil.rmtree(work, ignore_errors=True)

    @staticmethod
    def _copy_formatting_parts(source: Path, target: Path, attempts: int = 12, delay: float = 0.5):
        """Restore source formatting parts after Word has fully released the output.

        Office and OneDrive can retain a file handle briefly after Document.Close()/
        Application.Quit().  Retry only sharing/permission failures; other ZIP/package
        errors are real failures and should surface immediately.
        """
        parts = (
            "word/styles.xml",
            "word/stylesWithEffects.xml",
            "word/theme/theme1.xml",
            "word/fontTable.xml",
        )
        tmp = target.with_name(target.name + ".formatting.tmp")
        last_exc = None
        for attempt in range(1, attempts + 1):
            try:
                if tmp.exists():
                    tmp.unlink()
                with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(target, "r") as dst, \
                        zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as out:
                    source_names = set(src.namelist())
                    replaced = set()
                    for item in dst.infolist():
                        name = item.filename
                        if name in parts and name in source_names:
                            out.writestr(item, src.read(name))
                            replaced.add(name)
                        else:
                            out.writestr(item, dst.read(name))
                    for name in parts:
                        if name in source_names and name not in replaced:
                            out.writestr(name, src.read(name))
                            replaced.add(name)
                os.replace(tmp, target)
                logger.info(
                    "   -> Restored source formatting package parts: %s",
                    ", ".join(sorted(replaced)),
                )
                return
            except PermissionError as exc:
                last_exc = exc
                try:
                    if tmp.exists():
                        tmp.unlink()
                except OSError:
                    pass
                if attempt == attempts:
                    break
                logger.info(
                    "   -> Output is still locked; retrying formatting restore (%d/%d)...",
                    attempt, attempts,
                )
                time.sleep(delay)
        raise RuntimeError(
            f"Microsoft Word/OneDrive did not release the output file in time: {target}"
        ) from last_exc

    @staticmethod
    def _configure_word(word):
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
            word.Options.ConfirmConversions = False
            word.Options.DoNotPromptForConvert = True
            word.Options.WarnBeforeSavingPrintingSendingMarkup = False
            word.Options.SaveInterval = 0
        except Exception:
            pass

    @staticmethod
    def _is_rpc_disconnect(exc: BaseException) -> bool:
        hresult = getattr(exc, "hresult", None)
        if hresult in (-2147023170, -2147023174, -2147417848, -2147418111):
            return True
        msg = str(exc).lower()
        return any(marker in msg for marker in (
            "0x800706be", "0x800706ba", "0x80010108", "0x80010001",
            "remote procedure call failed", "rpc server is unavailable",
            "object invoked has disconnected", "call was rejected by callee",
        ))

    @staticmethod
    def _apply_label_in_fresh_word(output: Path):
        """Apply the corporate label only after package patching is complete.

        A dedicated Word instance isolates sensitivity-label add-in/policy failures
        from the merge session.  Label failure is deliberately fatal: returning an
        unlabeled corporate document would be misleading.
        """
        word = None
        doc = None
        server_dead = False
        try:
            word = win32com.client.DispatchEx("Word.Application")
            WordExcerptEngine._configure_word(word)
            logger.info("   -> Reopening finalized excerpt for sensitivity labeling...")
            doc = word.Documents.Open(
                os.path.normpath(str(output)),
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            set_sensitivity_label(doc)
            doc.Save()
            doc.Close(SaveChanges=False)
            doc = None
            word.Quit(SaveChanges=0)
            word = None
            logger.info("   -> Sensitivity label applied successfully.")
        except Exception as exc:
            server_dead = WordExcerptEngine._is_rpc_disconnect(exc)
            raise RuntimeError(
                "The excerpt was created and formatted, but Microsoft Word failed "
                "while applying the required corporate sensitivity label. The output "
                "was not reported as successfully completed."
            ) from exc
        finally:
            if not server_dead:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception as exc:
                        if WordExcerptEngine._is_rpc_disconnect(exc):
                            server_dead = True
                if word is not None and not server_dead:
                    try:
                        word.Quit(SaveChanges=0)
                    except Exception:
                        pass
            doc = None
            word = None

    @staticmethod
    def _merge_with_word(fragments: Sequence[Path], output: Path):
        if not fragments:
            raise ValueError("No document fragments were generated.")

        word = None
        dest = None
        com_initialized = False
        com_server_dead = False

        try:
            pythoncom.CoInitialize()
            com_initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            WordExcerptEngine._configure_word(word)

            # Merge in a fresh writable document. Do NOT invoke sensitivity labeling
            # in this session: managed-label add-ins/policy can disconnect Word COM.
            dest = word.Documents.Add()
            for index, fragment in enumerate(fragments, start=1):
                logger.info(
                    "   -> Merging excerpt fragment %d/%d: %s",
                    index, len(fragments), fragment.name,
                )
                rng = dest.Range(dest.Content.End - 1, dest.Content.End - 1)
                rng.InsertFile(FileName=os.path.normpath(str(fragment)))

            if output.exists():
                output.unlink()
            dest.SaveAs2(
                FileName=os.path.normpath(str(output)),
                FileFormat=12,  # wdFormatXMLDocument
            )
            dest.Close(SaveChanges=False)
            dest = None
            word.Quit(SaveChanges=0)
            word = None

            if not output.exists() or output.stat().st_size == 0:
                raise RuntimeError("Microsoft Word returned without producing the excerpt file.")

            # Word must be completely out of the way before direct package access.
            WordExcerptEngine._copy_formatting_parts(fragments[0], output)

            # Only now invoke managed sensitivity-label handling, in an isolated Word
            # instance so an add-in/RPC failure cannot strand the merge output locked.
            WordExcerptEngine._apply_label_in_fresh_word(output)

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



class WordExcerptThread(QThread):
    finished_path = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)
    finished = pyqtSignal()

    def __init__(self, source_paths, headings, selected_ids, output_path):
        super().__init__()
        self.source_paths = source_paths
        self.headings = headings
        self.selected_ids = selected_ids
        self.output_path = output_path

    def run(self):
        try:
            logger.info("Extracting %d selected heading(s)...", len(self.selected_ids))
            out = WordExcerptEngine(
                self.source_paths, self.headings,
                progress_callback=lambda percent, message: self.progress.emit(percent, message),
            ).extract(self.selected_ids, self.output_path)
            logger.info("Excerpt created: %s", out)
            self.finished_path.emit(str(out))
        except Exception as exc:
            logger.exception("Word excerpt extraction failed")
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()

class WordHeadingScanThread(QThread):
    scanned = pyqtSignal(object, object)  # ordered paths, headings
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, source_paths, order_mode="filename"):
        super().__init__()
        self.source_paths = list(source_paths)
        self.order_mode = order_mode

    def run(self):
        try:
            ordered = order_sources(self.source_paths, self.order_mode)
            def report(file_index, total_files, block_index, total_blocks, path):
                file_fraction = block_index / max(1, total_blocks)
                percent = int(((file_index + file_fraction) / max(1, total_files)) * 100)
                percent = max(0, min(99, percent))
                self.progress.emit(percent, f"Parsing {Path(path).name} - document {file_index + 1} of {total_files}")

            headings = WordHeadingScanner().scan(ordered, "drop", report)
            self.progress.emit(100, f"Parsed {len(ordered)} document(s)")
            self.scanned.emit(ordered, headings)
        except Exception as exc:
            logger.exception("Heading scan failed")
            self.failed.emit(str(exc))
