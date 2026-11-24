"""CR/pCR assembly engine, Word COM fragment splicing, and PyQt5 worker thread."""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_temp_root, to_extended_path
from modules.word_tools.core.cr_pcr.cr_builder import build_cr_base_document
from modules.word_tools.core.cr_pcr.models import (
    ContributionMetadata,
    extract_clause_summary,
    group_selected_headings_into_blocks,
)
from modules.word_tools.core.cr_pcr.pcr_builder import build_pcr_base_document
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import (
    ensure_3gpp_package_styles,
    promote_3gpp_styles_in_word,
)
from modules.word_tools.core.word_excerpt_extractor import (
    HeadingInfo,
    WordExcerptEngine,
)

logger = logging.getLogger(__name__)


class CrPcrGeneratorEngine:
    """
    Orchestrates end-to-end CR and pCR generation:
    1. Slices contiguous specification fragments via worker subprocess.
    2. Builds base template with 3GPP styles and change anchors.
    3. Splices fragments into anchors via Word COM Range.InsertFile.
    4. Enforces corporate sensitivity labeling and 3GPP styles lock-in.
    """

    def __init__(
        self,
        metadata: ContributionMetadata,
        source_paths: Sequence[str],
        headings: Sequence[HeadingInfo],
        selected_ids: Sequence[str],
        all_new_map: Optional[Dict[str, bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None,
    ):
        self.meta = metadata
        self.source_paths = list(source_paths)
        self.headings = list(headings)
        self.selected_ids = list(selected_ids)
        self.all_new_map = all_new_map or {}
        self.progress_callback = progress_callback

    def _progress(self, percent: int, msg: str):
        if self.progress_callback:
            self.progress_callback(max(0, min(100, int(percent))), msg)

    def _make_fragment(self, source: Path, ranges: List[Tuple[int, int]], target: Path) -> None:
        """Invokes isolated worker to slice the source specification."""
        helper = Path(__file__).resolve().parent.parent / "word_excerpt_fragment_worker.py"
        if not helper.is_file():
            raise RuntimeError(f"Fragment worker not found at {helper}")

        merged = WordExcerptEngine._merge_ranges(ranges, 2**31 - 1)
        cmd = [
            sys.executable,
            str(helper),
            str(source),
            str(target),
            json.dumps(merged),
        ]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        res = subprocess.run(cmd, capture_output=True, text=True, creationflags=flags, check=False)
        if res.returncode != 0:
            err = res.stderr.strip() or res.stdout.strip() or "Unknown worker error"
            raise RuntimeError(f"Fragment extraction failed for {source.name}: {err}")

    def generate(self, output_path: str) -> Path:
        """Executes full generation pipeline."""
        output = Path(output_path).resolve()
        os.makedirs(to_extended_path(output.parent), exist_ok=True)
        work_dir = Path(tempfile.mkdtemp(prefix="cr_gen_", dir=str(get_temp_root())))

        try:
            # 1. Group selected clauses into contiguous change blocks
            self._progress(5, "Analyzing clause contiguity and change blocks...")
            change_blocks = group_selected_headings_into_blocks(
                self.headings, self.selected_ids, self.all_new_map
            )
            if not change_blocks:
                raise ValueError("No valid clauses selected for contribution generation.")

            # Auto-compute affected clauses if empty
            if not self.meta.clauses_affected:
                self.meta.clauses_affected = extract_clause_summary(change_blocks)

            logger.info(
                "Grouped %d selected heading(s) into %d change block(s): %s",
                len(self.selected_ids),
                len(change_blocks),
                self.meta.clauses_affected,
            )

            # 2. Extract specification fragments
            fragment_paths = []
            total_blocks = len(change_blocks)
            for idx, block in enumerate(change_blocks):
                pct = 10 + int(40 * idx / total_blocks)
                self._progress(pct, f"Extracting change section {idx + 1}/{total_blocks}...")
                frag = work_dir / f"change_{idx:03d}.docx"
                self._make_fragment(Path(block.source_path), block.ranges, frag)
                fragment_paths.append(frag)

            # 3. Build Base Document (CR or pCR)
            self._progress(55, f"Constructing base {self.meta.mode.upper()} template...")
            base_doc_path = work_dir / f"base_{self.meta.mode}.docx"
            if self.meta.mode.lower() == "cr":
                build_cr_base_document(self.meta, change_blocks, base_doc_path)
            else:
                build_pcr_base_document(self.meta, change_blocks, base_doc_path)

            # 4. Splice fragments into anchors via Word COM
            self._progress(65, "Splicing change blocks with Microsoft Word COM...")
            merged_working_path = work_dir / "finalized_working.docx"
            shutil.copy2(to_extended_path(base_doc_path), to_extended_path(merged_working_path))

            self._splice_with_word(merged_working_path, fragment_paths)

            # 5. Lock in 3GPP styles, table formatting, and sensitivity label
            self._progress(85, "Applying 3GPP styles and corporate sensitivity label...")
            ensure_3gpp_package_styles(merged_working_path)

            # Apply sensitivity label in a fresh Word session
            pythoncom.CoInitialize()
            try:
                WordExcerptEngine._apply_label_in_fresh_word(merged_working_path)
            finally:
                pythoncom.CoUninitialize()

            # Final style assertion: ensure styles.xml retains all 3GPP QuickStyles and settings.xml has linkStyles stripped
            ensure_3gpp_package_styles(merged_working_path)

            # 6. Copy finalized file to output path
            ext_output = to_extended_path(output)
            if os.path.exists(ext_output):
                try:
                    os.remove(ext_output)
                except OSError:
                    pass
            shutil.copy2(to_extended_path(merged_working_path), ext_output)

            self._progress(100, f"Successfully created {output.name}!")
            return output

        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _splice_with_word(self, base_doc_path: Path, fragment_paths: List[Path]) -> None:
        """Opens base document, finds each anchor, and inserts the corresponding fragment."""
        word = None
        doc = None
        com_initialized = False
        com_dead = False

        try:
            from core.utils.office_printer_guard import temporary_virtual_default_printer
            pythoncom.CoInitialize()
            com_initialized = True
            with temporary_virtual_default_printer(log=logger):
                word = win32com.client.DispatchEx("Word.Application")
                WordExcerptEngine._configure_word(word)

                doc = word.Documents.Open(
                    os.path.normpath(str(base_doc_path)),
                    ReadOnly=False,
                    AddToRecentFiles=False,
                )

                # Unprotect and pre-label to prevent enterprise editing lockouts
                try:
                    if doc.ProtectionType != -1:
                        doc.Unprotect()
                except Exception:
                    pass

                try:
                    set_sensitivity_label(doc)
                except Exception as label_exc:
                    logger.warning("Pre-labeling in COM encountered an issue: %s", label_exc)

                # Insert each fragment at its corresponding anchor
                for idx, frag_path in enumerate(fragment_paths):
                    anchor_text = f"{{{{CHANGE_BLOCK_ANCHOR_{idx}}}}}"
                    rng = doc.Content
                    find = rng.Find
                    find.ClearFormatting()
                    find.Text = anchor_text
                    found = find.Execute()
                    if found:
                        rng.Text = ""  # Clear placeholder anchor text
                        rng.InsertFile(FileName=os.path.normpath(str(frag_path)))
                    else:
                        logger.warning("Could not locate anchor %s in base document.", anchor_text)

                # Update fields so DOCPROPERTY numbers render
                try:
                    doc.Fields.Update()
                except Exception:
                    pass

                # Promote all 3GPP styles in live Word model
                promote_3gpp_styles_in_word(doc)

                doc.Save()
                doc.Close(SaveChanges=False)
                doc = None
                word.Quit(SaveChanges=0)
                word = None

        except Exception as exc:
            com_dead = WordExcerptEngine._is_rpc_disconnect(exc)
            raise RuntimeError(f"Word COM assembly failed: {exc}") from exc
        finally:
            if not com_dead:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if word is not None:
                    try:
                        word.Quit(SaveChanges=0)
                    except Exception:
                        pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class CrPcrGeneratorThread(QThread):
    """Asynchronous background worker thread for generating CR/pCR contributions."""
    finished_path = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        metadata: ContributionMetadata,
        source_paths: Sequence[str],
        headings: Sequence[HeadingInfo],
        selected_ids: Sequence[str],
        output_path: str,
        all_new_map: Optional[Dict[str, bool]] = None,
    ):
        super().__init__()
        self.metadata = metadata
        self.source_paths = source_paths
        self.headings = headings
        self.selected_ids = selected_ids
        self.output_path = output_path
        self.all_new_map = all_new_map or {}

    def run(self):
        try:
            logger.info("Starting contribution generation (%s)...", self.metadata.mode.upper())
            engine = CrPcrGeneratorEngine(
                metadata=self.metadata,
                source_paths=self.source_paths,
                headings=self.headings,
                selected_ids=self.selected_ids,
                all_new_map=self.all_new_map,
                progress_callback=lambda pct, msg: self.progress.emit(pct, msg),
            )
            out_file = engine.generate(self.output_path)
            self.finished_path.emit(str(out_file))
        except Exception as exc:
            logger.exception("Contribution generation failed: %s", exc)
            self.failed.emit(str(exc))
