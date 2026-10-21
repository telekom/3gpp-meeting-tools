# --- File: src/modules/specifications/ui/threads.py ---
import logging
import os
import re
import zipfile
from pathlib import Path
from typing import List, Sequence, Union

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path
from core.network.session import (
    DownloadCancelledError,
    HttpError,
    NetworkError,
    NetworkSession,
)
from modules.spec_search.core.spec_search_threads import is_change_mark_file
from modules.word_tools.core.word_converter import convert_doc_to_docx
from modules.word_tools.core.word_excerpt_extractor import (
    HeadingInfo,
    WordHeadingScanner,
    order_sources,
)

logger = logging.getLogger(__name__)


class SpecDownloadThread(QThread):
    """
    Background worker thread to download specification archives from 3GPP FTP
    using the centralized NetworkSession.
    """

    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)
    ui_log_msg = pyqtSignal(str, int)
    progress = pyqtSignal(int)

    def __init__(self, url: str, dest_path: Union[str, Path], timeout: int = 30):
        super().__init__()
        self.url = str(url).strip()
        self.dest_path = Path(dest_path)
        self.timeout = timeout

    def _log(self, message: str, level: int = logging.INFO):
        logger.log(level, message)

    def run(self):
        if not self.url:
            self.error.emit("Download URL is empty.")
            return

        try:
            self._log(
                f"⬇️ Downloading: {self.dest_path.name}...", logging.INFO
            )

            # Delegate connection, delay, headers, streaming, and progress to NetworkSession
            downloaded_bytes = NetworkSession.download_file(
                url=self.url,
                dest_path=self.dest_path,
                timeout=self.timeout,
                progress_cb=self.progress.emit,
                cancel_cb=self.isInterruptionRequested,
                atomic=True,
            )

            size_kb = downloaded_bytes / 1024.0
            self._log(
                f"✅ Successfully downloaded {self.dest_path.name} ({size_kb:.1f} KB)",
                logging.INFO,
            )
            self.finished_success.emit(self.dest_path)

        except DownloadCancelledError:
            self._log(
                f"⚠️ Download cancelled: {self.dest_path.name}", logging.WARNING
            )

        except HttpError as e:
            err_msg = f"HTTP {e.status_code} downloading {self.dest_path.name}: {e}"
            logging.error(err_msg)
            self._log(f"❌ {err_msg}", logging.ERROR)
            self.error.emit(str(e))

        except NetworkError as e:
            err_msg = f"Network error downloading {self.dest_path.name}: {e}"
            logging.error(err_msg)
            self._log(f"❌ {err_msg}", logging.ERROR)
            self.error.emit(str(e))


class SpecExcerptPrepThread(QThread):
    """
    Background worker that handles the complete preparation pipeline for clause extraction:
      1. Downloads specification archive from 3GPP FTP if missing locally.
      2. Extracts clean Word documents (.docx, .doc), filtering out revision marks (-rm)
         and non-document files (.yaml, .asn, .pdf).
      3. Converts legacy .doc files to .docx if necessary.
      4. Orders multi-part split documents in natural clause order.
      5. Scans heading blocks into a HeadingInfo hierarchy.
    """

    progress = pyqtSignal(int, str)
    prepared = pyqtSignal(list, list)  # ordered_paths: List[str], headings: List[HeadingInfo]
    error = pyqtSignal(str)

    def __init__(self, spec_num: str, version: str, fname: str, url: str, spec_dl_dir: Path):
        super().__init__()
        self.spec_num = spec_num.strip()
        self.version = version.strip()
        self.fname = fname.strip()
        self.url = url.strip()
        self.spec_dl_dir = Path(spec_dl_dir)

    def run(self):
        try:
            os.makedirs(to_extended_path(self.spec_dl_dir), exist_ok=True)
            stem = Path(self.fname).stem.lower() if self.fname else ""
            zip_path = self.spec_dl_dir / self.fname if self.fname else None

            # --- STAGE 1: Check Local Word Documents or Download Archive ---
            self.progress.emit(5, "Checking local specification files...")

            existing_local_docs = self._find_local_version_docs(stem)

            if not existing_local_docs and zip_path and not zip_path.exists():
                if not self.url:
                    self.error.emit(f"No local files and no download URL available for {self.spec_num} v{self.version}.")
                    return

                self.progress.emit(10, f"Downloading {self.fname} from 3GPP FTP...")

                def on_dl_progress(pct: int):
                    scaled = 10 + int(pct * 0.35)
                    self.progress.emit(scaled, f"Downloading {self.fname}... ({pct}%)")

                try:
                    NetworkSession.download_file(
                        url=self.url,
                        dest_path=zip_path,
                        progress_cb=on_dl_progress,
                        cancel_cb=self.isInterruptionRequested,
                        atomic=True,
                    )
                except Exception as dl_err:
                    self.error.emit(f"Failed to download {self.fname}: {dl_err}")
                    return

            if self.isInterruptionRequested():
                return

            # --- STAGE 2: Extract & Filter from ZIP Archive ---
            extracted_docs: List[Path] = []
            if zip_path and os.path.exists(to_extended_path(zip_path)):
                self.progress.emit(45, f"Inspecting archive: {zip_path.name}...")
                try:
                    with zipfile.ZipFile(to_extended_path(zip_path), "r") as zf:
                        all_members = [
                            m for m in zf.namelist()
                            if not m.startswith("._") and "__MACOSX" not in m
                        ]

                        # Filter for Word candidates
                        word_members = [
                            m for m in all_members
                            if m.lower().endswith((".docx", ".doc"))
                        ]

                        # Change-mark filtering: Prefer clean documents over -rm versions
                        clean_word_members = [m for m in word_members if not is_change_mark_file(m)]
                        candidates = clean_word_members if clean_word_members else word_members

                        for member in candidates:
                            target_file = self.spec_dl_dir / Path(member).name
                            ext_target_file = to_extended_path(target_file)
                            if not os.path.exists(ext_target_file):
                                with open(ext_target_file, "wb") as f:
                                    f.write(zf.read(member))
                            if target_file not in extracted_docs:
                                extracted_docs.append(target_file)
                except Exception as zip_err:
                    logger.warning(f"Error extracting archive {zip_path.name}: {zip_err}")

            # Merge with any already extracted local files
            candidate_docs = list(set(extracted_docs + existing_local_docs))
            clean_candidates = [p for p in candidate_docs if not is_change_mark_file(p.name)]
            selected_docs = clean_candidates if clean_candidates else candidate_docs

            if not selected_docs:
                self.error.emit(
                    f"No Word (.docx / .doc) documents found for {self.spec_num} v{self.version}."
                )
                return

            if self.isInterruptionRequested():
                return

            # --- STAGE 3: Convert Legacy .doc to .docx ---
            self.progress.emit(55, "Resolving document formats...")
            docx_files: List[Path] = []

            # Deduplicate stems preferring .docx over .doc
            stem_map = {}
            for doc_path in selected_docs:
                doc_stem = doc_path.stem.lower()
                if doc_path.suffix.lower() == ".docx":
                    stem_map[doc_stem] = doc_path
                elif doc_stem not in stem_map:
                    stem_map[doc_stem] = doc_path

            total_items = len(stem_map)
            for idx, (doc_stem, doc_path) in enumerate(stem_map.items()):
                if self.isInterruptionRequested():
                    return

                if doc_path.suffix.lower() == ".doc":
                    target_docx = doc_path.with_suffix(".docx")
                    if target_docx.exists() and target_docx.stat().st_size > 0:
                        docx_files.append(target_docx)
                    else:
                        pct = 55 + int((idx / max(1, total_items)) * 25)
                        self.progress.emit(pct, f"Converting legacy .doc: {doc_path.name}...")
                        try:
                            converted = convert_doc_to_docx(doc_path)
                            docx_files.append(converted)
                        except Exception as conv_err:
                            logger.error(f"Failed to convert {doc_path.name}: {conv_err}")
                else:
                    docx_files.append(doc_path)

            if not docx_files:
                self.error.emit(f"Could not obtain .docx documents for {self.spec_num} v{self.version}.")
                return

            # --- STAGE 4: Multi-Part Natural Ordering ---
            ordered_paths = order_sources([str(p) for p in docx_files], mode="filename")

            if self.isInterruptionRequested():
                return

            # --- STAGE 5: Scan Heading Blocks ---
            self.progress.emit(85, f"Scanning headings in {len(ordered_paths)} document part(s)...")

            def report_scan(file_index, total_files, block_index, total_blocks, path):
                file_fraction = block_index / max(1, total_blocks)
                scan_pct = 85 + int(((file_index + file_fraction) / max(1, total_files)) * 14)
                scan_pct = max(85, min(99, scan_pct))
                self.progress.emit(
                    scan_pct,
                    f"Scanning {Path(path).name} ({file_index + 1}/{total_files})...",
                )

            scanner = WordHeadingScanner()
            headings: List[HeadingInfo] = scanner.scan(ordered_paths, order_mode="drop", progress_callback=report_scan)

            self.progress.emit(100, f"Ready: {len(headings)} heading(s) detected across {len(ordered_paths)} file(s).")
            self.prepared.emit(ordered_paths, headings)

        except Exception as e:
            logger.exception("Error during specification excerpt preparation")
            self.error.emit(f"Preparation failed: {e}")

    def _find_local_version_docs(self, stem: str) -> List[Path]:
        """Locates cached Word files belonging to the version stem."""
        if not self.spec_dl_dir.exists() or not stem:
            return []

        clean_stem = stem.replace("_cl", "").replace("-cl", "")
        found = []
        try:
            for item in self.spec_dl_dir.iterdir():
                if not item.is_file():
                    continue
                name_lower = item.name.lower()
                if not (name_lower.endswith(".docx") or name_lower.endswith(".doc")):
                    continue
                if name_lower.startswith(stem) or name_lower.startswith(clean_stem):
                    found.append(item)
        except Exception:
            pass
        return found


class SpecMarkdownConverterThread(QThread):
    """
    Background worker thread to convert single or multi-part specifications
    into single unified Markdown (.md) or clause-split dossiers with version-isolated media.
    """
    finished_success = pyqtSignal(Path, str)
    error = pyqtSignal(str)
    progress_status = pyqtSignal(str)

    def __init__(
        self,
        spec_dir: Path,
        stem: str,
        mode: str = "single",
        meta: dict = None,
    ):
        super().__init__()
        self.spec_dir = Path(spec_dir)
        self.stem = stem
        self.mode = mode
        self.meta = meta or {}

    def run(self):
        try:
            from modules.specifications.core.spec_markdown_converter import convert_spec_to_markdown

            def on_progress(msg: str):
                self.progress_status.emit(msg)

            result_path = convert_spec_to_markdown(
                spec_dir=self.spec_dir,
                stem=self.stem,
                mode=self.mode,
                meta=self.meta,
                progress_callback=on_progress,
                cancellation_check=self.isInterruptionRequested,
            )
            mode_desc = "Clause Dossier" if self.mode == "split" else "Unified Markdown Document"
            self.finished_success.emit(result_path, f"Successfully generated {mode_desc} at:\n{result_path}")
        except Exception as e:
            logger.exception("Error during specification markdown conversion")
            self.error.emit(str(e))