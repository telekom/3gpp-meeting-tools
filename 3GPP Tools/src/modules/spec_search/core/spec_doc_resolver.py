"""
Core resolver and dispatcher for opening 3GPP specification documents directly
from Spec Search, including automatic local cache discovery, ZIP extraction,
multi-part clause range targeting, and clipboard search assistance.
"""

import logging
import os
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressDialog, QWidget

from core.network.session import NetworkSession
from core.utils.paths import get_specs_root, to_extended_path
from modules.spec_search.core.spec_search_threads import is_change_mark_file
from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)


def version_to_letter_code(version: str) -> str:
    """
    Converts a standard 3GPP version (e.g. '16.20.0') to its 3-character filename code ('gk0').
    """
    parts = str(version).lstrip("vV").strip().split(".")
    if len(parts) != 3:
        return ""

    def num_to_char(n: int) -> str:
        if n < 10:
            return str(n)
        return chr(ord("a") + n - 10)

    try:
        return f"{num_to_char(int(parts[0]))}{num_to_char(int(parts[1]))}{num_to_char(int(parts[2]))}"
    except Exception:
        return ""


def letter_code_to_version(code: str) -> str:
    """
    Converts a 3-character filename code (e.g. 'gk0' -> '16.20.0', 'k00' -> '20.0.0') back to version.
    """
    clean = str(code or "").strip()
    if len(clean) != 3:
        return ""

    def char_to_num(c: str) -> int:
        if c.isdigit():
            return int(c)
        return ord(c.lower()) - ord("a") + 10

    try:
        return f"{char_to_num(clean[0])}.{char_to_num(clean[1])}.{char_to_num(clean[2])}"
    except Exception:
        return ""


def parse_clause_token(t: str) -> Optional[Tuple[int, ...]]:
    """Parses a multi-part filename clause token like 's00', 's04', 's0504', 's0506' into a numeric tuple."""
    clean = t.lstrip("sS_").strip()
    if clean.isdigit():
        if len(clean) == 2:
            return (int(clean),)
        if len(clean) == 4:
            return (int(clean[:2]), int(clean[2:]))
        if len(clean) == 6:
            return (int(clean[:2]), int(clean[2:4]), int(clean[4:]))
        return (int(clean),)
    parts = [int(p) for p in re.findall(r"\d+", clean)]
    return tuple(parts) if parts else None


def match_multipart_document_for_clause(docs: List[Path], clause_number: Optional[str]) -> Optional[Path]:
    """
    If multiple split documents exist (e.g., '_s00_s04.docx', '_s05_s08.docx'),
    picks the document corresponding to the specified clause or Annex.
    """
    if not docs:
        return None
    if len(docs) == 1 or not clause_number:
        return docs[0]

    c_str = str(clause_number).strip()

    # Check annexes (e.g. "Annex A", "A.1", "D.6")
    annex_m = re.match(r"^(?:annex\s+)?([A-Za-z])", c_str, re.IGNORECASE)
    if annex_m and annex_m.group(1).upper() not in ("C", "V"):  # avoid false match on Cl. / v.
        letter = annex_m.group(1).upper()
        for doc in docs:
            stem_l = doc.stem.lower()
            if "annex" in stem_l or f"_s{letter.lower()}" in stem_l or f"s{letter.lower()}_" in stem_l:
                return doc

    nums = [int(p) for p in re.findall(r"\d+", c_str)]
    if not nums:
        return docs[0]
    target_tup = tuple(nums[:4]) + (0,) * max(0, 4 - len(nums))

    # Match clause ranges, e.g. _s00_s04, _s05_s0504, _s06_s08, _s09_s10
    for doc in docs:
        m = re.search(r"_(?:cover_)?s([0-9A-Za-z]+)_s([0-9A-Za-z]+)", doc.stem, re.IGNORECASE)
        if m:
            st = parse_clause_token(m.group(1))
            et = parse_clause_token(m.group(2))
            if st and et:
                st_pad = st + (0,) * (4 - len(st))
                et_pad = et + (999999,) * (4 - len(et))
                if st_pad <= target_tup <= et_pad:
                    return doc

    # Match single split files, e.g. _s06_17
    for doc in docs:
        m = re.search(r"_s([0-9A-Za-z]+)(?:_([0-9A-Za-z]+))?", doc.stem, re.IGNORECASE)
        if m:
            t1 = parse_clause_token(m.group(1))
            t2 = parse_clause_token(m.group(2)) if m.group(2) else None
            t_full = (t1 + t2) if (t1 and t2) else t1
            if t_full:
                if target_tup[:len(t_full)] == t_full:
                    return doc

    return docs[0]


class SpecDownloadWorker(QThread):
    """Background worker for on-demand specification ZIP download."""
    progress = pyqtSignal(int)
    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, url: str, target_zip_path: Path):
        super().__init__()
        self.url = url
        self.target_zip_path = target_zip_path

    def run(self):
        try:
            self.target_zip_path.parent.mkdir(parents=True, exist_ok=True)
            NetworkSession.download_file(self.url, self.target_zip_path)
            self.finished_success.emit(self.target_zip_path)
        except Exception as e:
            self.error.emit(str(e))


class SpecDocResolver:
    """Dispatches specification document requests to local Word, PDF, or Explorer."""

    @staticmethod
    def resolve_spec_files(
        spec_number: str,
        version: str,
        cache_dir: Optional[Path] = None,
        specs_db: Optional[Any] = None,
    ) -> Dict[str, Any]:
        """
        Locates the specification directory and files for the given spec and version.
        """
        clean_spec = re.sub(r"^(?:TS|TR)\s*", "", str(spec_number or "").strip(), flags=re.IGNORECASE)
        clean_ver = str(version or "").lstrip("vV").strip()

        base_cache = Path(cache_dir) if cache_dir else get_specs_root()
        spec_dir = base_cache / clean_spec

        local_files: List[Path] = []
        if spec_dir.exists():
            try:
                local_files = [p for p in spec_dir.iterdir() if p.is_file()]
            except OSError:
                local_files = []

        # If version was not specified, discover latest available version from local files
        if not clean_ver and local_files:
            discovered_versions = []
            for f in local_files:
                if f.name.lower().endswith((".docx", ".doc", ".zip")) and not is_change_mark_file(f.name):
                    m = re.search(r"[-_\s]([a-z0-9]{3})(?:[-_.]|$)", f.name.lower())
                    if m:
                        v_str = letter_code_to_version(m.group(1))
                        if v_str:
                            discovered_versions.append(v_str)
            if discovered_versions:
                from modules.nas.core.nas_db import parse_version_tuple
                discovered_versions.sort(key=parse_version_tuple, reverse=True)
                clean_ver = discovered_versions[0]

        db_filename = None
        db_url = None
        if specs_db and clean_ver:
            try:
                rows = specs_db.search_files(spec_number=clean_spec, release_version=clean_ver)
                if rows:
                    db_filename = rows[0][4]
                    db_url = rows[0][6]
            except Exception as e:
                logger.warning(f"Could not query specs_db for {clean_spec} v{clean_ver}: {e}")

        # Compute fallback stem
        if db_filename:
            stem = Path(db_filename).stem.lower()
        else:
            code = version_to_letter_code(clean_ver)
            num_clean = clean_spec.replace(".", "").replace("-", "")
            stem = f"{num_clean}-{code}".lower() if code else clean_spec.lower()

        word_docs: List[Path] = []
        pdf_docs: List[Path] = []
        zip_file: Optional[Path] = None

        for f in local_files:
            fname_lower = f.name.lower()
            if fname_lower.endswith((".docx", ".doc")):
                if stem and (fname_lower.startswith(stem) or stem in fname_lower):
                    if not is_change_mark_file(f.name):
                        word_docs.append(f)
            elif fname_lower.endswith(".pdf"):
                if stem and (fname_lower.startswith(stem) or stem in fname_lower):
                    pdf_docs.append(f)
            elif fname_lower.endswith(".zip"):
                if db_filename and fname_lower == db_filename.lower():
                    zip_file = f
                elif stem and (fname_lower.startswith(stem) or stem in fname_lower):
                    zip_file = f

        word_docs.sort(key=lambda p: p.name)

        return {
            "spec_number": clean_spec,
            "version": clean_ver,
            "spec_dir": spec_dir,
            "filename": db_filename,
            "url": db_url,
            "stem": stem,
            "word_docs": word_docs,
            "pdf_docs": pdf_docs,
            "zip_file": zip_file,
            "is_downloaded": bool(word_docs or zip_file),
        }

    @staticmethod
    def extract_word_from_zip(zip_path: Path, target_dir: Path) -> List[Path]:
        """
        Extracts clean Word documents from an archive, filtering revision-mark documents.
        """
        extracted: List[Path] = []
        ext_zip_path = to_extended_path(zip_path)
        if not os.path.exists(ext_zip_path):
            return extracted

        try:
            with zipfile.ZipFile(ext_zip_path, "r") as zf:
                for member in zf.namelist():
                    if member.startswith("._") or "__MACOSX" in member:
                        continue
                    if member.lower().endswith((".docx", ".doc")) and not is_change_mark_file(member):
                        target_file = target_dir / Path(member).name
                        ext_target_file = to_extended_path(target_file)
                        if not os.path.exists(ext_target_file):
                            with open(ext_target_file, "wb") as f:
                                f.write(zf.read(member))
                        if target_file not in extracted:
                            extracted.append(target_file)
        except Exception as e:
            logger.error(f"Failed to extract {zip_path}: {e}")

        extracted.sort(key=lambda p: p.name)
        return extracted

    @staticmethod
    def open_spec(
        spec_number: str,
        version: str,
        clause_number: Optional[str] = None,
        doc_format: str = "word",
        cache_dir: Optional[Path] = None,
        specs_db: Optional[Any] = None,
        parent_widget: Optional[QWidget] = None,
        on_opened_callback: Optional[Any] = None,
    ) -> bool:
        """
        Main entry point to resolve and launch a specification document.
        Assists search by copying 'Clause X.Y.Z' to clipboard when opening Word.
        """
        info = SpecDocResolver.resolve_spec_files(spec_number, version, cache_dir, specs_db)
        spec_dir = info["spec_dir"]

        # --- 1. OPEN FOLDER ---
        if doc_format == "folder":
            if spec_dir.exists():
                try:
                    os.startfile(str(spec_dir))
                    return True
                except Exception as e:
                    logger.error(f"Failed to open folder {spec_dir}: {e}")
            else:
                QMessageBox.information(
                    parent_widget,
                    "Folder Not Found",
                    f"The local folder for TS {spec_number} has not been created yet:\n{spec_dir}",
                )
            return False

        # --- 2. OPEN PDF ---
        if doc_format == "pdf":
            if info["pdf_docs"]:
                try:
                    open_document(info["pdf_docs"][0])
                    return True
                except Exception as e:
                    logger.error(f"Failed to open PDF {info['pdf_docs'][0]}: {e}")
                    QMessageBox.warning(parent_widget, "Open Error", f"Could not open PDF:\n{e}")
                    return False
            else:
                QMessageBox.information(
                    parent_widget,
                    "PDF Not Available",
                    f"No converted PDF document found for TS {spec_number} v{version}.\n\n"
                    f"You can open the Word (.docx) version instead, or convert it via the Specifications tab.",
                )
                return False

        # --- 3. OPEN WORD (.docx / .doc) ---
        target_doc = None
        if info["word_docs"]:
            target_doc = match_multipart_document_for_clause(info["word_docs"], clause_number)
        elif info["zip_file"] and info["zip_file"].exists():
            # Extract from ZIP on the fly
            extracted = SpecDocResolver.extract_word_from_zip(info["zip_file"], spec_dir)
            if extracted:
                target_doc = match_multipart_document_for_clause(extracted, clause_number)

        if target_doc and target_doc.exists():
            # Clipboard search assistance for delegates
            clip_text = f"Clause {clause_number}" if clause_number else f"TS {spec_number}"
            clipboard = QApplication.clipboard()
            if clipboard and clause_number:
                clipboard.setText(clause_number.strip())

            try:
                open_document(target_doc)
                msg = f"Opening '{target_doc.name}'"
                if clause_number:
                    msg += f" — '{clause_number}' copied to clipboard (press Ctrl+F in Word to jump directly)."
                logger.info(f"✅ {msg}")

                if parent_widget and hasattr(parent_widget, "matrix_title") and not hasattr(parent_widget, "current_selected_message_name"):
                    parent_widget.matrix_title.setText(f"📖 {msg}")

                if on_opened_callback:
                    on_opened_callback(target_doc)
                return True
            except Exception as e:
                logger.error(f"Failed to open document {target_doc}: {e}")
                QMessageBox.critical(parent_widget, "Word Open Error", f"Could not launch Microsoft Word:\n{e}")
                return False

        # --- 4. FILE NOT PRESENT: OFFER DOWNLOAD ---
        if not info["url"]:
            QMessageBox.warning(
                parent_widget,
                "Document Not Found",
                f"No local Word document or ZIP archive found for TS {spec_number} v{version}.\n\n"
                f"Please open the '📚 Specifications' tab to locate or download this release.",
            )
            return False

        reply = QMessageBox.question(
            parent_widget,
            "Download Specification",
            f"TS {spec_number} v{version} is not yet downloaded locally.\n\n"
            f"Would you like to download it now from 3GPP FTP?\n"
            f"File: {info['filename'] or 'Archive'}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply != QMessageBox.Yes:
            return False

        # Background download with progress dialog
        zip_target = spec_dir / (info["filename"] or f"{info['stem']}.zip")
        prog = QProgressDialog(f"Downloading TS {spec_number} v{version} from 3GPP FTP...", "Cancel", 0, 0, parent_widget)
        prog.setWindowTitle("Downloading 3GPP Specification")
        prog.setWindowModality(1)  # Qt.WindowModal

        worker = SpecDownloadWorker(info["url"], zip_target)

        def on_download_success(downloaded_zip: Path):
            prog.close()
            extracted = SpecDocResolver.extract_word_from_zip(downloaded_zip, spec_dir)
            if extracted:
                final_doc = match_multipart_document_for_clause(extracted, clause_number)
                if final_doc:
                    if clause_number:
                        QApplication.clipboard().setText(clause_number.strip())
                    open_document(final_doc)
                    logger.info(f"✅ Downloaded and opened '{final_doc.name}'")
            else:
                QMessageBox.warning(parent_widget, "Extraction Issue", "Downloaded ZIP did not contain clean Word documents.")

        def on_download_error(err_str: str):
            prog.close()
            QMessageBox.critical(parent_widget, "Download Failed", f"Could not download specification:\n{err_str}")

        worker.finished_success.connect(on_download_success)
        worker.error.connect(on_download_error)
        prog.canceled.connect(worker.requestInterruption)

        worker.start()
        prog.exec_()
        return True
