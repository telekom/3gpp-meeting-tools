"""On-demand PDF resolution, confirmation, download, and Word-to-PDF conversion helpers."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PyQt5.QtCore import Qt, QThread
from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressDialog, QWidget

from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)


def find_matching_pdf(
    info: Dict[str, Any],
    clause_number: Optional[str],
    multipart_matcher: Callable[[List[Path], Optional[str]], Optional[Path]],
) -> Optional[Path]:
    """Locates the converted PDF matching the target clause (multi-part aware)."""
    pdf_docs: List[Path] = info.get("pdf_docs") or []
    word_docs: List[Path] = info.get("word_docs") or []

    if word_docs:
        expected_word = multipart_matcher(word_docs, clause_number)
        if expected_word:
            expected_stem = expected_word.stem.lower()
            for pdf in pdf_docs:
                if pdf.stem.lower() == expected_stem and pdf.exists():
                    return pdf
            sibling_pdf = expected_word.with_suffix(".pdf")
            if sibling_pdf.exists():
                return sibling_pdf
            return None

    if pdf_docs:
        return multipart_matcher(pdf_docs, clause_number)
    return None


def track_doc_worker(parent_widget: Optional[QWidget], worker: QThread) -> None:
    """Retains active document worker reference on parent_widget until finished."""
    if parent_widget is not None:
        workers = getattr(parent_widget, "_active_doc_workers", None)
        if workers is None:
            workers = []
            setattr(parent_widget, "_active_doc_workers", workers)
        workers.append(worker)

    def _cleanup(w=worker, pw=parent_widget) -> None:
        if pw is not None:
            active = getattr(pw, "_active_doc_workers", None)
            if isinstance(active, list) and w in active:
                active.remove(w)
        w.deleteLater()

    worker.finished.connect(_cleanup)


def open_pdf_with_feedback(
    pdf_path: Path,
    clause_number: Optional[str] = None,
    parent_widget: Optional[QWidget] = None,
    on_opened_callback: Optional[Any] = None,
) -> bool:
    """Copies clause number for Ctrl+F, opens the PDF, and updates status feedback."""
    clipboard = QApplication.clipboard()
    if clipboard and clause_number:
        clipboard.setText(clause_number.strip())

    try:
        open_document(pdf_path)
        msg = f"Opening '{pdf_path.name}'"
        if clause_number:
            msg += (
                f" — '{clause_number.strip()}' copied to clipboard "
                "(press Ctrl+F in PDF reader to jump directly)."
            )
        logger.info(f"✅ {msg}")

        if (
            parent_widget
            and hasattr(parent_widget, "matrix_title")
            and not hasattr(parent_widget, "current_selected_message_name")
        ):
            parent_widget.matrix_title.setText(f"📕 {msg}")

        if on_opened_callback:
            on_opened_callback(pdf_path)
        return True
    except Exception as e:
        logger.error(f"Failed to open PDF {pdf_path}: {e}")
        QMessageBox.warning(parent_widget, "Open Error", f"Could not open PDF:\n{e}")
        return False


def convert_word_to_pdf_and_open(
    word_doc: Path,
    spec_number: str,
    version: str,
    clause_number: Optional[str] = None,
    parent_widget: Optional[QWidget] = None,
    on_opened_callback: Optional[Any] = None,
) -> bool:
    """Converts a local Word specification document to PDF via WordConverterThread and opens it."""
    from modules.word_tools.core.word_converter import WordConverterThread

    prog = QProgressDialog(
        f"Converting TS {spec_number} v{version} ({word_doc.name}) to PDF...",
        "Cancel",
        0,
        0,
        parent_widget,
    )
    prog.setWindowTitle("Converting Specification to PDF")
    prog.setWindowModality(Qt.WindowModal)

    worker = WordConverterThread(str(word_doc), "pdf")
    track_doc_worker(parent_widget, worker)
    state = {"done": False, "success": False}

    def _on_log(msg: str, _level: int = 0) -> None:
        if not state["done"]:
            try:
                prog.setLabelText(str(msg))
            except RuntimeError:
                pass

    def _on_success(out_path: str) -> None:
        state["done"] = True
        prog.close()
        state["success"] = open_pdf_with_feedback(
            Path(out_path),
            clause_number=clause_number,
            parent_widget=parent_widget,
            on_opened_callback=on_opened_callback,
        )

    def _on_failed(err: str) -> None:
        state["done"] = True
        prog.close()
        QMessageBox.critical(
            parent_widget,
            "PDF Conversion Failed",
            f"Could not convert TS {spec_number} v{version} to PDF:\n{err}",
        )

    worker.ui_log_msg.connect(_on_log)
    worker.finished_path.connect(_on_success)
    worker.conversion_failed.connect(_on_failed)
    prog.canceled.connect(worker.requestInterruption)

    worker.start()
    if not state["done"]:
        prog.exec_()
    return state["success"]


def handle_pdf_open_request(
    resolver_cls: Any,
    info: Dict[str, Any],
    spec_number: str,
    version: str,
    clause_number: Optional[str],
    cache_dir: Optional[Path],
    specs_db: Optional[Any],
    parent_widget: Optional[QWidget],
    on_opened_callback: Optional[Any],
    multipart_matcher: Callable[[List[Path], Optional[str]], Optional[Path]],
    download_worker_cls: Any,
) -> bool:
    """Opens cached PDF or prompts the user to convert/download+convert on demand."""
    pdf_target = resolver_cls.find_matching_pdf(info, clause_number)
    if pdf_target and pdf_target.exists():
        return resolver_cls._open_pdf_with_feedback(
            pdf_target,
            spec_number=spec_number,
            clause_number=clause_number,
            parent_widget=parent_widget,
            on_opened_callback=on_opened_callback,
        )

    word_target = multipart_matcher(info["word_docs"], clause_number) if info["word_docs"] else None
    has_local_zip = bool(info["zip_file"] and info["zip_file"].exists())
    has_remote_url = bool(info["url"])

    if not (word_target or has_local_zip or has_remote_url):
        QMessageBox.information(
            parent_widget,
            "PDF Not Available",
            f"No converted PDF or local/FTP Word source was found for TS {spec_number} v{version}.\n\n"
            "Please open the Specifications tab to locate or download this release.",
        )
        return False

    source_hint = (
        f"Would you like to convert '{word_target.name}' to PDF now?"
        if word_target
        else (
            f"Would you like to extract and convert TS {spec_number} v{version} to PDF now?"
            if has_local_zip
            else f"Would you like to download TS {spec_number} v{version} from 3GPP FTP and convert it to PDF now?"
        )
    )
    box = QMessageBox(parent_widget)
    box.setIcon(QMessageBox.Question)
    box.setWindowTitle("PDF Not Available")
    box.setText(f"No converted PDF document found for TS {spec_number} v{version}.")
    box.setInformativeText(source_hint)
    btn_convert = box.addButton("Convert && Open PDF", QMessageBox.AcceptRole)
    btn_word = box.addButton("Open Word Instead", QMessageBox.ActionRole)
    box.addButton(QMessageBox.Cancel)
    box.setDefaultButton(btn_convert)
    box.exec_()

    clicked = box.clickedButton()
    if clicked == btn_word:
        return resolver_cls.open_spec(
            spec_number, version, clause_number, "word",
            cache_dir, specs_db, parent_widget, on_opened_callback,
        )
    if clicked != btn_convert:
        return False

    if word_target and word_target.exists():
        return resolver_cls._convert_word_to_pdf_and_open(
            word_target, spec_number, version, clause_number, parent_widget, on_opened_callback
        )

    spec_dir = info["spec_dir"]
    if has_local_zip:
        extracted = resolver_cls.extract_word_from_zip(info["zip_file"], spec_dir)
        target = multipart_matcher(extracted, clause_number)
        if target and target.exists():
            return resolver_cls._convert_word_to_pdf_and_open(
                target, spec_number, version, clause_number, parent_widget, on_opened_callback
            )
        QMessageBox.warning(
            parent_widget, "Extraction Issue", "Local ZIP archive did not contain clean Word documents."
        )
        return False

    zip_target = spec_dir / (info["filename"] or f"{info['stem']}.zip")
    prog = QProgressDialog(
        f"Downloading TS {spec_number} v{version} from 3GPP FTP...", "Cancel", 0, 0, parent_widget
    )
    prog.setWindowTitle("Downloading 3GPP Specification")
    prog.setWindowModality(Qt.WindowModal)

    dl_worker = download_worker_cls(info["url"], zip_target)
    track_doc_worker(parent_widget, dl_worker)
    dl_state: Dict[str, Any] = {"done": False, "downloaded": None}

    def _on_dl_ok(downloaded_zip: Path) -> None:
        dl_state["done"] = True
        dl_state["downloaded"] = downloaded_zip
        prog.close()

    def _on_dl_err(err_str: str) -> None:
        dl_state["done"] = True
        prog.close()
        QMessageBox.critical(parent_widget, "Download Failed", f"Could not download specification:\n{err_str}")

    dl_worker.finished_success.connect(_on_dl_ok)
    dl_worker.error.connect(_on_dl_err)
    prog.canceled.connect(dl_worker.requestInterruption)

    dl_worker.start()
    if not dl_state["done"]:
        prog.exec_()

    downloaded_zip = dl_state["downloaded"]
    if not downloaded_zip:
        return False

    extracted = resolver_cls.extract_word_from_zip(downloaded_zip, spec_dir)
    target = multipart_matcher(extracted, clause_number)
    if target and target.exists():
        return resolver_cls._convert_word_to_pdf_and_open(
            target, spec_number, version, clause_number, parent_widget, on_opened_callback
        )
    QMessageBox.warning(
        parent_widget, "Extraction Issue", "Downloaded ZIP did not contain clean Word documents."
    )
    return False
