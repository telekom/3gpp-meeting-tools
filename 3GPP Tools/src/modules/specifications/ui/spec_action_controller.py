# --- File: src/modules/specifications/ui/spec_action_controller.py ---
"""
Document and ZIP archive action controller for specifications.
Handles asynchronous archive downloading, zip extraction, format conversions
(Word, PDF, HTML, TXT, Markdown), and system file dispatching.
"""
import logging
import os
import zipfile
from pathlib import Path

from PyQt5 import sip
from PyQt5.QtWidgets import QComboBox, QMessageBox, QPushButton

from core.utils.paths import to_extended_path
from modules.specifications.ui.threads import (
    SpecDownloadThread,
    SpecMarkdownConverterThread,
)
from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)


def handle_document_action(tab, combo: QComboBox, doc_type: str, btn: QPushButton):
    """
    Dispatches document actions (open Word, convert to PDF/HTML/TXT/MD, or download archive).

    :param tab: SpecificationsTab instance providing download_dir, db, and _download_threads.
    :param combo: QComboBox representing the specification version dropdown.
    :param doc_type: Target document format ('word', 'pdf', 'html', 'txt', 'md', 'md_split').
    :param btn: Triggering action QPushButton for state indication.
    """
    if sip.isdeleted(combo):
        return
    c_data = combo.currentData()
    if not c_data:
        return

    spec_dl_dir = Path(tab.download_dir) / c_data["spec_num"]
    zip_path = spec_dl_dir / c_data["fname"]
    stem = Path(c_data["fname"]).stem

    def _process_and_open():
        extracted_docs = []

        ext_zip_path = to_extended_path(zip_path)
        if os.path.exists(ext_zip_path):
            try:
                with zipfile.ZipFile(ext_zip_path, "r") as z:
                    for member in z.namelist():
                        if "__MACOSX" in member or member.startswith("._"):
                            continue

                        if member.lower().endswith((".doc", ".docx")):
                            target_file = spec_dl_dir / Path(member).name
                            ext_target_file = to_extended_path(target_file)
                            if not os.path.exists(ext_target_file):
                                with open(ext_target_file, "wb") as f:
                                    f.write(z.read(member))

                            if target_file not in extracted_docs:
                                extracted_docs.append(target_file)
            except Exception as e:
                QMessageBox.warning(tab, "Extraction Error", f"Failed to extract archive:\n{e}")
                return

        if not extracted_docs:
            extracted_docs = list(spec_dl_dir.glob(f"{stem}*.doc*"))

        if not extracted_docs:
            QMessageBox.warning(tab, "Not Found", "No Word documents found on disk or inside the zip archive.")
            return

        try:
            from modules.word_tools.core.word_converter import WordConverterThread
        except ImportError as e:
            QMessageBox.warning(tab, "Import Error", f"Could not import word_converter:\n{e}")
            return

        for doc_path in extracted_docs:
            try:
                if doc_type == "word":
                    open_document(doc_path)

                    # Refresh badges to show 📝 immediately after Word extraction
                    try:
                        if not sip.isdeleted(combo):
                            tab._update_combo_item_badges(combo, combo.currentIndex())
                            combo.currentIndexChanged.emit(combo.currentIndex())
                    except RuntimeError:
                        pass

                elif doc_type in ("pdf", "html", "txt"):
                    target_ext = f".{doc_type}"
                    target_path = doc_path.with_suffix(target_ext)

                    if not target_path.exists():
                        orig_text = btn.text() if not sip.isdeleted(btn) else "Converting..."
                        if not sip.isdeleted(btn):
                            btn.setText("⏳ Converting...")
                            btn.setEnabled(False)

                        conv_thread = WordConverterThread(str(doc_path), doc_type)

                        def on_success(p, c=combo, b=btn, txt=orig_text):
                            try:
                                os.startfile(p)
                            except Exception as e:
                                logger.error(f"Error opening converted file: {e}", exc_info=True)

                            try:
                                if not sip.isdeleted(c):
                                    tab._update_combo_item_badges(c, c.currentIndex())
                                    c.currentIndexChanged.emit(c.currentIndex())
                                if not sip.isdeleted(b):
                                    b.setText(txt)
                                    b.setEnabled(True)
                            except RuntimeError:
                                pass

                        conv_thread.finished_path.connect(on_success)

                        def cleanup(t=conv_thread, b=btn, txt=orig_text):
                            if t in tab._download_threads:
                                tab._download_threads.remove(t)
                            try:
                                if not sip.isdeleted(b) and not b.isEnabled():
                                    b.setText(txt)
                                    b.setEnabled(True)
                            except RuntimeError:
                                pass

                        conv_thread.finished.connect(cleanup)
                        tab._download_threads.append(conv_thread)
                        conv_thread.start()

                    else:
                        os.startfile(str(target_path))

                elif doc_type in ("md", "md_split"):
                    is_split = (doc_type == "md_split")
                    target_path = spec_dl_dir / f"{stem}_clauses" / "index.md" if is_split else spec_dl_dir / f"{stem}.md"

                    if not target_path.exists():
                        orig_text = btn.text() if not sip.isdeleted(btn) else "Converting..."
                        if not sip.isdeleted(btn):
                            btn.setText("⏳ Converting MD...")
                            btn.setEnabled(False)

                        spec_meta = {
                            "spec": c_data.get("spec_num", ""),
                            "version": c_data.get("version", ""),
                            "release": c_data.get("release", ""),
                            "upload_date": c_data.get("upload_date", ""),
                        }
                        spec_details = tab.db.get_spec_details(c_data.get("spec_num", ""))
                        if spec_details:
                            spec_meta["title"] = spec_details.get("title", "")
                            spec_meta["wg"] = spec_details.get("wg", "")

                        mode = "split" if is_split else "single"
                        md_thread = SpecMarkdownConverterThread(spec_dl_dir, stem, mode=mode, meta=spec_meta)

                        def on_md_progress(status_msg, b=btn):
                            try:
                                if not sip.isdeleted(b):
                                    b.setText(f"⏳ {status_msg}"[:25])
                            except RuntimeError:
                                pass

                        def on_md_success(p, msg, c=combo, b=btn, txt=orig_text):
                            try:
                                os.startfile(str(p))
                            except Exception as e:
                                logger.error(f"Error opening converted Markdown file: {e}", exc_info=True)

                            try:
                                if not sip.isdeleted(c):
                                    tab._update_combo_item_badges(c, c.currentIndex())
                                    c.currentIndexChanged.emit(c.currentIndex())
                                if not sip.isdeleted(b):
                                    b.setText(txt)
                                    b.setEnabled(True)
                            except RuntimeError:
                                pass

                        def on_md_error(err, b=btn, txt=orig_text):
                            logger.error(f"Spec Markdown conversion failed: {err}")
                            QMessageBox.warning(tab, "Markdown Error", f"Failed to convert specification:\n{err}")
                            try:
                                if not sip.isdeleted(b):
                                    b.setText(txt)
                                    b.setEnabled(True)
                            except RuntimeError:
                                pass

                        def md_cleanup(t=md_thread, b=btn, txt=orig_text):
                            if t in tab._download_threads:
                                tab._download_threads.remove(t)
                            try:
                                if not sip.isdeleted(b) and not b.isEnabled():
                                    b.setText(txt)
                                    b.setEnabled(True)
                            except RuntimeError:
                                pass

                        md_thread.progress_status.connect(on_md_progress)
                        md_thread.finished_success.connect(on_md_success)
                        md_thread.error.connect(on_md_error)
                        md_thread.finished.connect(md_cleanup)

                        tab._download_threads.append(md_thread)
                        md_thread.start()
                    else:
                        os.startfile(str(target_path))

                    break

            except Exception as e:
                QMessageBox.warning(tab, "Open Error", f"Could not open/convert {doc_type.upper()}:\n{e}")

    word_exists = any(spec_dl_dir.glob(f"{stem}*.doc*"))
    target_exists = False

    if doc_type == "word":
        target_exists = word_exists
    elif doc_type == "pdf":
        target_exists = any(spec_dl_dir.glob(f"{stem}*.pdf"))
    elif doc_type == "html":
        target_exists = any(spec_dl_dir.glob(f"{stem}*.html"))
    elif doc_type == "txt":
        target_exists = any(spec_dl_dir.glob(f"{stem}*.txt"))
    elif doc_type == "md":
        target_exists = any(f for f in spec_dl_dir.glob(f"{stem}*.md") if "_clauses" not in str(f))
    elif doc_type == "md_split":
        target_exists = (spec_dl_dir / f"{stem}_clauses" / "index.md").exists()

    if target_exists or word_exists or zip_path.exists():
        _process_and_open()
    else:
        os.makedirs(to_extended_path(spec_dl_dir), exist_ok=True)

        idx = combo.currentIndex()
        combo.setItemText(idx, "⏳ Downloading...")
        combo.setEnabled(False)

        if not sip.isdeleted(btn):
            btn.setText("⏳ Downloading...")
            btn.setEnabled(False)

        thread = SpecDownloadThread(c_data["url"], zip_path)

        def _on_success(zp):
            try:
                if not sip.isdeleted(combo):
                    c_data_inner = combo.itemData(idx)
                    if c_data_inner:
                        c_data_inner["is_downloaded"] = True
                        combo.setItemData(idx, c_data_inner)
                    tab._update_combo_item_badges(combo, idx)
                    combo.setEnabled(True)
                if not sip.isdeleted(btn):
                    btn.setEnabled(True)
            except RuntimeError:
                pass

            _process_and_open()

            try:
                if not sip.isdeleted(combo):
                    combo.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass

        def _on_err(err):
            try:
                if not sip.isdeleted(combo):
                    tab._update_combo_item_badges(combo, idx)
                    combo.setEnabled(True)
                if not sip.isdeleted(btn):
                    btn.setEnabled(True)
                    combo.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass
            QMessageBox.critical(tab, "Download Failed", f"Network error:\n{err}")

        thread.finished_success.connect(_on_success)
        thread.error.connect(_on_err)

        tab._download_threads.append(thread)
        thread.finished.connect(
            lambda t=thread: tab._download_threads.remove(t) if t in tab._download_threads else None
        )
        thread.start()


def handle_zip_action(tab, combo: QComboBox, btn: QPushButton):
    """
    Dispatches ZIP actions (opens local directory if zip exists, or downloads archive).

    :param tab: SpecificationsTab instance providing download_dir and _download_threads.
    :param combo: QComboBox representing the specification version dropdown.
    :param btn: Triggering action QPushButton for state indication.
    """
    if sip.isdeleted(combo):
        return
    c_data = combo.currentData()
    if not c_data:
        return

    spec_dl_dir = Path(tab.download_dir) / c_data["spec_num"]
    zip_path = spec_dl_dir / c_data["fname"]

    if zip_path.exists():
        try:
            os.startfile(str(spec_dl_dir))
        except Exception as e:
            QMessageBox.warning(tab, "Explorer Error", f"Could not open directory:\n{e}")
        return

    os.makedirs(to_extended_path(spec_dl_dir), exist_ok=True)

    idx = combo.currentIndex()
    combo.setItemText(idx, "⏳ Downloading...")
    combo.setEnabled(False)

    if not sip.isdeleted(btn):
        btn.setText("⏳ Downloading...")
        btn.setEnabled(False)

    thread = SpecDownloadThread(c_data["url"], zip_path)

    def _on_success(zp):
        try:
            if not sip.isdeleted(combo):
                c_data_inner = combo.itemData(idx)
                if c_data_inner:
                    c_data_inner["is_downloaded"] = True
                    combo.setItemData(idx, c_data_inner)
                tab._update_combo_item_badges(combo, idx)
                combo.setEnabled(True)
            if not sip.isdeleted(btn):
                btn.setEnabled(True)
                combo.currentIndexChanged.emit(idx)
        except RuntimeError:
            pass

    def _on_err(err):
        try:
            if not sip.isdeleted(combo):
                tab._update_combo_item_badges(combo, idx)
                combo.setEnabled(True)
            if not sip.isdeleted(btn):
                btn.setEnabled(True)
                combo.currentIndexChanged.emit(idx)
        except RuntimeError:
            pass
        QMessageBox.critical(tab, "Download Failed", f"Network error:\n{err}")

    thread.finished_success.connect(_on_success)
    thread.error.connect(_on_err)

    tab._download_threads.append(thread)
    thread.finished.connect(
        lambda t=thread: tab._download_threads.remove(t) if t in tab._download_threads else None
    )
    thread.start()
