# --- File: src/modules/meetings/ui/tdocs_agenda_controller.py ---
"""
Agenda and Chairman's Notes import/download controller for TDocsWindow.
Manages batch Word document agenda importing, CSV agenda synchronization,
and Chairman's Notes retrieval threads.
"""
import logging
import re
from pathlib import Path

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from core.network.network_state import NetworkState
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.agenda_manager import AgendaDownloaderThread, AgendaManager
from modules.meetings.core.chair_notes_downloader import (
    ChairNotesDownloaderThread,
    ChairNotesListingThread,
)
from modules.meetings.core.tdocs.tdocs_threads import WordAgendaImporterThread
from modules.meetings.core.url_router import URLRouter

logger = logging.getLogger(__name__)


def import_word_agenda_dialog(window):
    """Prompts file dialog to select agenda / chair notes Word, ZIP, or HTML files to import."""
    agenda_dir = window.meeting_dir / "Agenda"
    start_dir = str(agenda_dir) if agenda_dir.exists() else str(window.meeting_dir)

    file_paths, _ = QFileDialog.getOpenFileNames(
        window,
        "Select TDocsByAgenda / Chairman's Notes File(s)",
        start_dir,
        "Supported Files (*.docx *.doc *.zip *.htm *.html);;Word Documents (*.docx *.doc);;ZIP Archives (*.zip);;HTML Files (*.htm *.html);;All Files (*.*)",
    )
    if file_paths:
        process_imported_agenda_files(window, [Path(p) for p in file_paths])


def process_imported_agenda_files(window, source_paths: list):
    """Queues files for sequential background Word agenda ingestion."""
    valid_paths = [p for p in source_paths if p.exists()]
    if not valid_paths:
        QMessageBox.warning(window, "Files Not Found", "None of the specified files exist on disk.")
        return

    window._import_queue.extend(valid_paths)
    if not window._is_importing_agenda:
        window._total_import_batch_count = len(window._import_queue)
        window._import_accumulated_merged = 0
        window._import_accumulated_agreements = 0
        window._import_success_files = []
        window._import_failed_files = []
        window._is_importing_agenda = True
        process_next_in_import_queue(window)


def process_next_in_import_queue(window):
    """Processes the next file in the agenda import batch queue."""
    if not window._import_queue:
        finalize_import_batch(window)
        return

    current_file = window._import_queue.pop(0)
    current_idx = window._total_import_batch_count - len(window._import_queue)

    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText(f"⏳ ({current_idx}/{window._total_import_batch_count}) Importing...")

    window.word_import_thread = WordAgendaImporterThread(
        current_file,
        window.meeting_dir,
        mtg_info=getattr(window, "mtg_info", None),
        agenda_map=getattr(window, "agenda_map", None),
    )
    window.word_import_thread.progress.connect(
        lambda msg, idx=current_idx, tot=window._total_import_batch_count: on_word_import_progress(
            window, f"({idx}/{tot}) {msg}"
        )
    )
    window.word_import_thread.finished.connect(
        lambda success, data, fname, err: on_word_import_finished(window, success, data, fname, err)
    )
    window.word_import_thread.start()


def on_word_import_progress(window, msg: str):
    """Updates button text with import progress status."""
    window.refresh_btn.setText(f"⏳ {msg}"[:32])


def on_word_import_finished(window, success: bool, agenda_data: dict, filename: str, error_msg: str):
    """Handles parsing result for a single agenda file and proceeds to next in queue."""
    extracted_ags = getattr(window.word_import_thread, "extracted_agreements", []) if hasattr(window, "word_import_thread") else []
    # In SA/CT meetings, do not extract individual pCR approvals as RAN-style agreements
    if not getattr(window, "is_ran", False):
        extracted_ags = []

    ag_count = len([a for a in (extracted_ags or []) if not a.get("is_previous_meeting_recap")])
    if success and (agenda_data or ag_count > 0):
        if extracted_ags and getattr(window, "is_ran", False):
            window.chair_agreements = list(getattr(window, "chair_agreements", [])) + extracted_ags
            window.model.set_chair_agreements(window.chair_agreements)
        if agenda_data:
            window.model.merge_agenda_data(agenda_data)
            window._import_accumulated_merged += len(agenda_data)
        window._import_accumulated_agreements = getattr(window, "_import_accumulated_agreements", 0) + ag_count
        window._import_success_files.append((filename, len(agenda_data), ag_count))
    else:
        window._import_failed_files.append((filename, error_msg or "No valid TDocs table or agreements found"))

    process_next_in_import_queue(window)


def finalize_import_batch(window):
    """Resets UI state and shows completion summary for agenda batch ingestion."""
    window._is_importing_agenda = False
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    ag_total = getattr(window, "_import_accumulated_agreements", 0)
    if hasattr(window, "_update_agreements_counter"):
        window._update_agreements_counter()
    if window._import_accumulated_merged > 0 or ag_total > 0:
        window._refresh_comboboxes()
        window.refresh_btn.setText(f"✅ {window._import_accumulated_merged} Merged")
        QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))

    if window._total_import_batch_count == 1:
        if window._import_success_files:
            entry = window._import_success_files[0]
            fname = entry[0]
            count = entry[1]
            ag_cnt = entry[2] if len(entry) > 2 else 0
            ag_line = f"\n• Extracted {ag_cnt} structured agreement(s) / working assumption(s)" if ag_cnt > 0 else ""
            QMessageBox.information(
                window,
                "Import Successful",
                f"Saved to Agenda/\nSuccessfully merged {count} TDoc entries{ag_line} from:\n{fname}",
            )
        else:
            fname, err = window._import_failed_files[0]
            QMessageBox.warning(window, "Import Failed", f"Failed to import {fname}:\n{err}")
        return

    summary_lines = [
        f"Batch import completed for {window._total_import_batch_count} files.",
        f"• Successfully imported: {len(window._import_success_files)} file(s)",
        f"• Total TDocs merged: {window._import_accumulated_merged}",
    ]
    if ag_total > 0:
        summary_lines.append(f"• Total Agreements / Working Assumptions extracted: {ag_total}")
    summary_lines.append("")

    if window._import_failed_files:
        summary_lines.append(f"⚠️ Failures ({len(window._import_failed_files)}):")
        for fname, err in window._import_failed_files:
            summary_lines.append(f"  - {fname}: {err}")
        QMessageBox.warning(window, "Batch Import Completed with Warnings", "\n".join(summary_lines))
    else:
        QMessageBox.information(window, "Batch Import Successful", "\n".join(summary_lines))


def download_agenda_csv(window):
    """Initiates asynchronous download of the official 3GPP meeting agenda CSV file."""
    wg_name = window.mtg_info.get("wg_name", "").upper()
    main_url = window.mtg_info.get("url_key", "")
    if main_url and not main_url.startswith("http"):
        main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

    is_active = window.mtg_info.get("is_active_sync", False)
    candidate_urls = []
    if NetworkState.get_instance().is_local_active():
        local_base = URLRouter._get_local_server_base(wg_name)
        candidate_urls.append(local_base)

    if is_active:
        sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
        candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

    if main_url:
        candidate_urls.append(main_url)

    docs_url = window.mtg_info.get("docs_folder_url", "")
    if docs_url:
        if not docs_url.startswith("http"):
            docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip('/')
        base_from_docs = re.sub(r'/Docs/?$', '', docs_url, flags=re.IGNORECASE)
        if base_from_docs not in candidate_urls:
            candidate_urls.append(base_from_docs)

    if not candidate_urls:
        return QMessageBox.warning(
            window, "No URL Available", "Cannot download Agenda CSV: No valid meeting URL found."
        )

    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText("⏳ Getting Agenda...")

    window.agenda_dl_thread = AgendaDownloaderThread(candidate_urls, window.agenda_dir, window)
    window.agenda_dl_thread.progress.connect(lambda msg: window.refresh_btn.setText(f"⏳ {msg}"[:30]))
    window.agenda_dl_thread.finished.connect(
        lambda success, msg: on_agenda_download_finished(window, success, msg)
    )
    window.agenda_dl_thread.start()


def on_agenda_download_finished(window, success: bool, msg: str):
    """Handles completion of the agenda CSV download."""
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if success:
        window.refresh_btn.setText("✅ Agenda Cached")
        QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))
        window.agenda_map = AgendaManager.load_local_agenda(window.agenda_dir)
        if hasattr(window, "model") and window.model:
            window.model.set_agenda_map(window.agenda_map)
        window._refresh_comboboxes()
        QMessageBox.information(window, "Agenda CSV Cached", msg)
    else:
        QMessageBox.warning(window, "Download Failed", msg)


def download_chair_notes(window):
    """
    Initiates asynchronous scanning of /Inbox/Chair_Notes or /Inbox/Chairs_Notes
    and displays the multi-chair picker dialog so delegates can choose which notes to download.
    """
    wg_name = window.mtg_info.get("wg_name", "").upper()
    main_url = window.mtg_info.get("url_key", "")
    if main_url and not main_url.startswith("http"):
        main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

    is_active = window.mtg_info.get("is_active_sync", False)
    candidate_urls = []
    if NetworkState.get_instance().is_local_active():
        local_base = URLRouter._get_local_server_base(wg_name)
        candidate_urls.append(local_base)

    if is_active:
        sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
        candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

    if main_url:
        candidate_urls.append(main_url)

    if not candidate_urls:
        return QMessageBox.warning(
            window, "No URL Available", "Cannot download Chairman's Notes: No valid meeting URL found."
        )

    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText("⏳ Checking Notes...")

    window.chair_notes_listing_thread = ChairNotesListingThread(candidate_urls, window)
    window.chair_notes_listing_thread.progress.connect(
        lambda msg: window.refresh_btn.setText(f"⏳ {msg}"[:30])
    )
    window.chair_notes_listing_thread.finished.connect(
        lambda success, files, dir_url, msg: on_chair_notes_listing_finished(
            window, success, files, dir_url, msg, candidate_urls
        )
    )
    window.chair_notes_listing_thread.start()


def on_chair_notes_listing_finished(
    window, success: bool, files: list, dir_url: str, msg: str, candidate_urls: list
):
    """Handles completion of server directory scanning and presents the picker dialog."""
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if not success:
        QMessageBox.warning(window, "Notes Not Found", msg)
        return

    if not files:
        QMessageBox.information(
            window,
            "No Chairman's Notes Found",
            msg or "Connected to Chair_Notes folder, but no matching Chairman's Notes files were found on the server.",
        )
        return

    chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"

    from modules.meetings.ui.chair_notes_picker_dialog import ChairNotesPickerDialog

    dlg = ChairNotesPickerDialog(
        parent=window,
        file_tuples=files,
        local_dir=chair_notes_dir,
        mtg_info=getattr(window, "mtg_info", None),
        target_dir_url=dir_url,
    )

    if dlg.exec_() == QDialog.Accepted:
        selected_files = dlg.get_selected_files()
        if not selected_files:
            return
        overwrite = dlg.should_overwrite()
        start_chair_notes_download(window, candidate_urls, chair_notes_dir, selected_files, overwrite)


def start_chair_notes_download(
    window, candidate_urls: list, chair_notes_dir: Path, selected_files: list, overwrite: bool
):
    """Launches the download thread for the user-selected Chairman's Notes files."""
    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText("⏳ Downloading Notes...")

    window.chair_notes_thread = ChairNotesDownloaderThread(
        candidate_base_urls=candidate_urls,
        target_dir=chair_notes_dir,
        parent=window,
        selected_files=selected_files,
        overwrite=overwrite,
    )
    window.chair_notes_thread.progress.connect(
        lambda msg: window.refresh_btn.setText(f"⏳ {msg}"[:30])
    )
    window.chair_notes_thread.finished.connect(
        lambda success, count, files, msg: on_chair_notes_finished(
            window, success, count, files, msg
        )
    )
    window.chair_notes_thread.start()


def on_chair_notes_finished(window, success: bool, count: int, files: list, msg: str):
    """Handles completion of Chairman's Notes download and provides direct open actions."""
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if success and count > 0:
        window.refresh_btn.setText(f"✅ {count} Notes")
        QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))
        chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"

        box = QMessageBox(window)
        box.setWindowTitle("Chairman's Notes Downloaded")
        box.setText(f"Successfully downloaded {count} Chairman's Notes file(s) to:\n{chair_notes_dir}")

        if files:
            file_sample = "\n".join(f"• {f}" for f in files[:8])
            if len(files) > 8:
                file_sample += f"\n...and {len(files) - 8} more"
            box.setInformativeText(file_sample)

        btn_word = box.addButton("📝 Open in Word", QMessageBox.ActionRole)
        btn_folder = box.addButton("📂 Open Folder", QMessageBox.ActionRole)
        btn_close = box.addButton("Close", QMessageBox.RejectRole)
        box.setDefaultButton(btn_word)
        box.exec_()

        clicked = box.clickedButton()
        if clicked == btn_word:
            if files:
                doc_files = [f for f in files if Path(f).suffix.lower() in (".docx", ".doc")]
                target_file = chair_notes_dir / (doc_files[0] if doc_files else files[0])
                _open_file_in_word(target_file, window)
        elif clicked == btn_folder:
            _open_folder(chair_notes_dir)

    elif success and count == 0:
        QMessageBox.information(window, "Chairman's Notes", msg)
    else:
        QMessageBox.warning(window, "Download Failed", msg)


def _open_file_in_word(filepath: Path, parent=None):
    """Opens a document in Microsoft Word using the system default handler."""
    import os
    try:
        if hasattr(os, "startfile"):
            os.startfile(str(filepath))
        else:
            import webbrowser
            webbrowser.open(f"file:///{filepath}")
    except Exception as e:
        if parent:
            QMessageBox.warning(parent, "Open Failed", f"Could not open document:\n{e}")


def open_chair_notes_folder(window):
    """Opens the local folder containing Chairman's Notes."""
    _open_folder(window.meeting_dir / "Agenda" / "ChairNotes")


def scan_existing_chair_notes(window):
    """
    Scans existing Chairman's Notes files (.docx, .doc, .zip) on disk,
    extracts current agreements and TDoc remarks, merges them into the table,
    and reports results to the user.
    """
    if not getattr(window, "is_ran", False):
        QMessageBox.information(
            window,
            "RAN Feature Only",
            "Scanning Chairman's Notes for inline text agreements is designed for RAN working groups (RAN1, RAN2, RAN3, etc.).\n\n"
            f"Meeting '{getattr(window, 'mtg_info', {}).get('wg_name', '')}' follows the SA/CT operational model where agreements "
            "are approved directly as TDoc status in the official meeting table or TdocsByAgenda.htm.",
        )
        return

    from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser

    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText("⏳ Scanning Notes...")

    try:
        agenda_data, agreements, files_scanned = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
            window.meeting_dir,
            mtg_info=getattr(window, "mtg_info", None),
            agenda_map=getattr(window, "agenda_map", None),
            force_rescan=True,
        )

        non_recap_agreements = [a for a in (agreements or []) if not a.get("is_previous_meeting_recap")]
        ag_count = len(non_recap_agreements)
        tdoc_count = len(agenda_data)

        if agreements:
            window.chair_agreements = agreements
            window.model.set_chair_agreements(agreements)

        if agenda_data:
            window.model.merge_agenda_data(agenda_data)
            window._refresh_comboboxes()

        if hasattr(window, "_update_agreements_counter"):
            window._update_agreements_counter()

        window.refresh_btn.setEnabled(True)
        window.refresh_btn.setText("🔄 Refresh")

        if files_scanned > 0:
            window.refresh_btn.setText(f"✅ {ag_count} Agreements")
            QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))

            wa_cnt = sum(1 for a in non_recap_agreements if "working assumption" in str(a.get("kind", "")).lower())
            ffs_cnt = sum(1 for a in non_recap_agreements if "ffs" in str(a.get("kind", "")).lower() or "open issue" in str(a.get("kind", "")).lower())
            pure_ag = ag_count - wa_cnt - ffs_cnt

            ais_covered = len({str(a.get("ai", "")).strip() for a in non_recap_agreements if a.get("ai")})
            wis_covered = len({str(a.get("wi", "")).strip() for a in non_recap_agreements if a.get("wi") and a.get("wi") != "General"})

            msg = (
                f"Successfully scanned {files_scanned} Chairman's Notes file(s) in:\n"
                f"{window.meeting_dir / 'Agenda' / 'ChairNotes'}\n\n"
                f"• Agreements & Decisions Extracted: {ag_count}\n"
                f"  - 🟢 Technical Agreements: {pure_ag}\n"
                f"  - 🟡 Working Assumptions: {wa_cnt}\n"
                f"  - 🔵 FFS / Open Issues: {ffs_cnt}\n"
                f"• Agenda Items Covered: {ais_covered}\n"
                f"• Work Items Covered: {wis_covered}\n"
                f"• Meeting Table Rows Updated: {tdoc_count}\n\n"
                f"All agreements are cached and ready for post-meeting and Markdown reports."
            )
            QMessageBox.information(window, "Chairman's Notes Scanned", msg)
        else:
            chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"
            msg = (
                f"No Chairman's Notes files (.docx, .doc, .zip) were found in:\n{chair_notes_dir}\n\n"
                "Would you like to check the 3GPP server and download Chairman's Notes now?"
            )

            reply = QMessageBox.question(
                window,
                "No Local Chairman's Notes Found",
                msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                download_chair_notes(window)
    except Exception as e:
        window.refresh_btn.setEnabled(True)
        window.refresh_btn.setText("🔄 Refresh")
        logger.error(f"[scan_existing_chair_notes] Error: {e}", exc_info=True)
        QMessageBox.warning(window, "Scan Failed", f"Failed to scan Chairman's Notes:\n{e}")


def start_window_tdocs_caching(window):
    """Initiates asynchronous parallel caching of all meeting TDocs from Docs/."""
    from modules.meetings.core.tdocs.tdocs_cacher import TDocsCacherThread

    if getattr(window, "cacher_thread", None) and window.cacher_thread.isRunning():
        QMessageBox.information(
            window,
            "Caching in Progress",
            "TDocs caching is already running in the background. Please wait for it to complete.",
        )
        return

    docs_url = getattr(window, "docs_ftp_url", "")
    if not docs_url and getattr(window, "main_ftp_url", ""):
        docs_url = f"{window.main_ftp_url.rstrip('/')}/Docs/"

    if not docs_url:
        QMessageBox.warning(
            window,
            "No Docs Folder URL",
            "Cannot cache TDocs: No valid Docs folder URL found for this meeting.",
        )
        return

    window.refresh_btn.setEnabled(False)
    window.refresh_btn.setText("⏳ Caching TDocs...")

    parent_obj = window if hasattr(window, "metaObject") else None
    window.cacher_thread = TDocsCacherThread(docs_url, window.meeting_dir, parent=parent_obj)
    window.cacher_thread.progress.connect(
        lambda msg: window.refresh_btn.setText(f"⏳ {msg}"[:30])
    )
    window.cacher_thread.finished.connect(
        lambda success, msg: on_window_tdocs_caching_finished(window, success, msg)
    )
    window.cacher_thread.start()


def on_window_tdocs_caching_finished(window, success: bool, msg: str):
    """Resets UI and displays the result of the meeting TDocs caching process."""
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if success:
        QMessageBox.information(window, "Caching Complete", msg)
    else:
        QMessageBox.warning(window, "Caching Failed", msg)
