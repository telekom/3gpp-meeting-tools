# --- File: src/modules/meetings/ui/tdocs_agenda_controller.py ---
"""
Agenda and Chairman's Notes import/download controller for TDocsWindow.
Manages batch Word document agenda importing, CSV agenda synchronization,
and Chairman's Notes retrieval threads.
"""
import re
from pathlib import Path

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from core.network.network_state import NetworkState
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.agenda_manager import AgendaDownloaderThread, AgendaManager
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.tdocs_threads import WordAgendaImporterThread
from modules.meetings.core.url_router import URLRouter


def import_word_agenda_dialog(window):
    """Prompts file dialog to select agenda / chair notes Word or HTML files to import."""
    agenda_dir = window.meeting_dir / "Agenda"
    start_dir = str(agenda_dir) if agenda_dir.exists() else str(window.meeting_dir)

    file_paths, _ = QFileDialog.getOpenFileNames(
        window,
        "Select TDocsByAgenda / Chairman's Notes File(s)",
        start_dir,
        "Supported Files (*.docx *.doc *.htm *.html);;Word Documents (*.docx *.doc);;DOCX Files (*.docx);;DOC Files (*.doc);;HTML Files (*.htm *.html);;All Files (*.*)",
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

    window.word_import_thread = WordAgendaImporterThread(current_file, window.meeting_dir)
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
    if success and agenda_data:
        window.model.merge_agenda_data(agenda_data)
        window._import_accumulated_merged += len(agenda_data)
        window._import_success_files.append((filename, len(agenda_data)))
    else:
        window._import_failed_files.append((filename, error_msg or "No valid TDocs table found"))

    process_next_in_import_queue(window)


def finalize_import_batch(window):
    """Resets UI state and shows completion summary for agenda batch ingestion."""
    window._is_importing_agenda = False
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if window._import_accumulated_merged > 0:
        window._refresh_comboboxes()
        window.refresh_btn.setText(f"✅ {window._import_accumulated_merged} Merged")
        QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))

    if window._total_import_batch_count == 1:
        if window._import_success_files:
            fname, count = window._import_success_files[0]
            QMessageBox.information(
                window,
                "Import Successful",
                f"Saved to Agenda/\nSuccessfully merged {count} TDocs from:\n{fname}",
            )
        else:
            fname, err = window._import_failed_files[0]
            QMessageBox.warning(window, "Import Failed", f"Failed to import {fname}:\n{err}")
        return

    summary_lines = [
        f"Batch import completed for {window._total_import_batch_count} files.",
        f"• Successfully imported: {len(window._import_success_files)} file(s)",
        f"• Total TDocs merged: {window._import_accumulated_merged}\n",
    ]

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
    """Initiates asynchronous download of Chairman's Notes files from 3GPP servers."""
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
    window.refresh_btn.setText("⏳ Downloading Notes...")

    chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"
    window.chair_notes_thread = ChairNotesDownloaderThread(candidate_urls, chair_notes_dir)
    window.chair_notes_thread.progress.connect(lambda msg: window.refresh_btn.setText(f"⏳ {msg}"[:30]))
    window.chair_notes_thread.finished.connect(
        lambda success, count, files, msg: on_chair_notes_finished(window, success, count, files, msg)
    )
    window.chair_notes_thread.start()


def on_chair_notes_finished(window, success: bool, count: int, files: list, msg: str):
    """Handles completion of Chairman's Notes download and offers instant ingestion."""
    window.refresh_btn.setEnabled(True)
    window.refresh_btn.setText("🔄 Refresh")

    if success and count > 0:
        window.refresh_btn.setText(f"✅ {count} Notes")
        QTimer.singleShot(4000, lambda: window.refresh_btn.setText("🔄 Refresh"))
        chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"
        reply = QMessageBox.question(
            window,
            "Chairman's Notes Downloaded",
            f"{msg}\n\nWould you like to parse and merge these {count} Chairman's Notes file(s) into your meeting table now?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            source_paths = [chair_notes_dir / f for f in files]
            process_imported_agenda_files(window, source_paths)
    elif success and count == 0:
        QMessageBox.information(window, "Chairman's Notes", msg)
    else:
        QMessageBox.warning(window, "Download Failed", msg)


def open_chair_notes_folder(window):
    """Opens the local folder containing Chairman's Notes."""
    _open_folder(window.meeting_dir / "Agenda" / "ChairNotes")
