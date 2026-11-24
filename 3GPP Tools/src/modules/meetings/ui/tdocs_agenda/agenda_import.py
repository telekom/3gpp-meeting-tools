"""Batch Word/ZIP/HTML agenda ingestion and official Agenda CSV download actions."""

import re
from pathlib import Path

from PyQt5.QtWidgets import QFileDialog

from core.network.network_state import NetworkState
from modules.meetings.core.agenda_manager import AgendaDownloaderThread, AgendaManager
from modules.meetings.core.tdocs.tdocs_threads import WordAgendaImporterThread
from modules.meetings.core.url_router import URLRouter

from .refresh_state import (
    _qmessagebox,
    _restore_refresh_btn,
    _set_refresh_btn_busy,
    _update_refresh_btn_progress,
)


def import_word_agenda_dialog(window):
    """Prompt file dialog to select agenda / chair notes Word, ZIP, or HTML files to import."""
    agenda_dir = window.meeting_dir / "Agenda"
    start_dir = str(agenda_dir) if agenda_dir.exists() else str(window.meeting_dir)

    file_paths, _ = QFileDialog.getOpenFileNames(
        window,
        "Select TDocsByAgenda / Chairman's Notes File(s)",
        start_dir,
        "Supported Files (*.docx *.doc *.zip *.htm *.html);;"
        "Word Documents (*.docx *.doc);;ZIP Archives (*.zip);;HTML Files (*.htm *.html);;All Files (*.*)",
    )
    if file_paths:
        process_imported_agenda_files(window, [Path(p) for p in file_paths])


def process_imported_agenda_files(window, source_paths: list):
    """Queue files for sequential background Word agenda ingestion."""
    valid_paths = [p for p in source_paths if p.exists()]
    if not valid_paths:
        _qmessagebox().warning(window, "Files Not Found", "None of the specified files exist on disk.")
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
    """Process the next file in the agenda import batch queue."""
    if not window._import_queue:
        finalize_import_batch(window)
        return

    current_file = window._import_queue.pop(0)
    current_idx = window._total_import_batch_count - len(window._import_queue)

    _set_refresh_btn_busy(
        window,
        f"({current_idx}/{window._total_import_batch_count}) Importing...",
    )

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
    """Update button text with import progress status."""
    _update_refresh_btn_progress(window, msg, max_len=34)


def on_word_import_finished(window, success: bool, agenda_data: dict, filename: str, error_msg: str):
    """Handle parsing result for a single agenda file and proceed to next in queue."""
    extracted_ags = (
        getattr(window.word_import_thread, "extracted_agreements", [])
        if hasattr(window, "word_import_thread")
        else []
    )
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
    """Reset UI state and show completion summary for agenda batch ingestion."""
    window._is_importing_agenda = False
    mb = _qmessagebox()

    ag_total = getattr(window, "_import_accumulated_agreements", 0)
    if hasattr(window, "_update_agreements_counter"):
        window._update_agreements_counter()
    if window._import_accumulated_merged > 0 or ag_total > 0:
        window._refresh_comboboxes()
        _restore_refresh_btn(
            window,
            f"✅ {window._import_accumulated_merged} Merged",
            reset_after_ms=4000,
        )
    else:
        _restore_refresh_btn(window)

    if window._total_import_batch_count == 1:
        if window._import_success_files:
            entry = window._import_success_files[0]
            fname = entry[0]
            count = entry[1]
            ag_cnt = entry[2] if len(entry) > 2 else 0
            ag_line = f"\n• Extracted {ag_cnt} structured agreement(s) / working assumption(s)" if ag_cnt > 0 else ""
            mb.information(
                window,
                "Import Successful",
                f"Saved to Agenda/\nSuccessfully merged {count} TDoc entries{ag_line} from:\n{fname}",
            )
        else:
            fname, err = window._import_failed_files[0]
            mb.warning(window, "Import Failed", f"Failed to import {fname}:\n{err}")
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
        mb.warning(window, "Batch Import Completed with Warnings", "\n".join(summary_lines))
    else:
        mb.information(window, "Batch Import Successful", "\n".join(summary_lines))


def download_agenda_csv(window):
    """Initiate asynchronous download of the official 3GPP meeting agenda CSV file."""
    wg_name = window.mtg_info.get("wg_name", "").upper()
    main_url = window.mtg_info.get("url_key", "")
    if main_url and not main_url.startswith("http"):
        main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

    is_active = window.mtg_info.get("is_active_sync", False)
    candidate_urls = []
    if NetworkState.get_instance().is_local_active():
        candidate_urls.append(URLRouter._get_local_server_base(wg_name))

    if is_active:
        sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
        candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

    if main_url:
        candidate_urls.append(main_url)

    docs_url = window.mtg_info.get("docs_folder_url", "")
    if docs_url:
        if not docs_url.startswith("http"):
            docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip("/")
        base_from_docs = re.sub(r"/Docs/?$", "", docs_url, flags=re.IGNORECASE)
        if base_from_docs not in candidate_urls:
            candidate_urls.append(base_from_docs)

    if not candidate_urls:
        return _qmessagebox().warning(
            window, "No URL Available", "Cannot download Agenda CSV: No valid meeting URL found."
        )

    _set_refresh_btn_busy(window, "Getting Agenda...")

    window.agenda_dl_thread = AgendaDownloaderThread(candidate_urls, window.agenda_dir, window)
    window.agenda_dl_thread.progress.connect(lambda msg: _update_refresh_btn_progress(window, msg))
    window.agenda_dl_thread.finished.connect(
        lambda success, msg: on_agenda_download_finished(window, success, msg)
    )
    window.agenda_dl_thread.start()


def on_agenda_download_finished(window, success: bool, msg: str):
    """Handle completion of the agenda CSV download."""
    mb = _qmessagebox()
    if success:
        _restore_refresh_btn(window, "✅ Agenda Cached", reset_after_ms=4000)
        window.agenda_map = AgendaManager.load_local_agenda(window.agenda_dir)
        if hasattr(window, "model") and window.model:
            window.model.set_agenda_map(window.agenda_map)
        window._refresh_comboboxes()
        mb.information(window, "Agenda CSV Cached", msg)
    else:
        _restore_refresh_btn(window)
        mb.warning(window, "Download Failed", msg)
