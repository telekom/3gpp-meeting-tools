# --- File: src/modules/meetings/ui/tdocs_agenda_controller.py ---
"""Backward-compatible facade re-exporting `modules.meetings.ui.tdocs_agenda`."""

import logging

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMessageBox

from modules.meetings.ui.tdocs_agenda import (
    _ensure_refresh_btn_unlocked,
    _next_refresh_seq,
    _open_file_in_word,
    _restore_refresh_btn,
    _set_refresh_btn_busy,
    _update_refresh_btn_progress,
    download_agenda_csv,
    download_chair_notes,
    finalize_import_batch,
    import_word_agenda_dialog,
    on_agenda_download_finished,
    on_chair_notes_finished,
    on_chair_notes_listing_finished,
    on_window_tdocs_caching_finished,
    on_word_import_finished,
    on_word_import_progress,
    open_chair_notes_folder,
    process_imported_agenda_files,
    process_next_in_import_queue,
    scan_existing_chair_notes,
    start_chair_notes_download,
    start_window_tdocs_caching,
)

logger = logging.getLogger(__name__)

__all__ = [
    "QMessageBox",
    "QTimer",
    "_ensure_refresh_btn_unlocked",
    "_next_refresh_seq",
    "_open_file_in_word",
    "_restore_refresh_btn",
    "_set_refresh_btn_busy",
    "_update_refresh_btn_progress",
    "download_agenda_csv",
    "download_chair_notes",
    "finalize_import_batch",
    "import_word_agenda_dialog",
    "logger",
    "on_agenda_download_finished",
    "on_chair_notes_finished",
    "on_chair_notes_listing_finished",
    "on_window_tdocs_caching_finished",
    "on_word_import_finished",
    "on_word_import_progress",
    "open_chair_notes_folder",
    "process_imported_agenda_files",
    "process_next_in_import_queue",
    "scan_existing_chair_notes",
    "start_chair_notes_download",
    "start_window_tdocs_caching",
]
