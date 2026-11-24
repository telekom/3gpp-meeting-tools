# --- File: src/modules/meetings/core/chair_notes_downloader.py ---
"""Backward-compatible facade re-exporting `modules.meetings.core.chair_notes`."""

import time

from modules.meetings.core.chair_notes import (
    _PURE_TDOC_ZIP_RE,
    _VALID_NOTES_EXTENSIONS,
    ChairNotesDownloaderThread,
    ChairNotesListingThread,
)

__all__ = [
    "ChairNotesDownloaderThread",
    "ChairNotesListingThread",
    "_PURE_TDOC_ZIP_RE",
    "_VALID_NOTES_EXTENSIONS",
    "time",
]
