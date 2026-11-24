"""Modular Chairman's Notes selection, transfer, and worker package."""

from .selector import (
    _PURE_TDOC_ZIP_RE,
    _VALID_NOTES_EXTENSIONS,
    SUBPATHS,
    extract_timestamp_sort_key,
    group_and_pick_latest_by_track,
    is_valid_chair_notes_filename,
    looks_like_directory_listing,
    parse_file_list,
    pick_latest_file,
    prefer_clean_variants,
    select_best_chair_notes_files,
)
from .threads import ChairNotesDownloaderThread, ChairNotesListingThread
from .transfer import (
    download_files_sync,
    extract_chair_notes_zip,
    format_byte_size,
    locate_chair_notes_dir,
)

__all__ = [
    "ChairNotesDownloaderThread",
    "ChairNotesListingThread",
    "SUBPATHS",
    "_PURE_TDOC_ZIP_RE",
    "_VALID_NOTES_EXTENSIONS",
    "download_files_sync",
    "extract_chair_notes_zip",
    "extract_timestamp_sort_key",
    "format_byte_size",
    "group_and_pick_latest_by_track",
    "is_valid_chair_notes_filename",
    "locate_chair_notes_dir",
    "looks_like_directory_listing",
    "parse_file_list",
    "pick_latest_file",
    "prefer_clean_variants",
    "select_best_chair_notes_files",
]
