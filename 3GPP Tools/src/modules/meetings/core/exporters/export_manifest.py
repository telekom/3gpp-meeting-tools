# --- File: src/modules/meetings/core/exporters/export_manifest.py ---
"""
Backward-compatible re-export shim for Persistent Export Manifest & Changelog Tracker.

Decomposed into the modular package `modules.meetings.core.exporters.manifest`.
"""

import logging

from modules.meetings.core.exporters.manifest import (
    MANIFEST_FILENAME,
    MANIFEST_VERSION,
    ExportManifest,
    IncrementalPlan,
    _clean_for_json,
    _json_fallback,
    compute_meeting_snapshot,
    diff_meeting_snapshots,
    extract_base_tdoc_id,
    format_local_wi_changelog_markdown,
    format_master_changelog_markdown,
    get_manifest_path,
    is_export_complete,
    load_manifest,
    load_or_reconstruct_manifest,
    merge_manifest_into_tree,
    plan_incremental_session,
    reconstruct_manifest_from_directory,
    sanitize_folder_name,
    save_manifest,
)

logger = logging.getLogger(__name__)

__all__ = [
    "MANIFEST_FILENAME",
    "MANIFEST_VERSION",
    "_clean_for_json",
    "_json_fallback",
    "sanitize_folder_name",
    "is_export_complete",
    "IncrementalPlan",
    "ExportManifest",
    "extract_base_tdoc_id",
    "compute_meeting_snapshot",
    "diff_meeting_snapshots",
    "get_manifest_path",
    "load_manifest",
    "save_manifest",
    "reconstruct_manifest_from_directory",
    "load_or_reconstruct_manifest",
    "plan_incremental_session",
    "merge_manifest_into_tree",
    "format_master_changelog_markdown",
    "format_local_wi_changelog_markdown",
]
