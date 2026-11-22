# --- File: src/modules/meetings/core/exporters/manifest/__init__.py ---
"""Modular Persistent Export Manifest & Changelog Tracker package."""

from .changelog import (
    format_local_wi_changelog_markdown,
    format_master_changelog_markdown,
)
from .models import ExportManifest, IncrementalPlan
from .persistence import (
    get_manifest_path,
    load_manifest,
    load_or_reconstruct_manifest,
    reconstruct_manifest_from_directory,
    save_manifest,
)
from .planning import merge_manifest_into_tree, plan_incremental_session
from .snapshots import (
    compute_meeting_snapshot,
    diff_meeting_snapshots,
    extract_base_tdoc_id,
)
from .validation import (
    MANIFEST_FILENAME,
    MANIFEST_VERSION,
    _clean_for_json,
    _json_fallback,
    is_export_complete,
    sanitize_folder_name,
)

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
