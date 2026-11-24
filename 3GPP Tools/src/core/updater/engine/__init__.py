# --- File: src/core/updater/engine/__init__.py ---
"""Modular Continuous GitHub `master` self-updater engine package."""

from .atom_feed import (
    _parse_iso_timestamp,
    _strip_html_content,
    fetch_remote_commits,
    parse_commits_atom,
)
from .git_state import (
    _estimate_source_mtime,
    _find_git_root,
    _git_is_ancestor,
    _inspect_git_state,
    _parse_porcelain_rel_paths,
    _read_git_head_sha_from_files,
    inspect_local_install,
)
from .models import (
    _ATOM_NS,
    _SHA_RE,
    _TAG_RE,
    DEFAULT_CHECK_INTERVAL_HOURS,
    CommitInfo,
    LocalInstallInfo,
    StagedUpdateInfo,
    UpdateCheckResult,
    get_app_root,
)
from .staging_and_apply import (
    _locate_staged_app_dir,
    _safe_extract_zip,
    _sha256_file,
    apply_git_update,
    launch_staged_update_and_relaunch,
    spawn_detached_relaunch,
    stage_zip_update,
)
from .status_and_state import (
    check_for_updates,
    evaluate_update_status,
    get_update_state_path,
    load_update_state,
    save_update_state,
    should_auto_check,
)

__all__ = [
    "_ATOM_NS",
    "_SHA_RE",
    "_TAG_RE",
    "DEFAULT_CHECK_INTERVAL_HOURS",
    "CommitInfo",
    "LocalInstallInfo",
    "UpdateCheckResult",
    "StagedUpdateInfo",
    "get_app_root",
    "_parse_iso_timestamp",
    "_strip_html_content",
    "parse_commits_atom",
    "fetch_remote_commits",
    "_read_git_head_sha_from_files",
    "_find_git_root",
    "_parse_porcelain_rel_paths",
    "_inspect_git_state",
    "_git_is_ancestor",
    "_estimate_source_mtime",
    "inspect_local_install",
    "evaluate_update_status",
    "get_update_state_path",
    "load_update_state",
    "save_update_state",
    "should_auto_check",
    "check_for_updates",
    "_safe_extract_zip",
    "_locate_staged_app_dir",
    "_sha256_file",
    "stage_zip_update",
    "launch_staged_update_and_relaunch",
    "apply_git_update",
    "spawn_detached_relaunch",
]
