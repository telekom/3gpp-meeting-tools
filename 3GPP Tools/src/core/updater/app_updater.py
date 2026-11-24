# --- File: src/core/updater/app_updater.py ---
"""
Backward-compatible facade & re-export shim for Continuous GitHub `master` self-updater.

Decomposed into the modular package `core.updater.engine`. Patchable module-level
dependencies (`get_temp_root`, `get_project_root`, `NetworkSession`, `subprocess`,
`_git_is_ancestor`, `spawn_detached_relaunch`) remain bound here so existing
`unittest.mock.patch("core.updater.app_updater.*")` hooks work transparently.
"""

from __future__ import annotations

import logging
import subprocess

from core.network.session import NetworkSession
from core.updater.engine import (
    _ATOM_NS,
    _SHA_RE,
    _TAG_RE,
    DEFAULT_CHECK_INTERVAL_HOURS,
    CommitInfo,
    LocalInstallInfo,
    StagedUpdateInfo,
    UpdateCheckResult,
    _estimate_source_mtime,
    _find_git_root,
    _git_is_ancestor,
    _inspect_git_state,
    _locate_staged_app_dir,
    _parse_iso_timestamp,
    _parse_porcelain_rel_paths,
    _read_git_head_sha_from_files,
    _safe_extract_zip,
    _sha256_file,
    _strip_html_content,
    apply_git_update,
    check_for_updates,
    evaluate_update_status,
    fetch_remote_commits,
    get_app_root,
    get_update_state_path,
    inspect_local_install,
    launch_staged_update_and_relaunch,
    load_update_state,
    parse_commits_atom,
    save_update_state,
    should_auto_check,
    spawn_detached_relaunch,
    stage_zip_update,
)
from core.utils.paths import get_project_root, get_temp_root

logger = logging.getLogger(__name__)

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
    "get_project_root",
    "get_temp_root",
    "NetworkSession",
    "subprocess",
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
