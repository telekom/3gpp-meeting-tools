# --- File: src/core/updater/engine/status_and_state.py ---
"""Update evaluation, state persistence, throttling, and update check orchestration."""

from __future__ import annotations

import json
import logging
import os
import time
from pathlib import Path
from typing import Dict, List, Optional

from core.updater.apply_worker import UPDATE_STATE_FILENAME
from core.utils.paths import get_app_data_root

from .atom_feed import fetch_remote_commits
from .git_state import _git_is_ancestor, inspect_local_install
from .models import (
    DEFAULT_CHECK_INTERVAL_HOURS,
    CommitInfo,
    LocalInstallInfo,
    UpdateCheckResult,
    _facade,
)

logger = logging.getLogger(__name__)


def evaluate_update_status(
    local_info: LocalInstallInfo,
    commits: List[CommitInfo],
) -> UpdateCheckResult:
    """Compare local installation baseline against parsed remote commits."""
    if not commits:
        return UpdateCheckResult(
            status="error",
            local_info=local_info,
            error="No commits were found in the remote GitHub feed.",
        )

    latest = commits[0]

    if local_info.installed_sha:
        is_ancestor_fn = _facade("_git_is_ancestor", _git_is_ancestor)
        if latest.sha == local_info.installed_sha or (
            local_info.install_mode == "git"
            and is_ancestor_fn(local_info.git_root, latest.sha)
        ):
            return UpdateCheckResult(
                status="up_to_date",
                local_info=local_info,
                latest_commit=latest,
                pending_commits=[],
                all_commits=commits,
            )

        pending: List[CommitInfo] = []
        found_local = False
        for commit in commits:
            if commit.sha == local_info.installed_sha:
                found_local = True
                break
            pending.append(commit)

        return UpdateCheckResult(
            status="update_available",
            local_info=local_info,
            latest_commit=latest,
            pending_commits=pending,
            all_commits=commits,
            history_truncated=(not found_local and len(pending) == len(commits)),
        )

    # Manual ZIP install without `.app_version.json`: compare commit timestamps
    # with a 120-second tolerance so a freshly downloaded ZIP isn't flagged.
    cutoff = (local_info.estimated_ts or 0.0) + 120.0
    pending = [c for c in commits if c.updated_ts > cutoff]
    if pending:
        return UpdateCheckResult(
            status="update_available",
            local_info=local_info,
            latest_commit=latest,
            pending_commits=pending,
            all_commits=commits,
            history_truncated=(len(pending) == len(commits)),
        )

    return UpdateCheckResult(
        status="up_to_date",
        local_info=local_info,
        latest_commit=latest,
        pending_commits=[],
        all_commits=commits,
    )


def get_update_state_path() -> Path:
    """Return the path to persisted updater preferences in user app data."""
    return get_app_data_root() / UPDATE_STATE_FILENAME


def load_update_state() -> Dict[str, object]:
    """Load persisted updater settings and throttling timestamps."""
    path = get_update_state_path()
    defaults: Dict[str, object] = {
        "auto_check_enabled": False,
        "check_interval_hours": DEFAULT_CHECK_INTERVAL_HOURS,
        "last_check_ts": 0.0,
        "skipped_sha": "",
    }
    if not path.is_file():
        return defaults
    try:
        loaded = json.loads(path.read_text(encoding="utf-8")) or {}
        if isinstance(loaded, dict):
            defaults.update(loaded)
    except Exception as exc:
        logger.warning("[Updater] Failed to read update state: %s", exc)
    return defaults


def save_update_state(updates: Dict[str, object]) -> Dict[str, object]:
    """Merge ``updates`` into persisted updater state and save atomically."""
    state = load_update_state()
    state.update(updates)
    path = get_update_state_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".json.tmp")
    tmp_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
    os.replace(str(tmp_path), str(path))
    return state


def should_auto_check(state: Dict[str, object], now_ts: Optional[float] = None) -> bool:
    """Return True when background startup update check is enabled and due."""
    if not bool(state.get("auto_check_enabled", True)):
        return False
    now = time.time() if now_ts is None else float(now_ts)
    last_ts = float(state.get("last_check_ts") or 0.0)
    interval_h = max(1.0, float(state.get("check_interval_hours") or DEFAULT_CHECK_INTERVAL_HOURS))
    return (now - last_ts) >= (interval_h * 3600.0)


def check_for_updates(
    app_root: Optional[Path] = None,
    timeout: int = 15,
    record_check_time: bool = True,
) -> UpdateCheckResult:
    """Check GitHub `master` for new commits and compare against the local install."""
    local_info = inspect_local_install(app_root=app_root)
    try:
        commits = fetch_remote_commits(timeout=timeout)
    except Exception as exc:
        return UpdateCheckResult(
            status="error",
            local_info=local_info,
            error=str(exc),
        )

    if record_check_time:
        try:
            save_update_state({"last_check_ts": time.time()})
        except Exception:
            pass

    return evaluate_update_status(local_info, commits)
