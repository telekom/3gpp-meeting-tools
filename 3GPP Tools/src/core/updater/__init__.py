"""Application self-update services for tracking and applying master updates."""

from core.updater.app_updater import (
    CommitInfo,
    LocalInstallInfo,
    StagedUpdateInfo,
    UpdateCheckResult,
    apply_git_update,
    check_for_updates,
    evaluate_update_status,
    fetch_remote_commits,
    inspect_local_install,
    launch_staged_update_and_relaunch,
    load_update_state,
    parse_commits_atom,
    save_update_state,
    should_auto_check,
    stage_zip_update,
)

__all__ = [
    "CommitInfo",
    "LocalInstallInfo",
    "StagedUpdateInfo",
    "UpdateCheckResult",
    "apply_git_update",
    "check_for_updates",
    "evaluate_update_status",
    "fetch_remote_commits",
    "inspect_local_install",
    "launch_staged_update_and_relaunch",
    "load_update_state",
    "parse_commits_atom",
    "save_update_state",
    "should_auto_check",
    "stage_zip_update",
]
