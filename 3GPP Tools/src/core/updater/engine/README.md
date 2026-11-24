# Updater Engine Package (`core/updater/engine`)

Modular self-updater backend for checking GitHub `master` commits, detecting local Git vs ZIP installations, staging ZipSlip-safe archives, running `git pull --ff-only`, and spawning detached relaunches.

## Module Inventory

| File | Responsibility |
|---|---|
| `models.py` | `CommitInfo`, `LocalInstallInfo`, `UpdateCheckResult`, `StagedUpdateInfo`, `get_app_root()`, `_facade()` patch bridge |
| `atom_feed.py` | `parse_commits_atom`, `fetch_remote_commits`, ISO/HTML feed normalization |
| `git_state.py` | `.git/HEAD` reader, porcelain dirty-tree checker (ignoring runtime seeds), `inspect_local_install` |
| `status_and_state.py` | `evaluate_update_status`, update state persistence, `should_auto_check`, `check_for_updates` |
| `staging_and_apply.py` | ZipSlip-safe `_safe_extract_zip`, `stage_zip_update`, `launch_staged_update_and_relaunch`, `apply_git_update`, `spawn_detached_relaunch` |

`core/updater/app_updater.py` remains as a thin backward-compatible re-export facade preserving all `patch("core.updater.app_updater.*")` test hooks.
