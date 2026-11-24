# Updater UI Package (`core.ui.updater_ui`)

Acyclic modular package extracted from `core/ui/update_dialog.py` (thin backward-compatible re-export shim):

- `workers.py`: `AppUpdateCheckThread` and `AppUpdateDownloadThread`.
- `html_renderer.py`: `_format_commit_date()`, `_format_local_baseline()`, and `_render_commits_html()`.
- `mixins/ui_setup.py`: Layout and design-token QSS styling (`_setup_ui`, `_apply_styles`).
- `mixins/actions.py`: Update checking, commit feed rendering, ZIP staging, Git pull, and detached relaunch helpers.
- `dialog.py`: `AppUpdateDialog` composing the setup and actions mixins cleanly.
