# Workspace Window Session Restore (`core/ui/window_restore`)

Acyclic modular breakdown behind the backward-compatible `core.ui.window_session` shim:

- **`preferences.py`**: Restoration mode preferences (`always`, `restart_only`, `off`), restart-originated flag handling, `--no-restore` / `Shift` escape hatches, Qt geometry base64 serialization + screen clamping, and QSettings persistence.
- **`collector.py`**: Lazy workspace window registry (`_resolve_supported_window_types`) and ordered visibility snapshotting (`collect_open_windows_state`).
- **`restorers.py`**: Per-window restoration handlers (`_RESTORERS`) for `TDocsWindow`, `SessionPlanWindow`, email windows, studios, spec history/excerpt/diff dialogs, `CrPcrDialog`, and `restore_open_windows`.
