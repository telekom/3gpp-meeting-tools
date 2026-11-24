# Email Manager Window Package (`modules.emails.ui.manager_window`)

Modular mixin package extracted from `ui/email_window.py` (thin backward-compatible re-export shim):

- `window.py`: `EmailManagerWindow` composing mixins, constructor, `closeEvent`, and session persistence.
- `mixins/ui_setup.py`: Config persistence and split-pane Qt layout setup (`_setup_ui`, `_configure_folders`).
- `mixins/filtering_and_selection.py`: SQLite reloads, dropdown filters, selection changes, reader HTML rendering, and context menus.
- `mixins/actions_and_workers.py`: Star/follow actions, Outlook file launching, and background sync/move/rescan/stats workers.
