# `modules.meetings.ui.tdocs_help`

Acyclic package backing the legacy `modules.meetings.ui.tdocs_help_dialog` shim with zero raw hex literals (`core.ui.theme.C`).

## Architecture

- `styles.py` — Shared `HELP_CSS` token stylesheet for help tabs.
- `tabs_html.py` — Rich HTML builders for Toolbar, Table, Search, and Shortcuts tabs.
- `dialog.py` — `TDocsHelpDialog(QDialog)` with `F1`/`Esc` shortcuts and live in-tab search.
- `__init__.py` — Public package exports.
