# Specification History Viewer (`modules/specifications/ui/spec_history/`)

Modular decomposition of `SpecVersionHistoryDialog` previously housed in the 1,223-line
`spec_history_dialog.py`. `spec_history_dialog.py` remains a thin, 100%-compatible
re-export shim preserving `SpecVersionHistoryDialog`, `NumericSortItem`, and `logger`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Header cards, version/clause filter bar, table & bottom notes drawer | `_HistoryUiMixin` |
| [`mixins/data_and_table.py`](mixins/data_and_table.py) | Numeric sorting, DB/auto-extract loading, origin CR ranking & table render | `NumericSortItem`, `_HistoryDataMixin` |
| [`mixins/details_and_notes.py`](mixins/details_and_notes.py) | HTML rapporteur notes, relevance banner, TDoc link hover/context menus | `_HistoryDetailsMixin` |
| [`mixins/actions_and_report.py`](mixins/actions_and_report.py) | Row context actions, clipboard TDoc list copy, re-scan & dossier launch | `_HistoryActionsMixin` |
| [`dialog.py`](dialog.py) | Modeless top-level `QDialog` shell (`Qt.Window`, `Qt.NonModal`) & session state | `SpecVersionHistoryDialog`, `logger` |
| [`__init__.py`](__init__.py) | Package exports | `SpecVersionHistoryDialog`, `NumericSortItem` |

## Invariants
- **Modeless Behavior:** `SpecVersionHistoryDialog` initializes with `super().__init__(None)`, `Qt.Window`, `Qt.NonModal`, and `get_session_state()` / `restore_session_state()`.
- **Zero Raw Hex Literals:** All colors reference `core.ui.theme.C`, enforced by `tests/test_ui_hex_lint.py`.
- Keep every file under **350 lines**.
