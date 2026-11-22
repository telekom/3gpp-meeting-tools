# Agreement Tracker Dialog (`modules/meetings/ui/agreement_tracker/`)

Modular decomposition of `AgreementTrackerDialog` and `AgreementResultsTableModel`
previously housed in the 1,381-line `agreement_tracker_dialog.py`. `agreement_tracker_dialog.py`
remains a thin, 100%-compatible re-export shim preserving `AgreementTrackerDialog`,
`AgreementResultsTableModel`, and `logger`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`model.py`](model.py) | Read-only results table model & live keyword/type filtering | `AgreementResultsTableModel` |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Splitter layout, filter ribbon, KPI cards, table & detail drawer UI | `_TrackerUiMixin` |
| [`mixins/filters_and_search.py`](mixins/filters_and_search.py) | WG/date/WI token helpers, async worker search & KPI updates | `_TrackerSearchMixin` |
| [`mixins/details_and_actions.py`](mixins/details_and_actions.py) | HTML detail rendering, TDoc archive download/open & context menus | `_TrackerActionsMixin` |
| [`mixins/export_report.py`](mixins/export_report.py) | HTML/Markdown agreement dossier compilation & Excel workbook export | `_TrackerExportMixin` |
| [`dialog.py`](dialog.py) | Modeless top-level `QDialog` shell (`Qt.Window`) & session state | `AgreementTrackerDialog`, `logger` |
| [`__init__.py`](__init__.py) | Package exports | `AgreementTrackerDialog`, `AgreementResultsTableModel` |

## Invariants
- **Modeless Top-Level Behavior:** `AgreementTrackerDialog` initializes with `super().__init__(None, Qt.Window | ...)` and `setModal(False)` to prevent Win32 `GW_OWNER` pinning.
- **Zero Raw Hex Literals:** Every stylesheet and generated HTML color uses `core.ui.theme.C`, enforced by `tests/test_ui_hex_lint.py`.
- Keep every file under **400 lines**.
