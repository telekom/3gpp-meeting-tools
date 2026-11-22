# Spec Clause Diff & Attribution Dialog (`modules/spec_search/ui/clause_diff/`)

Modular decomposition of `SpecClauseDiffDialog` previously housed in the 1,159-line
`spec_clause_diff_dialog.py`. `spec_clause_diff_dialog.py` remains a thin,
100%-compatible re-export shim preserving `SpecClauseDiffDialog` and `logger`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Visual Diff tab, AI Prompt tab & Word Compare footer controls | `_DiffUiMixin` |
| [`mixins/diff_preview.py`](mixins/diff_preview.py) | Version loading, HTML diff renderer, prompt copy/save & Word Compare launch | `_DiffPreviewMixin` |
| [`mixins/attribution_ui.py`](mixins/attribution_ui.py) | Tab 3 attribution controls & Stage 1 worker lifecycle | `_AttributionUiMixin` |
| [`mixins/attribution_results.py`](mixins/attribution_results.py) | Hero banner, candidate table, Stage 2 deep OpenXML verification & TDoc actions | `_AttributionResultsMixin` |
| [`dialog.py`](dialog.py) | Modeless top-level `QDialog` shell (`Qt.Window`, `Qt.NonModal`) & session state | `SpecClauseDiffDialog`, `logger` |
| [`__init__.py`](__init__.py) | Package exports | `SpecClauseDiffDialog` |

## Invariants
- **Modeless Behavior:** `SpecClauseDiffDialog` initializes with `super().__init__(None)`, `Qt.Window`, `Qt.NonModal`, and `get_session_state()` / `restore_session_state()`.
- **Zero Raw Hex Literals:** All colors reference `core.ui.theme.C`, enforced by `tests/test_ui_hex_lint.py`.
- Keep every file under **350 lines**.
