# Spec Search UI Dialogs (`modules/spec_search/ui/dialogs/`)

Modular decomposition of the three ingestion/maintenance dialogs previously housed in the
1,315-line `spec_search_dialogs.py`. `spec_search_dialogs.py` remains a thin,
100%-compatible re-export shim preserving `SpecSearchVersionSelectDialog`,
`SpecSearchPruneDialog`, `SpecSmartImportDialog`, and legacy `SpecClauseDiffDialog`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`version_select_dialog.py`](version_select_dialog.py) | Master-detail release selection & batch ingestion dialog | `SpecSearchVersionSelectDialog` |
| [`prune_dialog.py`](prune_dialog.py) | Superseded minor release review & pruning dialog | `SpecSearchPruneDialog` |
| [`smart_import_ui.py`](smart_import_ui.py) | Search bar, quick chips, delta table & footer UI mixin | `_SmartImportUiMixin` |
| [`smart_import_dialog.py`](smart_import_dialog.py) | Smart single-spec release delta detection & selection | `SpecSmartImportDialog` |
| [`__init__.py`](__init__.py) | Package exports & legacy `SpecClauseDiffDialog` re-export | `__all__` |

## Invariants
- **Zero Raw Hex Literals:** Every stylesheet and color uses `core.ui.theme.C` tokens, guarded by `tests/test_ui_hex_lint.py`.
- Keep every file under **500 lines**.

The version picker loads catalogue queries in the Qt thread pool and displays the latest result through a read-only table model. Each release selection shares one lazy directory index across its cache lookups. A new selection sees newly downloaded files. Imported-version badges use a snapshot loaded on the first selection.
