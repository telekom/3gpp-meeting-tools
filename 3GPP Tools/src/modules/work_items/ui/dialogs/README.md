# Work Items UI Dialogs (`modules/work_items/ui/dialogs/`)

Modular decomposition of the dialogs and Quick Copy builders previously housed in the
1,086-line `work_items_dialogs.py`. `work_items_dialogs.py` remains a thin,
100%-compatible re-export shim preserving all 8 public symbols and `logger`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`copy_menus.py`](copy_menus.py) | Plain/Markdown/Word-HTML clipboard builders and context menus | `build_wi_copy_data`, `populate_wi_copy_menu`, `populate_wi_batch_copy_menu`, `copy_wis_default` |
| [`info_dialog.py`](info_dialog.py) | Work Item metadata inspector, impacted specs, remarks feed & star/follow toggles | `WorkItemInfoDialog` |
| [`linked_specs_dialog.py`](linked_specs_dialog.py) | Scrollable impacted specifications list with DynaReport & FTP buttons | `LinkedSpecsDialog` |
| [`add_dialog.py`](add_dialog.py) | Background portal fetcher thread & manual Add/Fetch Work Item dialog | `ManualWIFetcherThread`, `AddWorkItemDialog` |
| [`__init__.py`](__init__.py) | Package exports (all 8 public symbols in `__all__` + `logger`) | `WorkItemInfoDialog`, `AddWorkItemDialog`, `LinkedSpecsDialog` |

## Invariants
- **Zero Raw Hex Literals:** Every stylesheet and color uses `core.ui.theme.C` tokens, guarded by `tests/test_ui_hex_lint.py`.
- Keep every file under **350 lines**.
