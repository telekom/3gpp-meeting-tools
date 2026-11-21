# `work_items_tab` Package

Modular implementation of `WorkItemsTab`, decomposed cleanly from the legacy `src/modules/work_items/ui/ui_tabs.py` monolith.

## Structure & Mixin Composition (MRO)

`WorkItemsTab` inherits from four focused mixins ahead of `QWidget`:

1. `WorkItemsUiSetupMixin` (`mixins/ui_setup.py`): Builds the 2-tier `PageRibbonCard`, filter controls, tracking chips strip, `QTableView` delegates/columns, and font-scale updates (`on_font_scale_changed`).
2. `WorkItemsFiltersTrackingMixin` (`mixins/filters_and_tracking.py`): Filter config persistence (`WorkItemsSettings`), TSG/chip toggles, dismissible tag pills, single-row & batch `starred`/`following` state mutations.
3. `WorkItemsDataSyncMixin` (`mixins/data_and_sync.py`): Incremental SQLite chunk pagination, `ensure_loaded()`, `focus_work_item()` deep linking, and background `WorkItemsScraperThread` / `TargetedWIScraperThread` orchestration.
4. `WorkItemsDialogsMenusMixin` (`mixins/dialogs_and_menus.py`): Row detail popups (`WorkItemInfoDialog`, `LinkedSpecsDialog`, `AddWorkItemDialog`), `Ctrl+C` clipboard copy, and single/batch right-click context menus.

## Backward Compatibility

`src/modules/work_items/ui/ui_tabs.py` acts as a thin re-export shim exposing `WorkItemsTab`, `WorkItemsTableModel`, all 8 table delegates in `__all__`, and `logger`.
