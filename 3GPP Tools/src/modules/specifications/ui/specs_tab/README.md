# Specifications Catalogue Tab Package (`specs_tab/`)

Modular decomposition of `SpecificationsTab` previously housed in `specifications/ui/ui_tabs.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`tab.py`](tab.py) | `QWidget` entry point, signals, debounce timers, deep-link jump | `SpecificationsTab`, `__init__`, `ensure_loaded`, `select_and_show_spec` |
| [`constants.py`](constants.py) | WG domain scoping, canonical WG ordering & filename part-number regex | `PART_NUM_REGEX`, `USEFUL_*`, `ACTIVE_WG_ORDER`, `LEGACY_*`, `FULL_WG_ORDER` |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | `PageRibbonCard` header, search/filter bar, table geometry, shortcuts, scaling | `SpecsTabUiSetupMixin` (`_setup_ui`, `on_font_scale_changed`, `set_bg_sync_active`) |
| [`mixins/filters.py`](mixins/filters.py) | Persisted JSON preferences, favorites, TSG/WG/Series/Tech/Tag filters, active chips | `SpecsTabFiltersMixin` (`_load_settings`, `_save_settings`, `_populate_wg_combo`, `_update_active_tags`, `_clear_all_filters`) |
| [`mixins/table_render.py`](mixins/table_render.py) | Filtered catalogue table population, WG/tag badges, Column 0 menus, version badges | `SpecsTabTableRenderMixin` (`refresh_table`, `_update_combo_item_badges`) |
| [`mixins/row_documents_cell.py`](mixins/row_documents_cell.py) | Column 2 version combo, disk state refresh, and rich document action menu | `SpecsTabRowDocumentsCellMixin` (`_build_spec_row_documents_cell`) |
| [`mixins/actions.py`](mixins/actions.py) | Context menu, Quick Copy (`Ctrl+C`), sync/history/clause/CR/annotation dialogs, deletions | `SpecsTabActionsMixin` (`_copy_table_selection`, `_open_version_history_dialog`, `_open_clause_excerpt_dialog`, `_open_cr_pcr_dialog`, `_delete_single_spec`) |
| [`__init__.py`](__init__.py) | Package exports | `SpecificationsTab` and WG constants |

## Backward Compatibility

- [`../ui_tabs.py`](../ui_tabs.py) remains as a thin re-export shim exposing `SpecificationsTab`, WG constants, and `get_wg_badge_style_and_desc` so existing imports across [`shell_setup.py`](../../../../core/ui/main_window/mixins/shell_setup.py), tests, and docs remain unaffected.
