# Spec Search Tab Package (`spec_search_tab/`)

Modular decomposition of `SpecSearchTab` previously housed in `spec_search/ui/spec_search_tabs.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`tab.py`](tab.py) | `QWidget` entry point, signals, debounce timers, spec title/short-name helpers | `SpecSearchTab`, `log_msg`, `navigate_to_specs_tab_requested`, `__init__`, `get_spec_title`, `get_spec_short_name` |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Main splitter, `Indexed Specs` card header, `PageRibbonCard`, evolution matrix `QTabWidget`, `SpecClauseInspector` | `SpecSearchTabUiSetupMixin._setup_ui` |
| [`mixins/search_and_config.py`](mixins/search_and_config.py) | Config persistence, inspector layout toggles, async FTS5 query dispatch, Evolution Matrix population | `SpecSearchTabSearchMixin` (`_save_config`, `_load_config_and_refresh`, `_execute_search`, `_on_search_results_ready`, `_on_tab_table_clicked`) |
| [`mixins/ingestion_and_maintenance.py`](mixins/ingestion_and_maintenance.py) | Version cards refresh, Smart Import, local/FTP indexing, release pruning, DB wipe/delete | `SpecSearchTabIngestionMixin` (`refresh_versions`, `_open_smart_import_dialog`, `_start_batch_ingestion`, `_on_prune_superseded_clicked`, `_on_wipe_db_clicked`) |
| [`mixins/clause_actions.py`](mixins/clause_actions.py) | Change History dialog, origin CR/TDoc attribution, Clause Diff, Word compare & Word/Markdown export | `SpecSearchTabClauseActionsMixin` (`_open_version_history_dialog`, `_on_find_origin_cr_clicked`, `_open_clause_diff_dialog`, `_handle_extract_clause_word`, `_handle_export_clause_markdown`) |
| [`mixins/matrix_navigation.py`](mixins/matrix_navigation.py) | Quick Diff, Jump-to-Version, double-click/context menus, `focus_spec`, and `shutdown` | `SpecSearchTabMatrixNavMixin` (`_handle_quick_diff`, `_handle_jump_to_version`, `_show_matrix_context_menu`, `focus_spec`, `shutdown`) |
| [`__init__.py`](__init__.py) | Package exports | `SpecSearchTab` |

## Backward Compatibility

- [`../spec_search_tabs.py`](../spec_search_tabs.py) remains as a thin re-export shim exposing `SpecSearchTab` for [`shell_setup.py`](../../../../core/ui/main_window/mixins/shell_setup.py) and the test suite.
