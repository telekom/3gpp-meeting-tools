# Protocol Matrix & Inspector Tab Package (`nas_tab/`)

Modular decomposition of `NASTab` previously housed in `nas/ui/nas_tabs.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`tab.py`](tab.py) | `QWidget` entry point, signals, debounce timers, DB wiring | `NASTab`, `log_msg`, `navigate_to_specs_tab_requested`, `search_in_spec_search_requested`, `__init__` |
| [`styles.py`](styles.py) | Context/ref regex patterns and toolbar/segmented QSS constants | `RE_SPEC_FROM_COL`, `RE_CLAUSE_FROM_REF`, `RE_CTX_*`, `SEGMENTED_MODE_STYLE`, `TOGGLE_BUTTON_STYLE`, `TINY_HEADER_BTN_STYLE`, `RIBBON_BADGE_STYLE` |
| [`mixins/ui_left_pane.py`](mixins/ui_left_pane.py) | Left vertical splitter: Indexed Protocols card and Protocol Library list | `NASTabLeftPaneMixin._build_left_pane` |
| [`mixins/ui_right_pane.py`](mixins/ui_right_pane.py) | Top-level horizontal splitter and Right vertical splitter (Matrix + Inspector) | `NASTabRightPaneMixin` (`_setup_ui`, `_build_right_pane`) |
| [`mixins/library_and_config.py`](mixins/library_and_config.py) | Quick spec selectors, matrix copy, stats ribbon, config persistence, version cards & banner | `NASTabLibraryAndConfigMixin` (`refresh_versions`, `_update_message_context_banner`, `_copy_matrix_formatted`, `_load_config`, `_save_config`) |
| [`mixins/messages_and_matrix.py`](mixins/messages_and_matrix.py) | Async message/IE list population, diff filtering, matrix model rebuild, double-click spec jump | `NASTabMessagesAndMatrixMixin` (`_populate_messages`, `_on_message_list_ready`, `_on_message_clicked`, `_rebuild_matrix_model`) |
| [`mixins/navigation_and_ingest.py`](mixins/navigation_and_ingest.py) | Context menus, Open Spec menus, reverse lookup jump, ingest delegation, shutdown | `NASTabNavigationAndIngestMixin` (`open_specification`, `_build_open_spec_menu`, `_jump_to_message`, `shutdown`) |
| [`__init__.py`](__init__.py) | Package exports | `NASTab` and regex/style constants |

## Backward Compatibility

- [`../nas_tabs.py`](../nas_tabs.py) remains as a thin re-export shim exposing `NASTab`, worker classes (`NASMessageListWorker`, `ReverseLookupWorker`), regexes, and style constants for [`shell_setup.py`](../../../../core/ui/main_window/mixins/shell_setup.py), [`ui/__init__.py`](../__init__.py), and the test suite.
