# Meetings Tab Package (`meetings_tab/`)

Modular decomposition of `MeetingsTab` previously housed in the 1,387-line `ui_tabs.py`.
`ui_tabs.py` remains a thin backward-compatible re-export shim exposing `MeetingsTab`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`tab.py`](tab.py) | Main `QWidget` composition, signals (`update_db_requested`, `update_specific_requested`, `update_meetings_requested`, `abort_sync_requested`), lifecycle shutdown | `MeetingsTab` (`__init__`, `shutdown`, `closeEvent`) |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | 2-tier `PageRibbonCard`, Quick TDoc Jump bar, filter strip, table + delegates, cart footer, DPI font scaling | `MeetingsTabUiSetupMixin` (`_setup_ui`, `on_font_scale_changed`, `_apply_scaled_table_columns`) |
| [`mixins/sync_controls.py`](mixins/sync_controls.py) | Intent-driven `Sync ▾` menu, inline `sync_status_banner` with `QProgressBar` + Stop button, custom `MeetingsSyncDialog` launcher | `MeetingsTabSyncMixin` (`request_quick_sync`, `on_sync_started`, `on_sync_progress`, `on_sync_finished`) |
| [`mixins/filters_and_data.py`](mixins/filters_and_data.py) | `QSettings`/`MeetingsSettings` filter persistence, TSG/WG/`Unknown Dates` presets, `refresh_table`, `focus_meeting` | `MeetingsTabFiltersDataMixin` (`_load_filters`, `_save_filters`, `refresh_table`, `focus_meeting`) |
| [`mixins/quick_tdoc_and_menus.py`](mixins/quick_tdoc_and_menus.py) | Quick TDoc alerts & context menu, right-click row menus, comparison cart hooks, meeting CRUD confirmations | `MeetingsTabQuickTdocMenusMixin` (`_on_tdoc_input_context_menu`, `_set_tdoc_metadata_warning`, `show_right_click_menu`) |
| [`mixins/windows_and_actions.py`](mixins/windows_and_actions.py) | `TDocsWindow` launcher, inline downloads, global TDoc actions, report dialogs, `CrPcrDialog` launcher | `MeetingsTabWindowsActionsMixin` (`_open_tdocs_window`, `_handle_global_action_from_window`, `open_cr_pcr_for_meeting`) |
| [`__init__.py`](__init__.py) | Package exports | `MeetingsTab` |

## Invariants
- **Zero Raw Hexes (`tests/test_ui_hex_lint.py`):** All stylesheets and rich HTML labels use `core.ui.theme.C` tokens (`0` raw `#RRGGBB` literals).
- **Responsive Quick TDoc Wrap:** Keep Quick TDoc action buttons as independent children of `ribbon_card`'s primary `FlowLayout` so they wrap cleanly without clipping.
- **Clean `shutdown()`:** Close tracked `TDocsWindow`, `_session_plan_window`, `contribution_dialog`, `agreement_tracker_dialog`, and cooperatively stop background threads on tab/app exit.
- Keep every file under **565 lines**.
