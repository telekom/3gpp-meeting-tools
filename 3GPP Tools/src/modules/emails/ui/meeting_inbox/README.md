# Meeting Email Inbox Package (`meeting_inbox/`)

Modular decomposition of `MeetingEmailInboxDialog` (Unified Unread & Flagged Email Triage Cockpit) previously housed in `meeting_email_inbox_dialog.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`dialog.py`](dialog.py) | Main `QDialog` composition, state initialization, `live_sync_toggled` signal & 2m live timer | `MeetingEmailInboxDialog`, `data_changed`, `live_sync_toggled`, `__init__` |
| [`tdoc_grouping.py`](tdoc_grouping.py) | Pure-Python `📄 By TDoc` / `🧵 By Thread` grouping, signal rollup & `🔥 Priority` radar | `enrich_message_item_signals`, `build_grouped_thread_items`, `build_grouped_tdoc_items`, `is_priority_item`, `is_item_effectively_parked` |
| [`followed_ais_helper.py`](followed_ais_helper.py) | Auto-inherited Followed AIs (from WIs, Session Plan, My TDocs) + per-meeting overrides | `compute_inherited_followed_ais`, `load_effective_followed_ais`, `save_followed_ais_override` |
| [`widgets.py`](widgets.py) | Shared custom Qt delegates/widgets, grouped `#inboxTopBar` / `#inboxViewSegment` / `#inboxFilterBar` builders, and QSS helpers (0 raw hexes) | `_STATUS_COLOR_MAP`, `build_inbox_top_toolbar`, `build_inbox_filter_bar`, `build_inbox_sync_controls`, `get_sync_toolbutton_qss`, `get_live_sync_pill_qss`, `get_compact_action_bar_qss`, `get_status_combo_qss`, `_InboxStatusStripDelegate`, `_NotePlainTextEdit`, `_InboxTableWidget` |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Master-detail splitter panes, detail header, inline bars, and 1-click filter reset | `InboxUiSetupMixin` (`_build_ui`, `_clear_all_filters`) |
| [`mixins/layout_and_settings.py`](mixins/layout_and_settings.py) | Per-meeting `QSettings` filter persistence, initial live catch-up on show, splitter orientation, responsive column fitting, session state | `InboxLayoutAndSettingsMixin` (`_load_filter_settings`, `_save_filter_settings`, `compute_fitted_column_widths`, `_fit_columns_to_viewport`, `showEvent`, `closeEvent`, `get_session_state`) |
| [`mixins/data_and_filtering.py`](mixins/data_and_filtering.py) | DB loading (`TrackedThreadsDatabase` + `GeneralEmailDatabase`), thread grouping, pill counts, table population | `InboxDataAndFilteringMixin` (`reload_from_dbs`, `_build_grouped_thread_items`, `_populate_filter_combos`, `_apply_filters`) |
| [`mixins/detail_rendering.py`](mixins/detail_rendering.py) | Single/multi selection state, assignment pill population, Markdown preview rendering | `InboxDetailRenderingMixin` (`_on_selection_changed`, `_render_multi_detail`, `_render_detail`) |
| [`mixins/inline_assignment.py`](mixins/inline_assignment.py) | In-place unpinned thread pinning, auto-suggestion, batch inline AI/Sub-Topic/WI linking | `InboxInlineAssignmentMixin` (`_pin_item_in_place`, `_on_pin_inline_clicked`, `_on_auto_assign_clicked`, `_on_inline_links_changed`) |
| [`mixins/triage_actions.py`](mixins/triage_actions.py) | Bidirectional read/flag/status/note mutations and interactive reader link dispatch | `InboxTriageActionsMixin` (`_mutate_read`, `_mutate_flag`, `_on_detail_status_changed`, `_flush_pending_note_save`, `_on_anchor_clicked`) |
| [`mixins/sync_and_merge.py`](mixins/sync_and_merge.py) | Context menu, split-thread merge workflow, `⚡ Sync New`, `🟢 Live (2m)` toggle, and `📅 Full Meeting Sync (Date Range)...` | `InboxSyncAndMergeMixin` (`_on_table_context_menu`, `_on_merge_selected_clicked`, `_start_outlook_sync`, `_on_full_meeting_sync_clicked`, `set_live_sync_enabled`, `_on_inbox_sync_finished`) |
| [`__init__.py`](__init__.py) | Package exports | `MeetingEmailInboxDialog` + shared widget/QSS exports |

## Backward Compatibility

- [`../meeting_email_inbox_dialog.py`](../meeting_email_inbox_dialog.py) remains as a thin re-export shim exposing `MeetingEmailInboxDialog`, custom widgets, and shared QSS helpers used by [`../tdoc_emails_dialog.py`](../tdoc_emails_dialog.py), [`window_session.py`](../../../../core/ui/window_session.py), and the test suite.
