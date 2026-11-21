# Tracked Thread Dialogs Package (`tracked_thread/`)

Modular decomposition of `TrackedThreadDialog` and `PinFromOutlookDialog` previously housed in `tracked_thread_dialog.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`thread_dialog.py`](thread_dialog.py) | Main `TrackedThreadDialog` `QDialog` composition and window lifecycle | `TrackedThreadDialog`, `thread_changed`, `__init__`, `showEvent`, `closeEvent` |
| [`mixins/ui_setup.py`](mixins/ui_setup.py) | Toolbar, collapsible Links & Tags panel, filter bar, reader pane, and LLM/export menu (0 raw hexes) | `TrackedThreadUiSetupMixin` (`_build_ui`, `_on_order_changed`, `_on_toggle_links_panel`, `_update_links_toggle_label`, `_populate_llm_export_menu`) |
| [`mixins/rendering.py`](mixins/rendering.py) | Company filter population, data loading, message filtering, Markdown rendering, unread jump navigation | `TrackedThreadRenderingMixin` (`_populate_company_filter`, `_load_data`, `_render_messages`, `_update_next_unread_button`, `_scroll_to_initial_unread`, `_on_next_unread_clicked`) |
| [`mixins/links_and_notes.py`](mixins/links_and_notes.py) | Cascading `AI -> Sub-Topic` dropdown scoping, TDoc family expansion, tags, and `InteractiveNotesDialog` launcher | `TrackedThreadLinksNotesMixin` (`_normalize_ai_sub_map`, `_get_valid_sub_items_for_ais`, `_refresh_sub_items_dropdown`, `_on_links_changed`, `_on_edit_notes_and_status`) |
| [`mixins/actions.py`](mixins/actions.py) | Interactive link dispatch (`tdoc:`/`outlook://`/docs), async Outlook refresh workers, split-thread merge/unpin, Markdown export & idle LLM CLI launch | `call_launch_cli_in_terminal`, `TrackedThreadActionsMixin` (`_on_anchor_clicked`, `_start_thread_worker`, `_on_merge_thread`, `_on_launch_cli`) |
| [`pin_from_outlook_dialog.py`](pin_from_outlook_dialog.py) | Outlook WG folder discovery, coverage/tag/date filtering, ignore list, and pinning dialog (0 raw hexes) | `_infer_meeting_tag`, `PinFromOutlookDialog` |
| [`__init__.py`](__init__.py) | Package exports | `TrackedThreadDialog`, `PinFromOutlookDialog`, `_infer_meeting_tag` |

## Backward Compatibility & Test Monkeypatching

- [`../tracked_thread_dialog.py`](../tracked_thread_dialog.py) remains as a thin re-export shim exposing `TrackedThreadDialog`, `PinFromOutlookDialog`, `_infer_meeting_tag`, and `launch_cli_in_terminal`.
- `mixins/actions.py` invokes `call_launch_cli_in_terminal()`, which checks `modules.emails.ui.tracked_thread_dialog.launch_cli_in_terminal` dynamically so existing unit-test `patch(...)` targets continue to intercept `_on_launch_cli()` without modification.
