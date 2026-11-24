# TDocs Menus Package (`modules.meetings.ui.tdocs.menus`)

Acyclic modular package extracted from `ui/tdocs_menus.py` (thin backward-compatible re-export shim):

- `quick_copy.py`: Cache checks (`is_tdoc_cached`), folder/portal launchers (`open_tdoc_local_folder`, `open_3gu_portal`), clipboard helper (`_copy_to_clipboard` resolved via facade for tests), and `populate_quick_copy_menu()` / `build_quick_copy_menu()`.
- `action_and_related.py`: Column 0 opener `build_action_menu()` and hyperlink context menu `build_related_menu()`.
- `row_context_menu.py`: Full right-click command center `build_row_context_menu()`.
