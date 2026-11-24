# `modules.specifications.ui.menu`

Modular context-menu and citation-copy package decomposed from the former `spec_menu.py` monolith:

- [`clipboard_and_citations.py`](clipboard_and_citations.py): Saved link-mode config, rich HTML/plain clipboard helper, metadata resolution, and citation payload builder.
- [`copy_menus.py`](copy_menus.py): Single-spec and multi-spec Quick Copy submenus plus default `Ctrl+C` copy handler.
- [`context_menu.py`](context_menu.py): Specifications table right-click context menu builder using `core.ui.theme.C` tokens.
