# `core.ui.palette`

Modular Ctrl+K Command Palette & Quick Switcher package decomposed from the former `command_palette.py` monolith:

- [`models.py`](models.py): `CommandItem` dataclass, Qt role constants, `_TDOC_PATTERN`, and `_score_command`.
- [`delegate.py`](delegate.py): `_PaletteRowDelegate` for group headers and action rows.
- [`searcher.py`](searcher.py): `DynamicPaletteSearcher` across cached specs/protocols and SQLite meetings/TDocs/WIs.
- [`static_commands.py`](static_commands.py): `build_static_command_items` for navigation tabs, side panels, workbenches, and dialogs.
- [`dialog.py`](dialog.py): `CommandPaletteDialog` modal switcher and keyboard event filter.
