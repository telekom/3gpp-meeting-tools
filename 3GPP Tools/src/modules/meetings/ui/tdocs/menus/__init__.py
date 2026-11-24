"""Modular TDoc browser context menus and centralized Quick Copy actions."""

from modules.meetings.ui.tdocs.menus.action_and_related import build_action_menu, build_related_menu
from modules.meetings.ui.tdocs.menus.quick_copy import (
    _copy_to_clipboard,
    build_quick_copy_menu,
    is_tdoc_cached,
    open_3gu_portal,
    open_tdoc_local_folder,
    populate_quick_copy_menu,
)
from modules.meetings.ui.tdocs.menus.row_context_menu import build_row_context_menu

__all__ = [
    "_copy_to_clipboard",
    "build_action_menu",
    "build_quick_copy_menu",
    "build_related_menu",
    "build_row_context_menu",
    "is_tdoc_cached",
    "open_3gu_portal",
    "open_tdoc_local_folder",
    "populate_quick_copy_menu",
]
