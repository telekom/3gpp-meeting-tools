"""Backward-compatible re-export shim for TDocs context menus and Quick Copy.

Canonical implementation lives in ``modules.meetings.ui.tdocs.menus``.
"""

from modules.meetings.ui.tdocs.menus import (
    _copy_to_clipboard,
    build_action_menu,
    build_quick_copy_menu,
    build_related_menu,
    build_row_context_menu,
    is_tdoc_cached,
    open_3gu_portal,
    open_tdoc_local_folder,
    populate_quick_copy_menu,
)
from modules.meetings.ui.tdocs.menus.quick_copy import logger

__all__ = [
    "_copy_to_clipboard",
    "build_action_menu",
    "build_quick_copy_menu",
    "build_related_menu",
    "build_row_context_menu",
    "is_tdoc_cached",
    "logger",
    "open_3gu_portal",
    "open_tdoc_local_folder",
    "populate_quick_copy_menu",
]
