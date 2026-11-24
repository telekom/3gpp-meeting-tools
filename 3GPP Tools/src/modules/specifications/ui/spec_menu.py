"""Backward-compatible re-export shim for specification context menus and copy helpers.

Canonical implementation lives in ``modules.specifications.ui.menu``.
"""

from __future__ import annotations

import logging

from modules.specifications.ui.menu import (
    CONFIG_FILE_NAME,
    _copy_to_clipboard,
    build_citation_data,
    copy_specs_default,
    format_canonical_spec_number,
    get_saved_copy_mode,
    populate_multi_spec_copy_menu,
    populate_spec_copy_menu,
    resolve_spec_metadata,
    save_copy_mode,
    show_spec_context_menu,
)

logger = logging.getLogger(__name__)

__all__ = [
    "CONFIG_FILE_NAME",
    "_copy_to_clipboard",
    "build_citation_data",
    "copy_specs_default",
    "format_canonical_spec_number",
    "get_saved_copy_mode",
    "logger",
    "populate_multi_spec_copy_menu",
    "populate_spec_copy_menu",
    "resolve_spec_metadata",
    "save_copy_mode",
    "show_spec_context_menu",
]
