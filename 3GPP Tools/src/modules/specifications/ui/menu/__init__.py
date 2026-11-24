"""Modular specification context-menu and clipboard-citation package."""

from modules.specifications.ui.menu.clipboard_and_citations import (
    CONFIG_FILE_NAME,
    _copy_to_clipboard,
    build_citation_data,
    format_canonical_spec_number,
    get_saved_copy_mode,
    resolve_spec_metadata,
    save_copy_mode,
)
from modules.specifications.ui.menu.context_menu import show_spec_context_menu
from modules.specifications.ui.menu.copy_menus import (
    copy_specs_default,
    populate_multi_spec_copy_menu,
    populate_spec_copy_menu,
)

__all__ = [
    "CONFIG_FILE_NAME",
    "_copy_to_clipboard",
    "build_citation_data",
    "copy_specs_default",
    "format_canonical_spec_number",
    "get_saved_copy_mode",
    "populate_multi_spec_copy_menu",
    "populate_spec_copy_menu",
    "resolve_spec_metadata",
    "save_copy_mode",
    "show_spec_context_menu",
]
