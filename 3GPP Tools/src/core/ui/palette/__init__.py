"""Modular Global Ctrl+K Command Palette package."""

from core.ui.palette.delegate import _PaletteRowDelegate
from core.ui.palette.dialog import CommandPaletteDialog
from core.ui.palette.models import (
    CommandItem,
    _KIND_HEADER,
    _KIND_ITEM,
    _ROLE_CMD,
    _ROLE_KIND,
    _TDOC_PATTERN,
    _score_command,
)
from core.ui.palette.searcher import DynamicPaletteSearcher
from core.ui.palette.static_commands import build_static_command_items

__all__ = [
    "CommandItem",
    "CommandPaletteDialog",
    "DynamicPaletteSearcher",
    "_KIND_HEADER",
    "_KIND_ITEM",
    "_PaletteRowDelegate",
    "_ROLE_CMD",
    "_ROLE_KIND",
    "_TDOC_PATTERN",
    "_score_command",
    "build_static_command_items",
]
