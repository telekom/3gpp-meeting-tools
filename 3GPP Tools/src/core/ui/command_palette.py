"""Backward-compatible re-export shim for the Ctrl+K Command Palette.

Canonical implementation lives in ``core.ui.palette``.
"""

from __future__ import annotations

import logging

from core.ui.palette import (
    CommandItem,
    CommandPaletteDialog,
    DynamicPaletteSearcher,
    _KIND_HEADER,
    _KIND_ITEM,
    _PaletteRowDelegate,
    _ROLE_CMD,
    _ROLE_KIND,
    _TDOC_PATTERN,
    _score_command,
    build_static_command_items,
)

logger = logging.getLogger(__name__)

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
    "logger",
]
