"""Modular TDocs Cockpit Quick Guide & Keyboard Shortcuts Dialog package."""

from __future__ import annotations

from .dialog import TDocsHelpDialog
from .styles import HELP_CSS
from .tabs_html import (
    _build_search_tab_html,
    _build_shortcuts_tab_html,
    _build_table_tab_html,
    _build_toolbar_tab_html,
)

__all__ = [
    "TDocsHelpDialog",
    "HELP_CSS",
    "_build_toolbar_tab_html",
    "_build_table_tab_html",
    "_build_search_tab_html",
    "_build_shortcuts_tab_html",
]
