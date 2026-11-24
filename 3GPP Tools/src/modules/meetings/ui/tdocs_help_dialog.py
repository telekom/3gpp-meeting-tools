# --- File: src/modules/meetings/ui/tdocs_help_dialog.py ---
"""Backward-compatible re-export shim for ``modules.meetings.ui.tdocs_help``."""

from __future__ import annotations

from .tdocs_help import (
    HELP_CSS,
    TDocsHelpDialog,
    _build_search_tab_html,
    _build_shortcuts_tab_html,
    _build_table_tab_html,
    _build_toolbar_tab_html,
)

__all__ = [
    "HELP_CSS",
    "TDocsHelpDialog",
    "_build_toolbar_tab_html",
    "_build_table_tab_html",
    "_build_search_tab_html",
    "_build_shortcuts_tab_html",
]
