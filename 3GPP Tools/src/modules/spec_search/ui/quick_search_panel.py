# --- File: src/modules/spec_search/ui/quick_search_panel.py ---
"""Backward-compatible re-export shim for ``modules.spec_search.ui.quick_search``."""

from __future__ import annotations

import logging

from .quick_search import (
    QuickSpecSearchPanel,
    _QuickSearchResultDelegate,
    _format_snippet_html,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "QuickSpecSearchPanel",
    "_QuickSearchResultDelegate",
    "_format_snippet_html",
]
