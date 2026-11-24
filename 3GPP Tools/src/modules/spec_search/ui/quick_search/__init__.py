"""Modular docked QuickSpecSearchPanel package."""

from __future__ import annotations

from .delegate import _QuickSearchResultDelegate, _format_snippet_html
from .panel import QuickSpecSearchPanel

__all__ = [
    "QuickSpecSearchPanel",
    "_QuickSearchResultDelegate",
    "_format_snippet_html",
]
