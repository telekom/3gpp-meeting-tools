"""Feature mixins composing QuickSpecSearchPanel."""

from __future__ import annotations

from .search_and_preview import QuickSearchLogicMixin
from .ui_setup import QuickSearchUiSetupMixin

__all__ = [
    "QuickSearchUiSetupMixin",
    "QuickSearchLogicMixin",
]
