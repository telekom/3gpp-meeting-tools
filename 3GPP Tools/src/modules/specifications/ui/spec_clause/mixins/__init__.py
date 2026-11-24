"""Feature mixins composing SpecClauseExcerptDialog."""

from __future__ import annotations

from .extraction import SpecClauseExtractionMixin
from .tree_and_filter import SpecClauseTreeMixin
from .ui_setup import SpecClauseUiSetupMixin

__all__ = [
    "SpecClauseUiSetupMixin",
    "SpecClauseTreeMixin",
    "SpecClauseExtractionMixin",
]
