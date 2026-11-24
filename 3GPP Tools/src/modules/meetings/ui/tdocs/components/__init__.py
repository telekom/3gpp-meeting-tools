# --- File: src/modules/meetings/ui/tdocs/components/__init__.py ---
"""Modular reusable UI widgets for TDocs & Specifications filter/comparison toolbars."""

from .checkable_combo import CheckableComboBox
from .combo_delegate import CheckableComboItemDelegate
from .comparison_cart import (
    ComparisonCartWidget,
    make_rich_tooltip,
    populate_slot_word_menu,
)

__all__ = [
    "CheckableComboItemDelegate",
    "CheckableComboBox",
    "populate_slot_word_menu",
    "ComparisonCartWidget",
    "make_rich_tooltip",
]
