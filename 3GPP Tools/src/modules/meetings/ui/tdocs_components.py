# --- File: src/modules/meetings/ui/tdocs_components.py ---
"""Backward-compatible re-export shim for modular TDocs UI components."""

from modules.meetings.ui.tdocs.components import (
    CheckableComboBox,
    CheckableComboItemDelegate,
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
