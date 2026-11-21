"""Modular UI components for Specification Search & Release Evolution."""

from .clause_inspector import SpecClauseInspector
from .delegates import SpecEvolutionStatusDelegate
from .spec_card import ReleasePillWidget, SpecCardWidget
from .styles import (
    INSPECTOR_BTN_AI,
    INSPECTOR_BTN_BASE,
    INSPECTOR_BTN_INFO,
    INSPECTOR_BTN_SUCCESS,
    INSPECTOR_BTN_WARNING,
)
from .version_cards import (
    SpecSearchVersionCardsWidget,
    SpecSearchVersionTreeWidget,
    _TreeCompatChildItem,
    _TreeCompatRootItem,
    _TreeCompatSpecItem,
)

__all__ = [
    "INSPECTOR_BTN_AI",
    "INSPECTOR_BTN_BASE",
    "INSPECTOR_BTN_INFO",
    "INSPECTOR_BTN_SUCCESS",
    "INSPECTOR_BTN_WARNING",
    "ReleasePillWidget",
    "SpecCardWidget",
    "SpecClauseInspector",
    "SpecEvolutionStatusDelegate",
    "SpecSearchVersionCardsWidget",
    "SpecSearchVersionTreeWidget",
    "_TreeCompatChildItem",
    "_TreeCompatRootItem",
    "_TreeCompatSpecItem",
]
