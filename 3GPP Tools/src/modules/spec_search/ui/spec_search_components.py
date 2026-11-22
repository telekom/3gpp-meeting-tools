"""Backward-compatible re-export shim for Spec Search UI components.

The implementation was split into ``modules.spec_search.ui.components``
(see ``components/README.md`` for a submodule map).
"""

from __future__ import annotations

from .components import (  # noqa: F401
    ActiveSpecsScopeBar,
    INSPECTOR_BTN_AI,
    INSPECTOR_BTN_BASE,
    INSPECTOR_BTN_INFO,
    INSPECTOR_BTN_SUCCESS,
    INSPECTOR_BTN_WARNING,
    ReleasePillWidget,
    SpecCardWidget,
    SpecClauseInspector,
    SpecEvolutionStatusDelegate,
    SpecSearchVersionCardsWidget,
    SpecSearchVersionTreeWidget,
    _TreeCompatChildItem,
    _TreeCompatRootItem,
    _TreeCompatSpecItem,
)

__all__ = [
    "ActiveSpecsScopeBar",
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
