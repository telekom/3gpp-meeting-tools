"""Backward-compatible re-export shim for Pill & Tag UI components.

Canonical implementation lives in ``modules.meetings.ui.pills``.
"""

from modules.meetings.ui.pills import (
    AgendaItemPill,
    AgendaItemPillsWidget,
    FilterPillButton,
    TagCloudWidget,
    TagPillItem,
    WgTogglePill,
    _TagLineEdit,
    get_wg_colors,
)

__all__ = [
    "AgendaItemPill",
    "AgendaItemPillsWidget",
    "FilterPillButton",
    "TagCloudWidget",
    "TagPillItem",
    "WgTogglePill",
    "_TagLineEdit",
    "get_wg_colors",
]
