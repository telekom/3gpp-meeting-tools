"""Modular Meeting Email Inbox cockpit package."""

from __future__ import annotations

from .dialog import MeetingEmailInboxDialog
from .widgets import (
    _STATUS_BADGE_BG_MAP,
    _STATUS_COLOR_MAP,
    _InboxStatusStripDelegate,
    _InboxTableWidget,
    _NotePlainTextEdit,
    _status_color_hex,
    get_compact_action_bar_qss,
    get_status_combo_qss,
)

__all__ = [
    "MeetingEmailInboxDialog",
    "_InboxStatusStripDelegate",
    "_InboxTableWidget",
    "_NotePlainTextEdit",
    "_STATUS_BADGE_BG_MAP",
    "_STATUS_COLOR_MAP",
    "_status_color_hex",
    "get_compact_action_bar_qss",
    "get_status_combo_qss",
]
