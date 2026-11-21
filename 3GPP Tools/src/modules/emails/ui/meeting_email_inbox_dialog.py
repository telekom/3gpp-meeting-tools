"""Backward-compatible re-export shim for MeetingEmailInboxDialog and shared inbox widgets.

The implementation is organized under ``modules.emails.ui.meeting_inbox``.
"""

from __future__ import annotations

from .meeting_inbox import (
    _STATUS_BADGE_BG_MAP,
    _STATUS_COLOR_MAP,
    MeetingEmailInboxDialog,
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
