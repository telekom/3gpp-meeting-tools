"""Backward-compatible re-export shim for PostMeetingExportDialog."""

from modules.meetings.ui.dialogs.post_meeting_export import (
    STYLE_COMPACT_BTN,
    STYLE_TOGGLE_ACTIVE,
    STYLE_TOGGLE_INACTIVE,
    STYLE_TREE,
    PostMeetingExportDialog,
    logger as _dialog_logger,
    natural_ai_sort_key,
)

logger = _dialog_logger

__all__ = [
    "PostMeetingExportDialog",
    "natural_ai_sort_key",
    "STYLE_TOGGLE_ACTIVE",
    "STYLE_TOGGLE_INACTIVE",
    "STYLE_COMPACT_BTN",
    "STYLE_TREE",
]
