"""Backward-compatible re-export shim for TrackedThreadDialog and PinFromOutlookDialog.

The implementation is organized under ``modules.emails.ui.tracked_thread``.
"""

from __future__ import annotations

from core.ai.cli_launcher import launch_cli_in_terminal

from .tracked_thread import (
    PinFromOutlookDialog,
    TrackedThreadDialog,
    _infer_meeting_tag,
)

__all__ = [
    "PinFromOutlookDialog",
    "TrackedThreadDialog",
    "_infer_meeting_tag",
    "launch_cli_in_terminal",
]
