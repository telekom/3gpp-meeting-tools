"""Modular UI dialogs for Tracked Email Threads and Outlook folder discovery."""

from .pin_from_outlook_dialog import PinFromOutlookDialog, _infer_meeting_tag
from .thread_dialog import TrackedThreadDialog

__all__ = [
    "PinFromOutlookDialog",
    "TrackedThreadDialog",
    "_infer_meeting_tag",
]
