"""Backward-compatible re-export shim for `modules.meetings.core.db`."""

from __future__ import annotations

from modules.meetings.core.db import MeetingsDatabase

__all__ = ["MeetingsDatabase"]
