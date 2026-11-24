"""Public facade for the meetings SQLite database package."""

from __future__ import annotations

from modules.meetings.core.db.database import MeetingsDatabase

__all__ = ["MeetingsDatabase"]
