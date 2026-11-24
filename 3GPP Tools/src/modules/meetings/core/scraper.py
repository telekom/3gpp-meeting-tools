# --- File: modules/meetings/core/scraper.py ---
"""Backward-compatible re-export shim for ``modules.meetings.core.crawler``."""

from __future__ import annotations

import logging

from .crawler import (
    MEETING_SOURCES,
    ManualMeetingFetcherThread,
    MeetingsCrawlerThread,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "MEETING_SOURCES",
    "MeetingsCrawlerThread",
    "ManualMeetingFetcherThread",
]
