"""Backward-compatible re-export shim for World Clock timezone service.

Canonical implementation lives in ``core.ui.world_clock_engine``.
"""

from __future__ import annotations

import logging

from core.network.session import NetworkSession
from core.ui.world_clock_engine import (
    VenueClockState,
    WebTimezoneSyncWorker,
    WorldClockService,
    WorldClockSnapshot,
    _COUNTRY_CITY_FALLBACKS,
    _clean_city_query,
    _format_wg_summary,
    _is_web_entry_fresh,
    _short_city_name,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone

logger = logging.getLogger(__name__)

__all__ = [
    "NetworkSession",
    "SessionPlanTimezone",
    "VenueClockState",
    "WebTimezoneSyncWorker",
    "WorldClockService",
    "WorldClockSnapshot",
    "_COUNTRY_CITY_FALLBACKS",
    "_clean_city_query",
    "_format_wg_summary",
    "_is_web_entry_fresh",
    "_short_city_name",
    "logger",
]
