# --- File: src/modules/meetings/core/session_plan/session_plan_parser.py ---
"""Backward-compatible re-export shim for the modular SA2 Session Plan Parser."""

import logging

from modules.meetings.core.session_plan.parser import (
    ADMIN_NOTE_RE,
    DAY_NAMES,
    PERIOD_ORDER,
    PERIOD_TIMES,
    QUARTER_TIMES,
    ROOM_OVERRIDE_RE,
    SessionPlanParser,
)

logger = logging.getLogger(__name__)

__all__ = [
    "PERIOD_ORDER",
    "PERIOD_TIMES",
    "QUARTER_TIMES",
    "DAY_NAMES",
    "ROOM_OVERRIDE_RE",
    "ADMIN_NOTE_RE",
    "SessionPlanParser",
]
