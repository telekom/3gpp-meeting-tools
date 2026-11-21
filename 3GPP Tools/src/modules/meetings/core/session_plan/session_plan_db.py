# --- File: src/modules/meetings/core/session_plan/session_plan_db.py ---
"""
Backward-compatible import shim for SessionPlanDatabase.

The implementation was split into ``src/modules/meetings/core/session_plan/db/``
(see ``db/README.md`` for a 'where to edit what' map).
"""

from modules.meetings.core.session_plan.db import (  # noqa: F401
    SessionPlanDatabase,
    logger,
)

__all__ = [
    "SessionPlanDatabase",
    "logger",
]
