# --- File: src/modules/meetings/core/session_plan/session_plan_exporter.py ---
"""
Backward-compatible import shim for the Session Plan Exporter.

The implementation was split into the ``exporter/`` package (see
``exporter/README.md`` for a 'where to edit what' map).
"""
from modules.meetings.core.session_plan.exporter import (
    SessionPlanExporter,
    logger,
)

__all__ = ["SessionPlanExporter", "logger"]
