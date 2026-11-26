# --- File: src/modules/meetings/ui/session_plan/session_plan_window.py ---
"""
Backward-compatible import shim for the Session Plan Studio window.

The implementation was split into the ``window/`` package (see window/README.md
for a 'where to edit what' map). Keep importing from here or from the package.
"""
from modules.meetings.ui.session_plan.window import (
    DAY_NAMES,
    QUARTER_ORDER,
    QUARTER_TIMES,
    GridRowDescriptor,
    SessionPlanGridTable,
    SessionPlanHorizontalHeader,
    SessionPlanVerticalHeader,
    SessionPlanWindow,
    _format_compact_snapshot_label,
)

__all__ = [
    "DAY_NAMES",
    "QUARTER_ORDER",
    "QUARTER_TIMES",
    "GridRowDescriptor",
    "SessionPlanGridTable",
    "SessionPlanHorizontalHeader",
    "SessionPlanVerticalHeader",
    "SessionPlanWindow",
    "_format_compact_snapshot_label",
]
