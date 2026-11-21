# --- File: src/modules/meetings/ui/session_plan/session_plan_dialogs.py ---
"""Backward-compatible re-export shim for SA2 Session Plan dialogs.

The implementation was split into ``src/modules/meetings/ui/session_plan/dialogs/``
(see ``dialogs/README.md`` for a submodule map).
"""

from __future__ import annotations

from .dialogs import (  # noqa: F401
    TIMEZONE_PRESETS,
    DelegationImportDialog,
    DelegationMemberDialog,
    MultiDelegateSelectDialog,
    SessionDispatchDialog,
    SessionPlanDiffDialog,
    SessionPlanImportDialog,
    SessionPlanNoteDialog,
    SessionPlanRoomLocationsDialog,
    SessionPlanTimezoneDialog,
    SessionPlanTrackedTopicsDialog,
    logger,
)

__all__ = [
    "TIMEZONE_PRESETS",
    "DelegationImportDialog",
    "DelegationMemberDialog",
    "MultiDelegateSelectDialog",
    "SessionDispatchDialog",
    "SessionPlanDiffDialog",
    "SessionPlanImportDialog",
    "SessionPlanNoteDialog",
    "SessionPlanRoomLocationsDialog",
    "SessionPlanTimezoneDialog",
    "SessionPlanTrackedTopicsDialog",
    "logger",
]
