# --- File: src/modules/meetings/ui/session_plan/dialogs/__init__.py ---
"""Modular SA2 Session Plan Dialogs package."""

import logging

from .core_dialogs import (
    TIMEZONE_PRESETS,
    SessionPlanDiffDialog,
    SessionPlanImportDialog,
    SessionPlanTimezoneDialog,
)
from .delegation_dispatch_dialogs import (
    MultiDelegateSelectDialog,
    SessionDispatchDialog,
)
from .delegation_roster_dialogs import (
    DelegationImportDialog,
    DelegationMemberDialog,
)
from .slot_and_topic_dialogs import (
    SessionPlanNoteDialog,
    SessionPlanRoomLocationsDialog,
    SessionPlanTrackedTopicsDialog,
)

logger = logging.getLogger("modules.meetings.ui.session_plan.session_plan_dialogs")

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
