"""Mixins for `MeetingAgreementsDialog`."""

from __future__ import annotations

from modules.meetings.ui.meeting_agreements.mixins.actions_and_export import (
    MeetingAgreementsActionsMixin,
)
from modules.meetings.ui.meeting_agreements.mixins.ui_setup import (
    MeetingAgreementsUiSetupMixin,
)

__all__ = [
    "MeetingAgreementsActionsMixin",
    "MeetingAgreementsUiSetupMixin",
]
