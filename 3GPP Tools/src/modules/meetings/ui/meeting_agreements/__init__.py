"""Public facade for the Meeting Agreements & Conclusions dialog package."""

from __future__ import annotations

from modules.meetings.ui.meeting_agreements.dialog import (
    MeetingAgreementsDialog,
    logger,
)
from modules.meetings.ui.meeting_agreements.model import (
    AgreementsTableModel,
    _clean_agreement_preview,
)

__all__ = [
    "AgreementsTableModel",
    "MeetingAgreementsDialog",
    "_clean_agreement_preview",
    "logger",
]
