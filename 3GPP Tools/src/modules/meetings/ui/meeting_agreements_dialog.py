"""Backward-compatible re-export shim for `modules.meetings.ui.meeting_agreements`."""

from __future__ import annotations

from modules.meetings.ui.meeting_agreements import (
    AgreementsTableModel,
    MeetingAgreementsDialog,
    _clean_agreement_preview,
    logger,
)

__all__ = [
    "AgreementsTableModel",
    "MeetingAgreementsDialog",
    "_clean_agreement_preview",
    "logger",
]
