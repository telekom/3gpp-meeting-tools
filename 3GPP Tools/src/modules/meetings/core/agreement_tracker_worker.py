"""Backward-compatible re-export shim for `modules.meetings.core.agreement_tracker`."""

from __future__ import annotations

from modules.meetings.core.agreement_tracker import (
    AgreementTrackerWorker,
    logger,
)

__all__ = [
    "AgreementTrackerWorker",
    "logger",
]
