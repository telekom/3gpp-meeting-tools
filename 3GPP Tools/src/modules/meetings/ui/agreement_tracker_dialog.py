# --- File: src/modules/meetings/ui/agreement_tracker_dialog.py ---
"""Backward-compatible re-export shim for :mod:`modules.meetings.ui.agreement_tracker`."""

from __future__ import annotations

from modules.meetings.ui.agreement_tracker import (
    AgreementResultsTableModel,
    AgreementTrackerDialog,
    logger as _tracker_logger,
)

logger = _tracker_logger

__all__ = [
    "AgreementResultsTableModel",
    "AgreementTrackerDialog",
]
