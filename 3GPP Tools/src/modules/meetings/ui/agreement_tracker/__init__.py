"""Modular Cross-Meeting Study Item & Topic Agreement Tracker package."""

from .dialog import AgreementTrackerDialog, logger as _dialog_logger
from .model import AgreementResultsTableModel

logger = _dialog_logger

__all__ = [
    "AgreementResultsTableModel",
    "AgreementTrackerDialog",
]
