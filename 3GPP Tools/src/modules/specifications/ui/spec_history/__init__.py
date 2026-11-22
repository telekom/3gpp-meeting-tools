"""Modular Specification Change History & Integrated TDocs Viewer package."""

from .dialog import SpecVersionHistoryDialog, logger as _dialog_logger
from .mixins import NumericSortItem

logger = _dialog_logger

__all__ = [
    "NumericSortItem",
    "SpecVersionHistoryDialog",
]
