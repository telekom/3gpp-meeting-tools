"""Backward-compatible re-export shim for TDocs table and proxy models."""

from modules.meetings.ui.tdocs.models import (
    TDocsFilterProxyModel,
    TDocsTableModel,
    logger as _models_logger,
    natural_sort_key,
)

logger = _models_logger

__all__ = [
    "natural_sort_key",
    "TDocsTableModel",
    "TDocsFilterProxyModel",
]
