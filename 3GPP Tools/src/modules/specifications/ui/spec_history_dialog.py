# --- File: src/modules/specifications/ui/spec_history_dialog.py ---
"""Backward-compatible re-export shim for :mod:`modules.specifications.ui.spec_history`."""

from __future__ import annotations

from modules.specifications.ui.spec_history import (
    NumericSortItem,
    SpecVersionHistoryDialog,
    logger as _history_logger,
)

logger = _history_logger

__all__ = [
    "NumericSortItem",
    "SpecVersionHistoryDialog",
]
