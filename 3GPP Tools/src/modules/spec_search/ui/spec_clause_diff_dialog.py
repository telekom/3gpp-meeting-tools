# --- File: src/modules/spec_search/ui/spec_clause_diff_dialog.py ---
"""Backward-compatible re-export shim for :mod:`modules.spec_search.ui.clause_diff`."""

from __future__ import annotations

from modules.spec_search.ui.clause_diff import (
    SpecClauseDiffDialog,
    logger as _diff_logger,
)

logger = _diff_logger

__all__ = [
    "SpecClauseDiffDialog",
]
