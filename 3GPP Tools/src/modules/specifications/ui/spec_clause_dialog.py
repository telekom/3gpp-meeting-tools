# --- File: src/modules/specifications/ui/spec_clause_dialog.py ---
"""Backward-compatible re-export shim for ``modules.specifications.ui.spec_clause``."""

from __future__ import annotations

import logging

from .spec_clause import SpecClauseDialog, SpecClauseExcerptDialog

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "SpecClauseExcerptDialog",
    "SpecClauseDialog",
]
