# --- File: src/modules/specifications/ui/spec_tdoc_report_dialog.py ---
"""Backward-compatible re-export shim for ``modules.specifications.ui.spec_tdoc_report``."""

from __future__ import annotations

import logging

from .spec_tdoc_report import SpecTDocReportDialog

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "SpecTDocReportDialog",
]
