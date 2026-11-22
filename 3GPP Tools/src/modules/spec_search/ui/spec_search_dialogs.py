# --- File: src/modules/spec_search/ui/spec_search_dialogs.py ---
"""Backward-compatible re-export shim for :mod:`modules.spec_search.ui.dialogs`."""

from __future__ import annotations

from modules.spec_search.ui.dialogs import (
    SpecClauseDiffDialog,
    SpecSearchPruneDialog,
    SpecSearchVersionSelectDialog,
    SpecSmartImportDialog,
)

__all__ = [
    "SpecClauseDiffDialog",
    "SpecSearchPruneDialog",
    "SpecSearchVersionSelectDialog",
    "SpecSmartImportDialog",
]
