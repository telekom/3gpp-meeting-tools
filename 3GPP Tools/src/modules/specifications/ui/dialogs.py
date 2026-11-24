# --- File: src/modules/specifications/ui/dialogs.py ---
"""Backward-compatible facade re-exporting `modules.specifications.ui.spec_dialogs` and legacy dialog aliases."""

from modules.specifications.ui.spec_clause_dialog import SpecClauseExcerptDialog
from modules.specifications.ui.spec_dialogs import (
    AddSpecDialog,
    AdvancedSyncDialog,
    ManualSpecFetcherThread,
    SpecsConfigDialog,
    TableFilterDialog,
    TargetedSyncDialog,
)
from modules.specifications.ui.spec_info_dialog import NumericTableWidgetItem, SpecInfoDialog
from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

__all__ = [
    "AddSpecDialog",
    "AdvancedSyncDialog",
    "CrPcrDialog",
    "ManualSpecFetcherThread",
    "NumericTableWidgetItem",
    "SpecClauseExcerptDialog",
    "SpecInfoDialog",
    "SpecsConfigDialog",
    "TableFilterDialog",
    "TargetedSyncDialog",
]
