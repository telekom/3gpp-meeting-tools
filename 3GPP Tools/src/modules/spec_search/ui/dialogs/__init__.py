"""Modular Spec Search dialogs package."""

from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog

from .prune_dialog import SpecSearchPruneDialog
from .smart_import_dialog import SpecSmartImportDialog
from .version_select_dialog import SpecSearchVersionSelectDialog

__all__ = [
    "SpecClauseDiffDialog",
    "SpecSearchPruneDialog",
    "SpecSearchVersionSelectDialog",
    "SpecSmartImportDialog",
]
