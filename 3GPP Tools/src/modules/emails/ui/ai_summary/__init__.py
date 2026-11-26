"""Shared AI email summary UI components (`InlineEmailSummaryCard`, model picker, defaults dialog)."""

from .m365_copilot_dialog import M365CopilotHandoffDialog
from .model_picker import SummaryModelDefaultsDialog, populate_summary_model_menu
from .summary_card import InlineEmailSummaryCard

__all__ = [
    "InlineEmailSummaryCard",
    "M365CopilotHandoffDialog",
    "SummaryModelDefaultsDialog",
    "populate_summary_model_menu",
]
