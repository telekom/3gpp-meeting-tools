"""Shared AI email summary UI components (`InlineEmailSummaryCard`, model picker, defaults dialog)."""

from .model_picker import SummaryModelDefaultsDialog, populate_summary_model_menu
from .summary_card import InlineEmailSummaryCard

__all__ = [
    "InlineEmailSummaryCard",
    "SummaryModelDefaultsDialog",
    "populate_summary_model_menu",
]
