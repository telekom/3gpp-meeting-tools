"""Feature mixins composing SpecTDocReportDialog."""

from __future__ import annotations

from .actions_and_export import SpecTDocReportActionsMixin
from .filtering_and_table import SpecTDocReportTableMixin
from .ui_setup import SpecTDocReportUiSetupMixin

__all__ = [
    "SpecTDocReportUiSetupMixin",
    "SpecTDocReportTableMixin",
    "SpecTDocReportActionsMixin",
]
