"""Backward-compatible re-export shim for `modules.meetings.ui.contributions.dialog`."""

from __future__ import annotations

from modules.meetings.ui.contributions.dialog import (
    AITokenInputWidget,
    CompanySelectorWidget,
    ContributionReportDialog,
    ContributionResultsTableModel,
    ContributionSearchWorker,
    KPICard,
    WITokenInputWidget,
    export_to_excel,
    generate_statistics,
    logger,
    on_export_finished,
    on_stats_finished,
)

__all__ = [
    "AITokenInputWidget",
    "CompanySelectorWidget",
    "ContributionReportDialog",
    "ContributionResultsTableModel",
    "ContributionSearchWorker",
    "KPICard",
    "WITokenInputWidget",
    "export_to_excel",
    "generate_statistics",
    "logger",
    "on_export_finished",
    "on_stats_finished",
]
