"""Public facade and compatibility re-exports for the Contribution Report dialog package."""

from __future__ import annotations

from modules.meetings.core.contribution_worker import ContributionSearchWorker
from modules.meetings.ui.contributions.contribution_models import (
    ContributionResultsTableModel,
)
from modules.meetings.ui.contributions.contribution_widgets import (
    AITokenInputWidget,
    CompanySelectorWidget,
    KPICard,
    WITokenInputWidget,
)
from modules.meetings.ui.contributions.dialog.report_dialog import (
    ContributionReportDialog,
    logger,
)
from modules.meetings.ui.dialogs.contribution_export_controller import (
    export_to_excel,
    generate_statistics,
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
