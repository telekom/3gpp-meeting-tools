"""Feature mixins composing `ContributionReportDialog`."""

from __future__ import annotations

from modules.meetings.ui.contributions.dialog.mixins.row_actions_and_session import (
    ContributionDialogActionsMixin,
)
from modules.meetings.ui.contributions.dialog.mixins.search_and_kpis import (
    ContributionDialogSearchMixin,
)
from modules.meetings.ui.contributions.dialog.mixins.ui_setup import (
    ContributionDialogUiSetupMixin,
)

__all__ = [
    "ContributionDialogActionsMixin",
    "ContributionDialogSearchMixin",
    "ContributionDialogUiSetupMixin",
]
