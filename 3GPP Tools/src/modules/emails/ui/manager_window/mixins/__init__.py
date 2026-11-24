"""Mixins package for EmailManagerWindow."""

from modules.emails.ui.manager_window.mixins.actions_and_workers import (
    EmailManagerActionsAndWorkersMixin,
)
from modules.emails.ui.manager_window.mixins.filtering_and_selection import (
    EmailManagerFilteringAndSelectionMixin,
)
from modules.emails.ui.manager_window.mixins.ui_setup import EmailManagerUiSetupMixin

__all__ = [
    "EmailManagerActionsAndWorkersMixin",
    "EmailManagerFilteringAndSelectionMixin",
    "EmailManagerUiSetupMixin",
]
