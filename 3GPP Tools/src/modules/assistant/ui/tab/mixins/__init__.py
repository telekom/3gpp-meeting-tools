"""UI mixin package for AssistantTab."""

from modules.assistant.ui.tab.mixins.docking_and_layout import (
    AssistantDockingLayoutMixin,
)
from modules.assistant.ui.tab.mixins.research_runner import (
    AssistantResearchRunnerMixin,
)
from modules.assistant.ui.tab.mixins.session_and_export import (
    AssistantSessionExportMixin,
)
from modules.assistant.ui.tab.mixins.ui_setup import AssistantUiSetupMixin

__all__ = [
    "AssistantDockingLayoutMixin",
    "AssistantResearchRunnerMixin",
    "AssistantSessionExportMixin",
    "AssistantUiSetupMixin",
]
