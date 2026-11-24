"""Feature mixins composing `WordSettingsDialog`."""

from __future__ import annotations

from modules.word_tools.ui.word_settings.mixins.actions_and_persistence import (
    WordSettingsActionsMixin,
)
from modules.word_tools.ui.word_settings.mixins.ui_setup import (
    WordSettingsUiSetupMixin,
)

__all__ = [
    "WordSettingsActionsMixin",
    "WordSettingsUiSetupMixin",
]
