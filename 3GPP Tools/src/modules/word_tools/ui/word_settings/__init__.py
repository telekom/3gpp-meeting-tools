"""Public facade for the Word Tools Settings dialog package."""

from __future__ import annotations

from modules.word_tools.ui.word_settings.dialog import WordSettingsDialog
from modules.word_tools.ui.word_settings.sensitivity_edit_dialog import (
    SensitivityLabelEditDialog,
)

__all__ = [
    "SensitivityLabelEditDialog",
    "WordSettingsDialog",
]
