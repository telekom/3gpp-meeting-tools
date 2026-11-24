"""Word Tools Settings & Preferences dialog orchestrator."""

from __future__ import annotations

from typing import List

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog

from modules.word_tools.ui.word_settings.mixins import (
    WordSettingsActionsMixin,
    WordSettingsUiSetupMixin,
)


class WordSettingsDialog(
    QDialog,
    WordSettingsUiSetupMixin,
    WordSettingsActionsMixin,
):
    """Configuration dialog for Word Tools preferences, sensitivity labels, and templates."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Word Tools Settings & Preferences")
        self.setMinimumWidth(640)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._labels_cache: List[dict] = []
        self._tenant_id: str = ""
        self._policy_upn: str = ""
        self._setup_ui()
        self._load_values()
