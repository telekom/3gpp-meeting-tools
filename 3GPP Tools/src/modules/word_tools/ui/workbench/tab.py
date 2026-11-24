"""Modernized Word Tools tab composed from modular workbench mixins."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from .mixins import (
    WordTabCompatAndCallbacksMixin,
    WordTabStatusAndPollMixin,
    WordTabUiSetupMixin,
)


class WordExtractorTab(
    WordTabUiSetupMixin,
    WordTabStatusAndPollMixin,
    WordTabCompatAndCallbacksMixin,
    QWidget,
):
    """Modernized Word Tools tab with 2-tier ribbon, live Word COM monitor, and modular operation cards."""

    # Public signals connected by MainWindow
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str)
    compare_doc_requested = pyqtSignal(str, str, dict)
    convert_doc_requested = pyqtSignal(str, str)
    convert_markdown_requested = pyqtSignal(str, dict)
    convert_markdown_to_word_requested = pyqtSignal(str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._active_word_file: Optional[Path] = None
        self._setup_ui()

        # Active Word background polling (every 3 seconds)
        self.word_poll_timer = QTimer(self)
        self.word_poll_timer.timeout.connect(self._poll_active_word)
        self.word_poll_timer.start(3000)
        self._poll_active_word()

        # Environment status checks
        self._check_libreoffice_status()
        self._check_dotx_status()
