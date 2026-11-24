"""Document Sanitizer & Revision Hygiene Hub card composed from modular mixins."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtWidgets import QWidget

from .mixins import SanitizerActionsMixin, SanitizerUiSetupMixin


class WordSanitizerCard(SanitizerUiSetupMixin, SanitizerActionsMixin, QWidget):
    """Document Sanitizer, Proofing Language Unification, Language Revision Purge, and Author Consolidation."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_author_file: str = ""
        self._last_modified_author_file: str = ""
        self._current_author_stats: list = []
        self._setup_ui()
