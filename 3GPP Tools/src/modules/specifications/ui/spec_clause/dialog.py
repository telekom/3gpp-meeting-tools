"""Modeless interactive clause extraction dialog composed from modular mixins."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from core.ui.theme import C
from core.ui.ui_components import CARD_FRAME_STYLE
from modules.specifications.ui.threads import SpecExcerptPrepThread
from modules.word_tools.core.word_excerpt_extractor import HeadingInfo, WordExcerptThread
from .mixins import (
    SpecClauseExtractionMixin,
    SpecClauseTreeMixin,
    SpecClauseUiSetupMixin,
)

logger = logging.getLogger(__name__)


class SpecClauseExcerptDialog(
    SpecClauseUiSetupMixin,
    SpecClauseTreeMixin,
    SpecClauseExtractionMixin,
    QDialog,
):
    """
    Independent dialog enabling interactive clause/subclause extraction directly
    from 3GPP specifications. Handles automated fetching, multi-part ordering,
    legacy .doc conversion, search filtering, and clean Word excerpt generation.
    """

    extraction_completed = pyqtSignal(Path)

    def __init__(
        self,
        spec_num: str,
        spec_type: str,
        version: str,
        fname: str,
        url: str,
        spec_dl_dir: Path,
        parent=None,
        auto_start: bool = True,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top of the main window)
        self._parent_ref = parent
        self.spec_num = spec_num.strip()
        self.spec_type = (spec_type or "TS").upper().strip()
        self.version = version.strip()
        self.fname = fname.strip()
        self.url = url.strip()
        self.spec_dl_dir = Path(spec_dl_dir)

        self._prep_thread: Optional[SpecExcerptPrepThread] = None
        self._extract_thread: Optional[WordExcerptThread] = None

        self._ordered_paths: List[str] = []
        self._headings: List[HeadingInfo] = []
        self._is_updating_checks = False

        self.setWindowTitle(
            f"✂️ Extract Clauses: {self.spec_type} {self.spec_num} v{self.version}"
        )
        # Modeless top-level independent window configuration:
        # Note: Must use Qt.Window rather than self.windowFlags() | Qt.Window so Qt.Dialog (0x3)
        # is cleared, allowing the window to freely move behind the main app window in z-order.
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent is not None:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.resize(900, 700)
        self.setMinimumWidth(740)
        self.setMinimumHeight(540)

        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()
        if auto_start and self.spec_num and self.version:
            self._start_preparation()

    def closeEvent(self, event):
        """Cleanly abort active worker threads if dialog is closed."""
        prep = getattr(self, "_prep_thread", None)
        if prep and prep.isRunning():
            prep.requestInterruption()
            if not prep.wait(1000):
                prep.terminate()
                prep.wait(500)
        ext = getattr(self, "_extract_thread", None)
        if ext and ext.isRunning():
            ext.requestInterruption()
            if not ext.wait(1000):
                ext.terminate()
                ext.wait(500)
        super().closeEvent(event)

    def get_session_state(self) -> dict:
        """Return serializable SpecClauseExcerptDialog context."""
        return {
            "spec_num": str(self.spec_num or ""),
            "spec_type": str(self.spec_type or "TS"),
            "version": str(self.version or ""),
            "fname": str(self.fname or ""),
            "url": str(self.url or ""),
            "spec_dl_dir": str(self.spec_dl_dir),
            "filter_text": self.filter_input.text() if hasattr(self, "filter_input") else "",
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore clause tree search filter."""
        if not isinstance(state, dict):
            return
        ft = state.get("filter_text")
        if isinstance(ft, str) and ft and hasattr(self, "filter_input"):
            self.filter_input.setText(ft)


SpecClauseDialog = SpecClauseExcerptDialog
