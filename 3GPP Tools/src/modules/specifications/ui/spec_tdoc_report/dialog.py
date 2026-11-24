"""Modeless TDoc Technical Dossier dialog composed from modular mixins."""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional, Sequence

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QWidget

from core.ui.theme import C
from core.ui.ui_components import CARD_FRAME_STYLE
from .mixins import (
    SpecTDocReportActionsMixin,
    SpecTDocReportTableMixin,
    SpecTDocReportUiSetupMixin,
)

logger = logging.getLogger(__name__)


class SpecTDocReportDialog(
    SpecTDocReportUiSetupMixin,
    SpecTDocReportTableMixin,
    SpecTDocReportActionsMixin,
    QDialog,
):
    """Modeless technical dossier dialog for inspecting all TDocs integrated in a specification."""

    def __init__(
        self,
        spec_number: str,
        version: str,
        tdoc_entries: Sequence[Dict[str, Any]],
        markdown_report: str = "",
        spec_title: str = "",
        parent: Optional[QWidget] = None,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top)
        self._parent_ref = parent
        self.spec_number = spec_number.strip()
        self.version = version.lstrip("vV").strip()
        self.all_entries = list(tdoc_entries)
        self.filtered_entries = list(tdoc_entries)
        self.markdown_report = markdown_report
        self.spec_title = spec_title

        # Top-level modeless independent window configuration
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass

        self.setWindowTitle(
            f"📑 TDoc Technical Dossier — TS {self.spec_number} v{self.version}"
        )
        self.resize(1120, 740)
        self.setMinimumSize(880, 540)

        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()
        self._populate_filters()
        self._refresh_table()
