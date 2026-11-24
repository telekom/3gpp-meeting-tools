"""Meeting Agreements & Conclusions Dialog orchestrator."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QDialog

from modules.meetings.ui.meeting_agreements.mixins import (
    MeetingAgreementsActionsMixin,
    MeetingAgreementsUiSetupMixin,
)
from modules.meetings.ui.meeting_agreements.model import AgreementsTableModel

logger = logging.getLogger(__name__)


class MeetingAgreementsDialog(
    QDialog,
    MeetingAgreementsUiSetupMixin,
    MeetingAgreementsActionsMixin,
):
    """
    Dedicated dialog providing structured viewing, searching, inspection,
    and export of all agreements in the current meeting.
    """

    locate_tdoc_requested = pyqtSignal(str)

    def __init__(
        self,
        agreements: List[dict],
        meeting_name: str = "",
        meeting_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.raw_agreements = list(agreements or [])
        self.meeting_name = meeting_name or "3GPP Meeting"
        self.meeting_dir = Path(meeting_dir) if meeting_dir else None

        self.setWindowTitle(f"⚖️ Meeting Agreements & Conclusions — {self.meeting_name}")
        self.resize(1020, 680)
        self.setMinimumSize(800, 500)

        self.model = AgreementsTableModel(self.raw_agreements, parent=self)
        self._init_ui()
        self._update_kpi_counters()

        # Pre-select first item if available
        if self.model.rowCount() > 0:
            self.table.selectRow(0)
            self._on_table_row_selected(self.model.index(0, 0))
