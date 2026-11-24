"""Independent modeless cross-meeting ContributionReportDialog orchestrator."""

from __future__ import annotations

import logging

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.ui.contributions.dialog.mixins import (
    ContributionDialogActionsMixin,
    ContributionDialogSearchMixin,
    ContributionDialogUiSetupMixin,
)
from modules.work_items.core.wi_database import WorkItemsDatabase

logger = logging.getLogger(__name__)


class ContributionReportDialog(
    QDialog,
    ContributionDialogUiSetupMixin,
    ContributionDialogSearchMixin,
    ContributionDialogActionsMixin,
):
    """Independent modeless cross-meeting contribution report window."""

    open_meeting_requested = pyqtSignal(str, str)  # (wg_name, meeting_number)

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        # Pass None as parent so the OS does not pin this window on top of the main window.
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)

        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.worker = None
        self.exporter_thread = None
        self.stats_thread = None
        self.active_tdoc_threads = {}

        self.setWindowTitle("3GPP Company Contribution Search & Report")
        self.resize(1340, 820)
        self._setup_ui()

    def _cleanup_threads(self):
        """Cleanly halts all worker threads to prevent QThread destroyed while running."""
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.quit()
            self.worker.wait(300)
        if self.exporter_thread and self.exporter_thread.isRunning():
            self.exporter_thread.quit()
            self.exporter_thread.wait(200)
        if self.stats_thread and self.stats_thread.isRunning():
            self.stats_thread.quit()
            self.stats_thread.wait(200)
        for th in list(self.active_tdoc_threads.values()):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(200)
        self.active_tdoc_threads.clear()

    def closeEvent(self, event):
        self._cleanup_threads()
        super().closeEvent(event)

    def reject(self):
        self._cleanup_threads()
        super().reject()
