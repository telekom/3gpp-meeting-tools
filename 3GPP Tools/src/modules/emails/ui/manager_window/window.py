"""EmailManagerWindow composing setup, filtering/selection, and worker mixins."""

from __future__ import annotations

import datetime
import logging
from pathlib import Path

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.ui.theme import C
from modules.emails.core.email_db import EmailDatabase
from modules.emails.ui.manager_window.mixins import (
    EmailManagerActionsAndWorkersMixin,
    EmailManagerFilteringAndSelectionMixin,
    EmailManagerUiSetupMixin,
)


class EmailManagerWindow(
    EmailManagerUiSetupMixin,
    EmailManagerFilteringAndSelectionMixin,
    EmailManagerActionsAndWorkersMixin,
    QWidget,
):
    """Split-pane eMeeting Email Manager window for TDoc-linked emails and Outlook sync."""

    tdoc_open_requested = pyqtSignal(str)
    tdoc_jump_requested = pyqtSignal(str)

    def __init__(
        self,
        meeting_dir: Path,
        ai_lookup: dict,
        meeting_start: str = "",
        meeting_end: str = "",
    ):
        super().__init__()
        logging.info("[EmailWindow] Initializing eMeeting Email Manager...")

        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup

        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self._apply_filters)

        self.db = EmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.config_path = self.meeting_dir / "Agenda" / "email_config.json"

        config_data = self._load_config()
        self.source_folder = config_data.get("source_folder", "")
        self.target_folder = config_data.get("target_folder", "")
        self.stats_config = config_data.get("stats_config", {})

        saved_start = config_data.get("start_date", "")
        saved_end = config_data.get("end_date", "")

        if saved_start and saved_end:
            self.start_date = saved_start
            self.end_date = saved_end
        else:
            try:
                if meeting_start and meeting_end:
                    s_dt = datetime.datetime.strptime(meeting_start, "%Y-%m-%d") - datetime.timedelta(days=3)
                    e_dt = datetime.datetime.strptime(meeting_end, "%Y-%m-%d") + datetime.timedelta(days=3)
                    self.start_date = s_dt.strftime("%Y-%m-%d")
                    self.end_date = e_dt.strftime("%Y-%m-%d")
                else:
                    self.start_date = ""
                    self.end_date = ""
            except Exception:
                self.start_date, self.end_date = "", ""

        self.setWindowTitle("📧 eMeeting Email Manager")
        self.resize(1300, 800)
        self.setStyleSheet(f"QWidget {{ background-color: {C.BG_ALT_LIGHT}; }}")

        self._setup_ui()
        self._refresh_table()

    def closeEvent(self, event):
        if hasattr(self, "search_timer"):
            self.search_timer.stop()
        for th in (
            getattr(self, "sync_thread", None),
            getattr(self, "move_thread", None),
            getattr(self, "rescan_thread", None),
            getattr(self, "stats_thread", None),
        ):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(150)
        super().closeEvent(event)

    def get_session_state(self) -> dict:
        """Return serializable EmailManagerWindow state."""
        return {
            "meeting_dir": str(self.meeting_dir),
            "tdoc_search": self.tdoc_search_input.text() if hasattr(self, "tdoc_search_input") else "",
            "email_search": self.search_input.text() if hasattr(self, "search_input") else "",
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore search filter text and apply filters."""
        if not isinstance(state, dict):
            return
        if hasattr(self, "tdoc_search_input") and isinstance(state.get("tdoc_search"), str):
            self.tdoc_search_input.setText(state["tdoc_search"])
        if hasattr(self, "search_input") and isinstance(state.get("email_search"), str):
            self.search_input.setText(state["email_search"])
        if hasattr(self, "search_timer"):
            self.search_timer.stop()
        if hasattr(self, "_apply_filters"):
            self._apply_filters()
