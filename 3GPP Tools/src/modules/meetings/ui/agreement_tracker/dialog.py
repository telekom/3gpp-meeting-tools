"""Modeless Cross-Meeting Study Item & Topic Agreement Tracker dialog shell."""

import logging
from typing import Dict, Optional

from PyQt5.QtCore import QDate, Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from modules.meetings.core.agreement_tracker_worker import AgreementTrackerWorker
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_settings import WorkItemsSettings

from .mixins import (
    _TrackerActionsMixin,
    _TrackerExportMixin,
    _TrackerSearchMixin,
    _TrackerUiMixin,
)

logger = logging.getLogger(__name__)


class AgreementTrackerDialog(
    _TrackerUiMixin,
    _TrackerSearchMixin,
    _TrackerActionsMixin,
    _TrackerExportMixin,
    QDialog,
):
    """
    Independent modeless Cross-Meeting Study Item & Topic Agreement Tracker.
    """

    open_meeting_requested = pyqtSignal(str, str)  # (wg_name, meeting_number)

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)

        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.wi_settings = WorkItemsSettings()
        self.worker: Optional[AgreementTrackerWorker] = None
        self.active_tdoc_threads: Dict[str, TDocActionThread] = {}

        self.setWindowTitle("⚖️ 3GPP Cross-Meeting Study Item & Topic Agreement Tracker")
        self.resize(1420, 860)
        self._setup_ui()

    def closeEvent(self, event):
        self._cleanup_threads()
        super().closeEvent(event)

    def reject(self):
        self._cleanup_threads()
        super().reject()

    def get_session_state(self) -> dict:
        """Return serializable AgreementTrackerDialog filter state."""
        return {
            "wgs": self.wg_filter.getCheckedItems() if hasattr(self, "wg_filter") else [],
            "date_from": self.date_from.date().toString("yyyy-MM-dd") if hasattr(self, "date_from") else "",
            "date_to": self.date_to.date().toString("yyyy-MM-dd") if hasattr(self, "date_to") else "",
            "wi_tokens": self.wi_token_widget.get_tokens() if hasattr(self, "wi_token_widget") else [],
            "ai_tokens": self.ai_token_widget.get_tokens() if hasattr(self, "ai_token_widget") else [],
            "wi_ai_or": bool(self.chk_wi_ai_or.isChecked()) if hasattr(self, "chk_wi_ai_or") else False,
            "keyword": self.keyword_edit.text() if hasattr(self, "keyword_edit") else "",
            "agreements": bool(self.chk_agreements.isChecked()) if hasattr(self, "chk_agreements") else True,
            "ffs": bool(self.chk_ffs.isChecked()) if hasattr(self, "chk_ffs") else True,
            "agreed_tdocs": bool(self.chk_agreed_tdocs.isChecked()) if hasattr(self, "chk_agreed_tdocs") else True,
            "prev_recaps": bool(self.chk_prev_recaps.isChecked()) if hasattr(self, "chk_prev_recaps") else False,
            "auto_download": bool(self.chk_auto_download.isChecked()) if hasattr(self, "chk_auto_download") else True,
            "bypass_cache": bool(self.chk_bypass_cache.isChecked()) if hasattr(self, "chk_bypass_cache") else False,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore WG, date, token, keyword, and checkbox filters."""
        if not isinstance(state, dict):
            return
        wgs = state.get("wgs")
        if isinstance(wgs, list) and hasattr(self, "set_active_wgs"):
            self.set_active_wgs([str(w) for w in wgs])

        d_from = QDate.fromString(str(state.get("date_from") or ""), "yyyy-MM-dd")
        if d_from.isValid() and hasattr(self, "date_from"):
            self.date_from.setDate(d_from)
        d_to = QDate.fromString(str(state.get("date_to") or ""), "yyyy-MM-dd")
        if d_to.isValid() and hasattr(self, "date_to"):
            self.date_to.setDate(d_to)

        if isinstance(state.get("wi_tokens"), list) and hasattr(self.wi_token_widget, "set_tokens"):
            self.wi_token_widget.set_tokens(state["wi_tokens"])
        if isinstance(state.get("ai_tokens"), list) and hasattr(self.ai_token_widget, "set_tokens"):
            self.ai_token_widget.set_tokens(state["ai_tokens"])
        if "wi_ai_or" in state and hasattr(self, "chk_wi_ai_or"):
            self.chk_wi_ai_or.setChecked(bool(state["wi_ai_or"]))
        if isinstance(state.get("keyword"), str) and hasattr(self, "keyword_edit"):
            self.keyword_edit.setText(state["keyword"])

        for key, attr in (
            ("agreements", "chk_agreements"),
            ("ffs", "chk_ffs"),
            ("agreed_tdocs", "chk_agreed_tdocs"),
            ("prev_recaps", "chk_prev_recaps"),
            ("auto_download", "chk_auto_download"),
            ("bypass_cache", "chk_bypass_cache"),
        ):
            if key in state and hasattr(self, attr):
                getattr(self, attr).setChecked(bool(state[key]))
