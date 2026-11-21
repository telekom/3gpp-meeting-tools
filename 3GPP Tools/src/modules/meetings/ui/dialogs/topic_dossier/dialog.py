"""TopicDossierDialog shell combining setup, topics/status, and export/CLI mixins."""

from typing import List, Optional

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import QDialog

from modules.meetings.core.cohort_resolver import MeetingCohort
from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import MetaTopicDefinition, MetaTopicManager
from modules.meetings.ui.dialogs.topic_dossier.mixins import (
    TopicDossierExportCliMixin,
    TopicDossierTopicsStatusMixin,
    TopicDossierUiSetupMixin,
)


class TopicDossierDialog(
    TopicDossierUiSetupMixin,
    TopicDossierTopicsStatusMixin,
    TopicDossierExportCliMixin,
    QDialog,
):
    """Main modeless studio for Cross-WG, Multi-Meeting topic dossiers."""

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        self.setWindowTitle("🌐 Cross-WG Topic Dossier & Briefing Studio")
        self.resize(1020, 680)
        self.meetings_db = meetings_db
        self.topic_manager = MetaTopicManager()
        self.active_topic: Optional[MetaTopicDefinition] = None
        self.selected_cohorts: List[MeetingCohort] = []
        self._worker: Optional[TopicDossierExporterThread] = None

        self._init_ui()
        self._load_topics_list()

    def closeEvent(self, event):
        if getattr(self, "_inputs_dlg", None) is not None:
            try:
                self._inputs_dlg.close()
            except Exception:
                pass
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait(5000)
        super().closeEvent(event)

    def get_session_state(self) -> dict:
        """Return serializable TopicDossierDialog state."""
        return {
            "topic_id": self.active_topic.id if self.active_topic is not None else "",
            "cycle_idx": self.combo_cycle_window.currentIndex() if hasattr(self, "combo_cycle_window") else 0,
            "date_from": self.edit_date_from.date().toString("yyyy-MM-dd") if hasattr(self, "edit_date_from") else "",
            "date_to": self.edit_date_to.date().toString("yyyy-MM-dd") if hasattr(self, "edit_date_to") else "",
            "dest_folder": self.edit_dest_folder.text() if hasattr(self, "edit_dest_folder") else "",
            "include_discussions": bool(self.chk_discussions.isChecked()) if hasattr(self, "chk_discussions") else True,
            "dedup": bool(self.chk_dedup.isChecked()) if hasattr(self, "chk_dedup") else True,
            "consolidated": bool(self.chk_consolidated.isChecked()) if hasattr(self, "chk_consolidated") else True,
            "download": bool(self.chk_download.isChecked()) if hasattr(self, "chk_download") else True,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore active topic, cycle preset, date range, and dossier options."""
        if not isinstance(state, dict):
            return
        topic_id = str(state.get("topic_id") or "").strip()
        if topic_id and hasattr(self, "_load_topics_list"):
            self._load_topics_list(preferred_id=topic_id)

        cycle_idx = state.get("cycle_idx")
        if isinstance(cycle_idx, int) and hasattr(self, "combo_cycle_window"):
            if 0 <= cycle_idx < self.combo_cycle_window.count():
                self.combo_cycle_window.setCurrentIndex(cycle_idx)

        d_from = QDate.fromString(str(state.get("date_from") or ""), "yyyy-MM-dd")
        if d_from.isValid() and hasattr(self, "edit_date_from"):
            self.edit_date_from.setDate(d_from)
        d_to = QDate.fromString(str(state.get("date_to") or ""), "yyyy-MM-dd")
        if d_to.isValid() and hasattr(self, "edit_date_to"):
            self.edit_date_to.setDate(d_to)

        dest = state.get("dest_folder")
        if isinstance(dest, str) and dest and hasattr(self, "edit_dest_folder"):
            self.edit_dest_folder.setText(dest)

        for key, attr in (
            ("include_discussions", "chk_discussions"),
            ("dedup", "chk_dedup"),
            ("consolidated", "chk_consolidated"),
            ("download", "chk_download"),
        ):
            if key in state and hasattr(self, attr):
                getattr(self, attr).setChecked(bool(state[key]))
