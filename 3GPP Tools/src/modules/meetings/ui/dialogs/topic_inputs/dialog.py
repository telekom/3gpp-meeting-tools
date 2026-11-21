"""TopicInputsViewerDialog shell combining UI setup, collection, table, and drawer mixins."""

import logging
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from modules.meetings.core.cohort_resolver import MeetingCohort
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.core.topic_manager import MetaTopicDefinition
from modules.meetings.ui.dialogs.topic_inputs.collector_thread import InputsCollectorThread
from modules.meetings.ui.dialogs.topic_inputs.mixins import (
    TopicInputsCollectionFilteringMixin,
    TopicInputsDrawerActionsMixin,
    TopicInputsTablesMixin,
    TopicInputsUiSetupMixin,
)

logger = logging.getLogger("modules.meetings.ui.dialogs.topic_inputs_dialog")


class TopicInputsViewerDialog(
    TopicInputsUiSetupMixin,
    TopicInputsCollectionFilteringMixin,
    TopicInputsTablesMixin,
    TopicInputsDrawerActionsMixin,
    QDialog,
):
    """Interactive dialog visualizing considered inputs, agreements, pCRs, CRs, and LSs."""

    export_requested = pyqtSignal()

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        selected_rounds: List[MeetingCohort],
        scope_mode: str = "full_dossier",
        timeframe_label: str = "",
        parent=None,
    ):
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowTitle(f"⚖️ Considered Inputs & Decisions — {topic.name}")
        self.resize(1140, 720)

        self.meetings_db = meetings_db
        self.topic = topic
        self.rounds = list(selected_rounds)
        self.scope_mode = scope_mode
        self.timeframe_label = timeframe_label or f"{len(self.rounds)} Synchronized Meeting Rounds"

        self.all_decisions: List[TopicDecision] = []
        self.all_contributions: List[Dict[str, Any]] = []
        self.ls_tracker = CrossWgLiaisonTracker()

        self._displayed_decisions: List[TopicDecision] = []
        self._displayed_contribs: List[Dict[str, Any]] = []
        self._displayed_ls: List[Dict[str, Any]] = []
        self._current_detail_text: str = ""

        self._collector_worker: Optional[InputsCollectorThread] = None

        self._init_ui()
        self._start_collection()

    def closeEvent(self, event):
        if self._collector_worker and self._collector_worker.isRunning():
            self._collector_worker.cancel()
            self._collector_worker.wait(3000)
        super().closeEvent(event)
