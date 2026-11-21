"""Cross-WG, Multi-Meeting Topic Dossier & Briefing Exporter Thread class."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import logging
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.cohort_resolver import MeetingCohort
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import MetaTopicDefinition
from .mixins import (
    TopicDossierArtifactsMixin,
    TopicDossierCacheAndConversionMixin,
    TopicDossierDiscoveryMixin,
    TopicDossierExportPipelineMixin,
)
from .path_utils import get_default_topic_export_base

logger = logging.getLogger(__name__)


class TopicDossierExporterThread(
    TopicDossierExportPipelineMixin,
    TopicDossierCacheAndConversionMixin,
    TopicDossierDiscoveryMixin,
    TopicDossierArtifactsMixin,
    QThread,
):
    progress_msg = pyqtSignal(str)
    progress_step = pyqtSignal(int, int)
    session_finished = pyqtSignal(bool, str, str)  # success, target_dir, summary

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        selected_cohorts: List[MeetingCohort],
        export_scope_mode: str = "full_dossier",  # "pre_meeting", "post_meeting", "full_dossier"
        output_base_dir: Optional[Path] = None,
        target_folder_name: Optional[str] = None,
        include_discussions: bool = True,
        deduplicate_revisions: bool = True,
        generate_consolidated_md: bool = True,
        allow_network_download: bool = True,
        parent=None,
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.topic = topic
        self.cohorts = list(selected_cohorts)
        self.export_scope_mode = export_scope_mode
        self.output_base_dir = Path(output_base_dir) if output_base_dir else get_default_topic_export_base()
        self.target_folder_name = target_folder_name
        self.include_discussions = bool(include_discussions)
        self.deduplicate_revisions = bool(deduplicate_revisions)
        self.generate_consolidated_md = bool(generate_consolidated_md)
        self.allow_network_download = bool(allow_network_download)

        self._is_cancelled = False
        self._executor: Optional[ThreadPoolExecutor] = None

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass

    @property
    def is_cancelled(self) -> bool:
        return self._is_cancelled or self.isInterruptionRequested()

    def run(self):
        try:
            self._run_export()
        except Exception as e:
            logger.exception(f"[TopicDossierExporter] Export failed: {e}")
            self.session_finished.emit(False, "", f"Export failed: {str(e)}")
