"""Protocol Matrix & Inspector tab composed from modular feature mixins."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.utils.paths import get_project_root, get_specs_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.nas_threads import NASFetchAndImportThread
from modules.nas.ui.nas_workers import NASMessageListWorker, ReverseLookupWorker
from modules.specifications.core.database import SpecsDatabase
from .mixins import (
    NASTabLeftPaneMixin,
    NASTabLibraryAndConfigMixin,
    NASTabMessagesAndMatrixMixin,
    NASTabNavigationAndIngestMixin,
    NASTabRightPaneMixin,
)


class NASTab(
    NASTabLeftPaneMixin,
    NASTabRightPaneMixin,
    NASTabLibraryAndConfigMixin,
    NASTabMessagesAndMatrixMixin,
    NASTabNavigationAndIngestMixin,
    QWidget,
):
    """Main UI tab for Protocol Evolution Matrix, Release Ingestion, and Definition Inspection."""

    log_msg = pyqtSignal(str, int)
    navigate_to_specs_tab_requested = pyqtSignal(str, str)  # spec_number, version
    search_in_spec_search_requested = pyqtSignal(str, str)  # spec_number, version

    def __init__(
        self,
        nas_db_path: Path,
        specs_db_path: Optional[Path] = None,
        config_path: Optional[Path] = None,
    ):
        super().__init__()
        self.nas_db_path = Path(nas_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = Path(config_path) if config_path else (get_project_root() / "config" / "nas_config.json")

        self._reverse_lookup_worker: Optional[ReverseLookupWorker] = None
        self._reverse_lookup_request_id: int = 0
        self._msg_query_workers: Dict[int, NASMessageListWorker] = {}
        self._msg_query_req_id: int = 0
        self._reverse_lookup_workers: Dict[int, ReverseLookupWorker] = {}
        self.thread: Optional[NASFetchAndImportThread] = None
        self._ingest_thread: Optional[NASFetchAndImportThread] = None

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = get_specs_root()

        self.db = NASDatabase(self.nas_db_path)
        self.specs_db = (
            SpecsDatabase(self.specs_db_path)
            if self.specs_db_path and self.specs_db_path.exists()
            else None
        )

        self.selected_version_ids: List[int] = []
        self._browse_mode: str = "messages"  # "messages" or "ies"
        self.current_selected_message_name: Optional[str] = None
        self.current_selected_ie_key: Optional[str] = None
        self._current_message_spec: Optional[str] = None
        self._current_message_clause: Optional[str] = None
        self._loading_config: bool = False
        self._initialized_filters: bool = False

        self._msg_search_timer = QTimer(self)
        self._msg_search_timer.setSingleShot(True)
        self._msg_search_timer.setInterval(200)
        self._msg_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._ie_search_timer = QTimer(self)
        self._ie_search_timer.setSingleShot(True)
        self._ie_search_timer.setInterval(200)
        self._ie_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._setup_ui()
        self.refresh_versions()
