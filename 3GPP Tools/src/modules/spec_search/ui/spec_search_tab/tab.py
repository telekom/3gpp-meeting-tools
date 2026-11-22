"""Specification Substring Search & Evolution Matrix tab composed from modular mixins."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.utils.paths import get_project_root, get_specs_root
from modules.meetings.core.settings import MeetingsSettings
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_search_threads import (
    SpecSearchImportThread,
    SpecSearchQueryWorker,
)
from modules.specifications.core.database import SpecsDatabase
from .mixins import (
    SpecSearchTabClauseActionsMixin,
    SpecSearchTabIngestionMixin,
    SpecSearchTabMatrixNavMixin,
    SpecSearchTabSearchMixin,
    SpecSearchTabUiSetupMixin,
)


class SpecSearchTab(
    SpecSearchTabUiSetupMixin,
    SpecSearchTabSearchMixin,
    SpecSearchTabIngestionMixin,
    SpecSearchTabClauseActionsMixin,
    SpecSearchTabMatrixNavMixin,
    QWidget,
):
    """Main Search Tab: Tree, Query Inputs, Cutoff Date Filter, Tabbed Evolution Matrix, and Inspector."""

    log_msg = pyqtSignal(str, int)
    navigate_to_specs_tab_requested = pyqtSignal(str, str)  # spec_number, version

    def __init__(self, search_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.search_db_path = Path(search_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = get_project_root() / "config" / "spec_search_config.json"

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = get_specs_root()

        self.db = SpecSearchDatabase(self.search_db_path)
        self.specs_db = SpecsDatabase(self.specs_db_path) if self.specs_db_path and self.specs_db_path.exists() else None
        self._spec_titles: Dict[str, str] = {}
        self._load_spec_titles()

        self._query_worker: Optional[SpecSearchQueryWorker] = None
        self._active_query_workers: Dict[int, SpecSearchQueryWorker] = {}
        self._current_req_id: int = 0
        self.selected_version_ids: List[int] = []
        self._specs_minimal_mode: bool = False
        self._solo_spec: Optional[str] = None
        self._pre_collapse_splitter_sizes: List[int] = [300, 750]
        self._active_doc_workers: List[object] = []
        self.auto_replace_older_versions: bool = True
        self.thread: Optional[SpecSearchImportThread] = None
        self._import_thread: Optional[SpecSearchImportThread] = None

        # Debounce timer for search execution
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self._execute_search)

        # Debounce timer for saving filter/selection state
        self._save_config_timer = QTimer(self)
        self._save_config_timer.setSingleShot(True)
        self._save_config_timer.setInterval(500)
        self._save_config_timer.timeout.connect(self._save_config)

        self._setup_ui()
        self._load_config_and_refresh()

    def _load_spec_titles(self):
        """Loads specification numbers and formal titles from SpecsDatabase into memory cache."""
        if not self.specs_db:
            default_path = get_project_root() / "db" / "3gpp_data.db"
            if default_path.exists():
                self.specs_db = SpecsDatabase(default_path)

        if self.specs_db:
            try:
                self._spec_titles = self.specs_db.get_all_spec_titles()
            except Exception as e:
                logging.getLogger(__name__).warning(f"Could not load spec titles: {e}")
                self._spec_titles = {}

    def get_spec_title(self, spec_number: str) -> str:
        """Returns the formal specification title (e.g. 'System architecture for the 5G System (5GS)')."""
        if not spec_number:
            return ""
        clean = re.sub(r'^(TS|TR)\s*', '', spec_number.strip(), flags=re.IGNORECASE)
        if clean in self._spec_titles:
            return self._spec_titles[clean]
        # Check base number without sub-part (e.g. 23.801-01 -> 23.801)
        base = clean.split('-')[0]
        if base in self._spec_titles:
            return self._spec_titles[base]
        # Fallback to PROTOCOL_SPEC_CATALOG if available
        try:
            from modules.nas.ui.nas_dialogs import PROTOCOL_SPEC_CATALOG
            if clean in PROTOCOL_SPEC_CATALOG:
                return PROTOCOL_SPEC_CATALOG[clean].get("full_title", "")
        except Exception:
            pass
        return ""

    def get_spec_short_name(self, spec_number: str) -> str:
        """Returns a concise acronym or short name for a spec (e.g. '5GS NAS', 'NR RRC', 'UAS', '5GS Architecture')."""
        if not spec_number:
            return ""
        clean = re.sub(r'^(TS|TR)\s*', '', spec_number.strip(), flags=re.IGNORECASE)
        # Check PROTOCOL_SPEC_CATALOG first for gold-standard curated telecom acronyms
        try:
            from modules.nas.ui.nas_dialogs import PROTOCOL_SPEC_CATALOG
            if clean in PROTOCOL_SPEC_CATALOG:
                return PROTOCOL_SPEC_CATALOG[clean].get("acronym", "")
        except Exception:
            pass

        title = self.get_spec_title(clean)
        if not title:
            return ""

        # Check for parenthesized acronym e.g. '(UAS)', '(5GS)', '(V2X)', '(ProSe)'
        acronyms = re.findall(r'\(([A-Z0-9\-_]{2,8})\)', title)
        for acr in acronyms:
            if acr.upper() not in ("TS", "TR", "STAGE 2", "STAGE 3", "REL"):
                return acr

        # Fallback to concise title (first phrase before semicolon)
        s = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', title, flags=re.IGNORECASE)
        parts = [p.strip() for p in s.split(';') if p.strip()]
        first_part = parts[0] if parts else title
        if len(first_part) > 22:
            return first_part[:21].rstrip() + "…"
        return first_part
