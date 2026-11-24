"""Docked QuickSpecSearchPanel widget composed from modular mixins."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_search_threads import SpecSearchQueryWorker
from .mixins import QuickSearchLogicMixin, QuickSearchUiSetupMixin


class QuickSpecSearchPanel(QuickSearchUiSetupMixin, QuickSearchLogicMixin, QWidget):
    """Compact sidebar quick search over indexed 3GPP clauses."""

    open_in_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_in_specs_requested = pyqtSignal(str, str)  # (spec, version)

    MAX_VISIBLE_RESULTS = 80
    DEBOUNCE_MS = 250

    def __init__(
        self,
        spec_search_db_path: Path,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("quickSpecSearchPanel")
        self.db_path = Path(spec_search_db_path)
        self.db = SpecSearchDatabase(self.db_path)

        self._versions: List[Dict[str, Any]] = []
        self._req_id: int = 0
        self._worker: Optional[SpecSearchQueryWorker] = None
        self._retired_workers: List[SpecSearchQueryWorker] = []
        self._selected_meta: Optional[Dict[str, Any]] = None
        self._selected_content: str = ""

        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(self.DEBOUNCE_MS)
        self._debounce_timer.timeout.connect(self._execute_search)

        self._setup_ui()
        self._apply_styles()
        self.refresh_versions()
