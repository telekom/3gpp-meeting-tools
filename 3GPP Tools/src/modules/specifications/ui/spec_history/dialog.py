"""Modeless Specification Change History & Integrated TDocs Viewer dialog."""

import logging
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QWidget

from core.ui.theme import C
from core.ui.ui_components import CARD_FRAME_STYLE
from core.utils.paths import get_project_root
from modules.specifications.core.spec_history_db import SpecHistoryDatabase

from .mixins import (
    _HistoryActionsMixin,
    _HistoryDataMixin,
    _HistoryDetailsMixin,
    _HistoryUiMixin,
)

logger = logging.getLogger(__name__)


class SpecVersionHistoryDialog(
    _HistoryUiMixin,
    _HistoryDataMixin,
    _HistoryDetailsMixin,
    _HistoryActionsMixin,
    QDialog,
):
    """Modern modal dialog inspecting all Change History and integrated TDocs for a specification."""

    def __init__(
        self,
        spec_number: str,
        initial_version: Optional[str] = None,
        parent: Optional[QWidget] = None,
        history_db: Optional[SpecHistoryDatabase] = None,
        specs_db: Optional[Any] = None,
        initial_search: str = "",
        initial_clause: str = "",
        sort_by_origin_relevance: bool = False,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top)
        self._parent_ref = parent
        self.spec_number = spec_number.strip()
        self.initial_version = initial_version.lstrip("vV").strip() if initial_version else ""
        self.initial_search = initial_search.strip()
        self.initial_clause = initial_clause.strip()
        self.sort_by_origin_relevance = sort_by_origin_relevance
        self.history_db = history_db or SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
        self.specs_db = specs_db
        self._attribution_engine = None

        # Modeless top-level independent window configuration
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass

        self.setWindowTitle(f"Change History & Integrated TDocs — {self.spec_number}")
        self.resize(1060, 680)
        self.setMinimumSize(850, 500)

        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            {CARD_FRAME_STYLE}
        """)

        self.all_records: List[Dict[str, Any]] = []
        self.filtered_records: List[Dict[str, Any]] = []
        self._active_report_thread = None
        self._report_dialog = None

        logger.info(f"📜 [Change History] Opened Change History & TDocs browser for TS {self.spec_number}{f' (v{self.initial_version})' if self.initial_version else ''}.")
        self._setup_ui()
        self._load_data(auto_extract_if_empty=True)

    def load_spec(
        self,
        spec_number: str,
        version: Optional[str] = None,
        initial_search: str = "",
        initial_clause: str = "",
        sort_by_origin_relevance: bool = False
    ):
        """Switches the active specification and reloads change history records."""
        self.spec_number = spec_number.strip()
        self.initial_version = version.lstrip("vV").strip() if version else ""
        if initial_search:
            self.initial_search = initial_search.strip()
        if initial_clause:
            self.initial_clause = initial_clause.strip()
        if sort_by_origin_relevance:
            self.sort_by_origin_relevance = sort_by_origin_relevance
        self.setWindowTitle(f"Change History & Integrated TDocs — {self.spec_number}")

        spec_title = "Specification Change History"
        if self.specs_db and hasattr(self.specs_db, "get_spec_details"):
            try:
                details = self.specs_db.get_spec_details(self.spec_number)
                if details and details.get("title"):
                    spec_title = details["title"]
            except Exception:
                pass
        self.title_lbl.setText(spec_title)
        self.spec_badge.setText(f"<b>{self.spec_number}</b>")
        logger.info(f"📜 [Change History] Reloading browser for TS {self.spec_number}{f' (v{self.initial_version})' if self.initial_version else ''}.")
        self._load_data(auto_extract_if_empty=True)

    def get_session_state(self) -> dict:
        """Return serializable SpecVersionHistoryDialog context and filters."""
        ver = ""
        if hasattr(self, "version_combo"):
            ver = str(self.version_combo.currentData() or "")
        return {
            "spec_number": str(self.spec_number or ""),
            "version": ver,
            "search_text": self.search_input.text() if hasattr(self, "search_input") else "",
            "clause_filter_enabled": bool(self.chk_clause_filter.isChecked()) if hasattr(self, "chk_clause_filter") else False,
            "clause_text": self.clause_input.text() if hasattr(self, "clause_input") else "",
            "clause_mode_idx": self.clause_mode_combo.currentIndex() if hasattr(self, "clause_mode_combo") else 0,
            "origin_ranking": bool(self.chk_origin_ranking.isChecked()) if hasattr(self, "chk_origin_ranking") else False,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore selected version, search query, clause mode, and origin ranking."""
        if not isinstance(state, dict):
            return
        ver = str(state.get("version") or "").strip()
        if hasattr(self, "version_combo"):
            idx = self.version_combo.findData(ver)
            if idx >= 0:
                self.version_combo.setCurrentIndex(idx)
        if isinstance(state.get("search_text"), str) and hasattr(self, "search_input"):
            self.search_input.setText(state["search_text"])
        if isinstance(state.get("clause_text"), str) and hasattr(self, "clause_input"):
            self.clause_input.setText(state["clause_text"])
        if "clause_filter_enabled" in state and hasattr(self, "chk_clause_filter"):
            self.chk_clause_filter.setChecked(bool(state["clause_filter_enabled"]))
        mode_idx = state.get("clause_mode_idx")
        if isinstance(mode_idx, int) and hasattr(self, "clause_mode_combo"):
            if 0 <= mode_idx < self.clause_mode_combo.count():
                self.clause_mode_combo.setCurrentIndex(mode_idx)
        if "origin_ranking" in state and hasattr(self, "chk_origin_ranking"):
            self.chk_origin_ranking.setChecked(bool(state["origin_ranking"]))
        self._apply_filters()
