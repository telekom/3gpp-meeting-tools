"""Modeless Unified Clause Comparison & Analysis Dialog."""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QDialog, QWidget

from modules.spec_search.core.spec_search_db import SpecSearchDatabase

from .mixins import (
    _AttributionResultsMixin,
    _AttributionUiMixin,
    _DiffPreviewMixin,
    _DiffUiMixin,
)

logger = logging.getLogger(__name__)


class SpecClauseDiffDialog(
    _DiffUiMixin,
    _DiffPreviewMixin,
    _AttributionUiMixin,
    _AttributionResultsMixin,
    QDialog,
):
    """
    Unified Clause Comparison & Analysis Dialog.
    Features:
      1. Tab 1 - Visual Diff: In-app color-coded text diff (+ green, - red, @@ headers)
         with change statistics and 1-click clipboard copy.
      2. Tab 2 - AI Analysis Prompt: Multi-tier hierarchical context prompt generator
         tailored for LLM standards/patent/signalling analysis.
      3. Tab 3 - Modifying TDoc Attribution: Zero-download metadata ranking and
         Track-Changes OpenXML text fingerprinting identifying originating CR/p-CR.
      4. Footer - Word Track Changes: 1-click extraction and launch into Microsoft Word
         COM automation with native Track Changes (strikethrough & underline).
    """

    def __init__(
        self,
        db: SpecSearchDatabase,
        spec_number: str,
        current_version: str,
        clause_number: str,
        clause_title: str,
        cache_dir: Optional[Path] = None,
        specs_db: Optional[Any] = None,
        base_version: Optional[str] = None,
        focus_attribution: bool = False,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(None, Qt.Window)
        self.db = db
        self.spec_number = str(spec_number).strip()
        self.current_version = str(current_version).lstrip("v").strip()
        self.clause_number = str(clause_number).strip()
        self.clause_title = str(clause_title).strip()
        self.cache_dir = cache_dir
        self.specs_db = specs_db
        self.initial_base_version = str(base_version).lstrip("v").strip() if base_version else None
        self.focus_attribution = focus_attribution
        self._parent_ref = parent

        self._raw_diff_text = ""
        self._generated_prompt = ""
        self._word_cmp_worker = None

        self._attr_worker = None
        self._deep_verify_worker = None
        self._cached_attribution_results: Optional[List[Dict[str, Any]]] = None
        self._attribution_query_key = None
        self._candidate_list: List[Dict[str, Any]] = []

        self.setWindowTitle(f"🔍 Compare Clause {self.clause_number} — 3GPP TS {self.spec_number}")
        self.resize(1060, 750)
        self._setup_ui()
        self._load_available_versions()
        self._generate_diff_preview()

        if self.focus_attribution:
            self.tab_widget.setCurrentIndex(2)
            QTimer.singleShot(150, lambda: self._run_attribution(force=True))

    @property
    def version_b(self) -> str:
        if hasattr(self, "target_ver_combo") and self.target_ver_combo.currentData():
            return self.target_ver_combo.currentData().get("version", "")
        return getattr(self, "current_version", "")

    @property
    def version_a(self) -> str:
        if hasattr(self, "base_ver_combo") and self.base_ver_combo.currentData():
            return self.base_ver_combo.currentData().get("version", "")
        return getattr(self, "initial_base_version", "")

    def get_session_state(self) -> dict:
        """Return serializable SpecClauseDiffDialog context and tab selection."""
        return {
            "spec_number": str(self.spec_number or ""),
            "version_a": str(self.version_a or ""),
            "version_b": str(self.version_b or ""),
            "clause_number": str(self.clause_number or ""),
            "clause_title": str(self.clause_title or ""),
            "active_tab_idx": self.tab_widget.currentIndex() if hasattr(self, "tab_widget") else 0,
            "focus_idx": self.focus_combo.currentIndex() if hasattr(self, "focus_combo") else 0,
            "tier_id": self.tier_group.checkedId() if hasattr(self, "tier_group") else 2,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore active tab, analysis focus, and prompt context tier."""
        if not isinstance(state, dict):
            return
        focus_idx = state.get("focus_idx")
        if isinstance(focus_idx, int) and hasattr(self, "focus_combo"):
            if 0 <= focus_idx < self.focus_combo.count():
                self.focus_combo.setCurrentIndex(focus_idx)
        tier_id = state.get("tier_id")
        if isinstance(tier_id, int) and hasattr(self, "tier_group"):
            btn = self.tier_group.button(tier_id)
            if btn is not None:
                btn.setChecked(True)
        tab_idx = state.get("active_tab_idx")
        if isinstance(tab_idx, int) and hasattr(self, "tab_widget"):
            if 0 <= tab_idx < self.tab_widget.count():
                self.tab_widget.setCurrentIndex(tab_idx)
