"""CrPcrDialog shell combining UI setup, clause tree/naming, and generation mixins."""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from core.ui.theme import C
from core.ui.ui_components import CARD_FRAME_STYLE
from modules.specifications.ui.threads import SpecExcerptPrepThread
from modules.word_tools.core.cr_pcr_generator import CrPcrGeneratorThread
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import HeadingInfo
from modules.word_tools.ui.cr_pcr.mixins import (
    CrPcrGenerationTemplateMixin,
    CrPcrTreeNamingMixin,
    CrPcrUiSetupMixin,
)

logger = logging.getLogger("modules.word_tools.ui.cr_pcr_dialog")


class CrPcrDialog(
    CrPcrUiSetupMixin,
    CrPcrTreeNamingMixin,
    CrPcrGenerationTemplateMixin,
    QDialog,
):
    """
    Comprehensive dialog for generating 3GPP CRs and pCRs.
    Can be launched from Specification tab or Meetings tab.
    """

    contribution_generated = pyqtSignal(Path)

    def __init__(
        self,
        spec_num: str = "",
        spec_type: str = "TS",
        version: str = "",
        release: str = "",
        fname: str = "",
        url: str = "",
        spec_dl_dir: Optional[Path] = None,
        ordered_paths: Optional[Sequence[str]] = None,
        headings: Optional[Sequence[HeadingInfo]] = None,
        meeting_info: Optional[Dict[str, str]] = None,
        tdoc_info: Optional[Dict[str, str]] = None,
        initial_mode: str = "pcr",
        parent=None,
        auto_start: bool = True,
        **kwargs,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top of main window)
        self._parent_ref = parent

        if "initial_spec" in kwargs and not spec_num:
            spec_num = kwargs.pop("initial_spec")
        if "mode" in kwargs:
            initial_mode = kwargs.pop("mode")
        if "initial_version" in kwargs and not version:
            version = kwargs.pop("initial_version")
        if "initial_file" in kwargs and not ordered_paths:
            f = kwargs.pop("initial_file")
            if f:
                ordered_paths = [f]

        self.spec_num = str(spec_num or "").strip()
        self.spec_type = (spec_type or "TS").upper().strip()
        self.version = str(version or "").strip()
        self.release = str(release or "").strip()
        self.fname = str(fname or "").strip()
        self.url = str(url or "").strip()
        self.spec_dl_dir = Path(spec_dl_dir) if spec_dl_dir else Path.cwd()
        self.meeting_info = meeting_info or {}
        self.tdoc_info = tdoc_info or {}
        self.current_mode = initial_mode.lower()

        self._prep_thread: Optional[SpecExcerptPrepThread] = None
        self._gen_thread: Optional[CrPcrGeneratorThread] = None

        self._ordered_paths: List[str] = list(ordered_paths) if ordered_paths else []
        self._headings: List[HeadingInfo] = list(headings) if headings else []
        self._all_new_map: Dict[str, bool] = {}
        self._is_updating_checks = False
        self._custom_filename_active = False

        cfg = WordConfig.load()
        self.default_company = cfg.get("default_source_company", "")

        self.setWindowTitle(
            f"📝 Draft 3GPP Contribution: {self.spec_type} {self.spec_num} v{self.version}"
        )
        # Modeless top-level independent window configuration:
        # Note: Must use Qt.Window rather than self.windowFlags() | Qt.Window so Qt.Dialog (0x3)
        # is cleared, allowing the window to freely move behind the main app window in z-order.
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent is not None:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.resize(1020, 780)
        self.setMinimumWidth(850)
        self.setMinimumHeight(600)

        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()

        # If headings are already provided, populate tree immediately
        if self._ordered_paths and self._headings:
            self._on_prep_finished(self._ordered_paths, self._headings)
        elif auto_start and self.spec_num and self.version:
            self._start_spec_preparation()

        self._update_preview_filename()

    @property
    def headings(self) -> List[HeadingInfo]:
        return self._headings

    @property
    def ordered_paths(self) -> List[str]:
        return self._ordered_paths

    def closeEvent(self, event):
        prep = getattr(self, "_prep_thread", None)
        if prep and prep.isRunning():
            prep.requestInterruption()
            if not prep.wait(1000):
                prep.terminate()
                prep.wait(500)
        gen = getattr(self, "_gen_thread", None)
        if gen and gen.isRunning():
            gen.requestInterruption()
            if not gen.wait(1000):
                gen.terminate()
                gen.wait(500)
        super().closeEvent(event)

    def get_session_state(self) -> dict:
        """Return serializable CrPcrDialog specification, meeting, and draft metadata."""
        return {
            "spec_num": str(self.spec_num or ""),
            "spec_type": str(self.spec_type or "TS"),
            "version": str(self.version or ""),
            "release": str(self.release or ""),
            "fname": str(self.fname or ""),
            "url": str(self.url or ""),
            "spec_dl_dir": str(self.spec_dl_dir) if self.spec_dl_dir else "",
            "mode": str(self.current_mode or "pcr"),
            "meeting_info": {
                "wg_name": self.input_wg.text() if hasattr(self, "input_wg") else "",
                "meeting_number": self.input_mtg_num.text() if hasattr(self, "input_mtg_num") else "",
                "location": self.input_location.text() if hasattr(self, "input_location") else "",
                "dates": self.input_dates.text() if hasattr(self, "input_dates") else "",
            },
            "tdoc_info": {
                "TDoc": self.input_tdoc_id.text() if hasattr(self, "input_tdoc_id") else "",
                "Source": self.input_source.text() if hasattr(self, "input_source") else "",
                "Title": self.input_title.text() if hasattr(self, "input_title") else "",
                "Work Item": self.input_work_item.text() if hasattr(self, "input_work_item") else "",
                "Agenda Item": self.input_ai.text() if hasattr(self, "input_ai") else "",
                "For": self.combo_doc_for.currentText() if hasattr(self, "combo_doc_for") else "",
                "Abstract": self.text_abstract.toPlainText() if hasattr(self, "text_abstract") else "",
            },
            "discussion": self.text_discussion.toPlainText() if hasattr(self, "text_discussion") else "",
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore mode, discussion text, and editable metadata fields."""
        if not isinstance(state, dict):
            return
        mode = str(state.get("mode") or "pcr").lower()
        if mode == "cr" and hasattr(self, "radio_cr"):
            self.radio_cr.setChecked(True)
        elif hasattr(self, "radio_pcr"):
            self.radio_pcr.setChecked(True)
        discussion = state.get("discussion")
        if isinstance(discussion, str) and discussion and hasattr(self, "text_discussion"):
            self.text_discussion.setPlainText(discussion)
