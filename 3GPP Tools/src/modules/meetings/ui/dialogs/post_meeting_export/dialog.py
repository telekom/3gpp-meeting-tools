"""PostMeetingExportDialog shell combining layout, analysis, tree, and config mixins."""

import logging
from pathlib import Path
from typing import Set, Tuple

from PyQt5.QtWidgets import QDialog

from core.ui.theme import C
from modules.meetings.ui.dialogs.post_meeting_export.mixins import (
    PostMeetingAnalysisMixin,
    PostMeetingOutputConfigMixin,
    PostMeetingTreeSelectionMixin,
    PostMeetingUiSetupMixin,
)

logger = logging.getLogger("modules.meetings.ui.dialogs.post_meeting_export_dialog")


class PostMeetingExportDialog(
    PostMeetingUiSetupMixin,
    PostMeetingAnalysisMixin,
    PostMeetingTreeSelectionMixin,
    PostMeetingOutputConfigMixin,
    QDialog,
):
    """
    Configuration dialog to customize and trigger post-meeting Markdown dossiers.
    Supports scoping by Followed/Starred Work Items, filtered view, whole meeting,
    or custom Work Item & Agenda Item selection (with dual WI->AI and AI->WI tree views),
    with live scope estimation and per-release/per-WI directory organization.
    """

    def __init__(
        self,
        parent,
        meeting_dir: Path,
        all_rows: list,
        visible_rows: list,
        agenda_map: dict,
        mtg_info: dict,
        followed_tokens: set,
        starred_tokens: set,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.all_rows = all_rows or []
        self.visible_rows = visible_rows or []
        self.agenda_map = agenda_map or {}
        self.mtg_info = mtg_info or {}
        self.followed_tokens = set(followed_tokens or [])
        self.starred_tokens = set(starred_tokens or [])
        self.tracked_tokens = self.followed_tokens | self.starred_tokens

        self._updating_tree = False
        self.selected_wi_ais: Set[Tuple[str, str]] = set()

        self.setWindowTitle("📑 Post-Meeting Report & Markdown Dossier Export")
        self.resize(960, 650)
        self.setMinimumSize(880, 560)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            QGroupBox {{
                font-weight: bold;
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 14px;
                background-color: {C.BG_CARD};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                background-color: {C.BG_CARD};
            }}
            QLabel {{ color: {C.TEXT_BODY}; }}
            QRadioButton, QCheckBox {{
                color: {C.TEXT_STRONG};
                padding: 2px;
            }}
        """)

        # Extract available meeting Work Items, Agenda Items, and their cross-correlations
        self._analyze_meeting_contributions()

        # Pre-populate custom selection with Followed & Starred WIs automatically
        self._initialize_default_custom_selection()

        self._setup_ui()

        # Update initial tree check states and live estimation banner
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._expand_checked_parents(self.tree_wi)
        self._update_estimates()
