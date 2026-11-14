# --- File: src/modules/meetings/ui/dialogs/topic_dossier_dialog.py ---
"""
Topic Dossier & Briefing Studio Dialog (`TopicDossierDialog`).

Modeless 2-column studio allowing delegates to manage cross-WG Meta-Topics,
select quarterly meeting cycles, preview scope across Working Groups,
and perform incremental or full Markdown dossier exports.
"""

import datetime
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, QSize, QDate
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QGroupBox,
    QRadioButton, QCheckBox, QPushButton, QLineEdit, QFileDialog,
    QListWidget, QListWidgetItem, QStackedWidget, QFrame, QScrollArea,
    QWidget, QButtonGroup, QMessageBox, QProgressBar, QComboBox, QTextEdit,
    QDateEdit,
)

from core.ui.flow_layout import FlowLayout
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.paths import get_project_root, to_extended_path, strip_extended_path
from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    MeetingCohortResolver,
    format_cohort_label,
)
from modules.meetings.core.exporters.topic_dossier_exporter import (
    TopicDossierExporterThread,
    get_default_topic_export_base,
    sanitize_folder_name,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import (
    MetaTopicDefinition,
    MetaTopicManager,
    MetaTopicRules,
)
from modules.meetings.ui.topic_pills import (
    WgTogglePill,
    TagCloudWidget,
    AgendaItemPillsWidget,
)
from core.ui.theme import C

ALL_COMMON_WGS = ["SA2", "RAN2", "RAN3", "CT1", "RAN1", "SA3", "CT3", "CT4", "SA1", "SA4", "SA5", "RAN4"]

STYLE_CARD = f"""
    QFrame#CardFrame {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 8px;
    }}
"""

STYLE_PRIMARY_BTN = f"""
    QPushButton {{
        background-color: {C.EMERALD_ACCENT};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.EMERALD_PRIMARY};
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 12px;
    }}
    QPushButton:hover {{
        background-color: {C.EMERALD_PRIMARY};
    }}
    QPushButton:disabled {{
        background-color: {C.TEXT_FAINT};
        border-color: {C.TEXT_FAINT};
    }}
"""

STYLE_SECONDARY_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:disabled {{
        color: {C.TEXT_FAINT};
        background-color: {C.BG_SURFACE};
        border-color: {C.BORDER_SUBTLE};
    }}
"""

STYLE_MODE_ACTIVE = f"""
    QPushButton {{
        background-color: {C.ACCENT_ALT};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.INFO};
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 11px;
    }}
"""

STYLE_MODE_INACTIVE = f"""
    QPushButton {{
        background-color: {C.BG_CARD};
        color: {C.TEXT_SECONDARY};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 11px;
    }}
    QPushButton:hover {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_PRIMARY};
    }}
"""


class TopicEditorDialog(QDialog):
    """Sub-dialog to create or edit a MetaTopicDefinition."""

    def __init__(
        self,
        topic: Optional[MetaTopicDefinition] = None,
        meetings_db: Optional[MeetingsDatabase] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle("Edit Meta-Topic Definition" if topic else "New Meta-Topic")
        self.resize(700, 660)
        self.topic = topic
        self.meetings_db = meetings_db
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Basic Info
        grid = QGridLayout()
        grid.addWidget(QLabel("<b>Topic Name:</b>"), 0, 0)
        self.edit_name = QLineEdit(self.topic.name if self.topic else "")
        self.edit_name.setPlaceholderText("e.g. Ambient IoT (Rel-19)")
        grid.addWidget(self.edit_name, 0, 1)

        grid.addWidget(QLabel("<b>Description:</b>"), 1, 0)
        self.edit_desc = QLineEdit(self.topic.description if self.topic else "")
        self.edit_desc.setPlaceholderText("Brief purpose or architectural scope...")
        grid.addWidget(self.edit_desc, 1, 1)
        layout.addLayout(grid)

        # Target WGs (Toggle Pills)
        wg_box = QGroupBox("Target Working Groups (Click to Toggle)")
        wg_layout = FlowLayout(wg_box, margin=8, h_spacing=6, v_spacing=6)
        self.wg_checks = {}
        target_set = set(self.topic.target_wgs) if self.topic else {"SA2", "RAN2", "CT1"}
        for wg in ALL_COMMON_WGS:
            pill = WgTogglePill(wg, checked=(wg in target_set))
            pill.toggled.connect(self._on_wg_toggled)
            self.wg_checks[wg] = pill
            wg_layout.addWidget(pill)
        layout.addWidget(wg_box)

        # Rules Group
        rules_box = QGroupBox("Matching Rules (Auto-Discovery)")
        r_layout = QVBoxLayout(rules_box)
        r_layout.setSpacing(8)

        r_layout.addWidget(QLabel("Work Item / Study Item Tokens (Tag Cloud):"))
        self.edit_wi = TagCloudWidget(placeholder="Type WI/SI token and press Enter or comma (e.g. ambientiot, fs_6g_df)...")
        if self.topic and self.topic.rules.wi_tokens:
            self.edit_wi.setText(", ".join(self.topic.rules.wi_tokens))
        r_layout.addWidget(self.edit_wi)

        ai_header_layout = QHBoxLayout()
        ai_header_layout.addWidget(QLabel("Agenda Item Patterns per Working Group:"))
        ai_header_layout.addStretch()
        btn_suggest = QPushButton("🔍 Auto-Suggest AIs...")
        btn_suggest.setToolTip("Auto-discover and recommend matching Agenda Items across target Working Groups")
        btn_suggest.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_suggest.clicked.connect(self._on_suggest_ais)
        ai_header_layout.addWidget(btn_suggest)
        r_layout.addLayout(ai_header_layout)

        self.edit_ai = AgendaItemPillsWidget()
        active_wgs = [wg for wg, p in self.wg_checks.items() if p.isChecked()]
        self.edit_ai.set_active_wgs(active_wgs)

        ai_lines = []
        if self.topic and self.topic.rules.ai_patterns:
            for wg, pats in self.topic.rules.ai_patterns.items():
                ai_lines.append(f"{wg}: {', '.join(pats)}")
        self.edit_ai.setText("; ".join(ai_lines))
        r_layout.addWidget(self.edit_ai)

        r_layout.addWidget(QLabel("Keywords in Title / Abstract (Tag Cloud):"))
        self.edit_kw = TagCloudWidget(placeholder="Type keyword and press Enter or comma (e.g. data framework, zero power)...")
        if self.topic and self.topic.rules.keywords:
            self.edit_kw.setText(", ".join(self.topic.rules.keywords))
        r_layout.addWidget(self.edit_kw)

        layout.addWidget(rules_box)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_save = QPushButton("Save Topic")
        btn_save.setStyleSheet(STYLE_PRIMARY_BTN)
        btn_save.clicked.connect(self._on_save)
        btn_layout.addWidget(btn_save)
        layout.addLayout(btn_layout)

    def _on_wg_toggled(self, checked: bool):
        active_wgs = [wg for wg, p in self.wg_checks.items() if p.isChecked()]
        self.edit_ai.set_active_wgs(active_wgs)

    def _on_suggest_ais(self):
        """Opens AiSuggestionDialog and applies recommended AIs."""
        from modules.meetings.core.ai_suggester import AgendaItemSuggester
        from modules.meetings.ui.dialogs.ai_suggestion_dialog import AiSuggestionDialog

        selected_wgs = [wg for wg, cb in self.wg_checks.items() if cb.isChecked()]
        if not selected_wgs:
            QMessageBox.warning(self, "No Working Groups", "Please select at least one Target Working Group first.")
            return

        if not self.meetings_db:
            from core.utils.paths import get_project_root
            db_path = get_project_root() / "db" / "3gpp_data.db"
            self.meetings_db = MeetingsDatabase(db_path)

        suggester = AgendaItemSuggester(self.meetings_db)

        # 1. Parse any existing AI patterns entered by the user (e.g. SA2: 20.6.21)
        known_patterns: Dict[str, List[str]] = {}
        cur_ai_text = self.edit_ai.text().strip()
        if cur_ai_text:
            for chunk in cur_ai_text.split(";"):
                chunk = chunk.strip()
                if ":" in chunk:
                    wg_k, pats = chunk.split(":", 1)
                    wg_k = wg_k.strip().upper()
                    p_list = [p.strip() for p in pats.split(",") if p.strip()]
                    if p_list:
                        known_patterns[wg_k] = p_list

        # 2. If known AI patterns exist, infer topic metadata (descriptions, keywords, WI tokens)
        inferred_summary = None
        inferred_kws: List[str] = []
        inferred_wis: List[str] = []
        if known_patterns:
            inferred = suggester.infer_topic_from_known_ais(known_patterns)
            inferred_summary = inferred.get("source_summary")
            inferred_kws = inferred.get("keywords", [])
            inferred_wis = inferred.get("wi_tokens", [])

        # 3. Gather user-specified keywords and WI tokens, supplementing with inferred ones
        kws = [k.strip() for k in self.edit_kw.text().split(",") if k.strip()]
        wis = [w.strip() for w in self.edit_wi.text().split(",") if w.strip()]

        for ik in inferred_kws:
            if ik.lower() not in [k.lower() for k in kws]:
                kws.append(ik)
        for iw in inferred_wis:
            if iw.lower() not in [w.lower() for w in wis]:
                wis.append(iw)

        if not kws and not wis:
            name = self.edit_name.text().strip()
            if name:
                clean_name = re.sub(r'\(.*?\)', '', name).strip()
                kws = [p.strip() for p in re.split(r'\s*(?:&|and)\s*', clean_name, flags=re.IGNORECASE) if p.strip()]

        if not kws and not wis:
            QMessageBox.information(
                self,
                "Keywords Needed",
                "Please enter keywords in 'Keywords in Title / Abstract', 'Topic Name', or specify an AI for at least one WG (e.g. SA2: 20.6.21).",
            )
            return

        dlg = AiSuggestionDialog(
            meetings_db=self.meetings_db,
            target_wgs=selected_wgs,
            initial_keywords=kws,
            initial_wis=wis,
            inferred_summary=inferred_summary,
            parent=self,
        )

        if dlg.exec_() == QDialog.Accepted:
            new_rule = dlg.get_formatted_rule_string()
            if new_rule:
                self._merge_ai_rules(new_rule)
            # Auto-populate empty keyword / WI fields with discovered terms
            active_kws = dlg.get_active_keywords()
            if not self.edit_kw.text().strip() and active_kws:
                self.edit_kw.setText(", ".join(active_kws))
            active_wis = dlg.get_active_wis()
            if not self.edit_wi.text().strip() and active_wis:
                self.edit_wi.setText(", ".join(active_wis))

    def _merge_ai_rules(self, new_rule_str: str):
        """Merges discovered WG AI rules with existing ones, avoiding duplicates."""
        existing_map: Dict[str, List[str]] = {}
        current_text = self.edit_ai.text().strip()
        if current_text:
            for chunk in current_text.split(";"):
                chunk = chunk.strip()
                if ":" in chunk:
                    wg, pats = chunk.split(":", 1)
                    wg = wg.strip().upper()
                    p_list = [p.strip() for p in pats.split(",") if p.strip()]
                    existing_map[wg] = p_list

        for chunk in new_rule_str.split(";"):
            chunk = chunk.strip()
            if ":" in chunk:
                wg, pats = chunk.split(":", 1)
                wg = wg.strip().upper()
                p_list = [p.strip() for p in pats.split(",") if p.strip()]
                if wg not in existing_map:
                    existing_map[wg] = []
                for p in p_list:
                    if p not in existing_map[wg]:
                        existing_map[wg].append(p)

        out_parts = []
        for wg, pats in existing_map.items():
            if pats:
                out_parts.append(f"{wg}: {', '.join(pats)}")
        self.edit_ai.setText("; ".join(out_parts))

    def _on_save(self):
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Invalid Name", "Please provide a valid topic name.")
            return

        wgs = [wg for wg, cb in self.wg_checks.items() if cb.isChecked()]
        if not wgs:
            QMessageBox.warning(self, "No Working Groups", "Please select at least one Working Group.")
            return

        # Parse WI tokens
        wi_tokens = [w.strip().lower() for w in self.edit_wi.text().split(",") if w.strip()]

        # Parse AI patterns
        ai_patterns = {}
        for chunk in self.edit_ai.text().split(";"):
            chunk = chunk.strip()
            if ":" in chunk:
                wg_k, pats = chunk.split(":", 1)
                wg_k = wg_k.strip().upper()
                p_list = [p.strip().lower() for p in pats.split(",") if p.strip()]
                if p_list:
                    ai_patterns[wg_k] = p_list

        keywords = [k.strip().lower() for k in self.edit_kw.text().split(",") if k.strip()]

        rules = MetaTopicRules(
            wi_tokens=wi_tokens,
            ai_patterns=ai_patterns,
            keywords=keywords,
            pinned_tdocs=self.topic.rules.pinned_tdocs if self.topic else [],
            excluded_tdocs=self.topic.rules.excluded_tdocs if self.topic else [],
        )

        topic_id = self.topic.id if self.topic else re.sub(r'[^a-zA-Z0-9_-]+', '_', name.lower()).strip('_')
        self.result_topic = MetaTopicDefinition(
            id=topic_id,
            name=name,
            description=self.edit_desc.text().strip(),
            color=self.topic.color if self.topic else C.EMERALD_ACCENT,
            target_wgs=wgs,
            cycle_mode=self.topic.cycle_mode if self.topic else "past_n_cycles",
            past_n_cycles=self.topic.past_n_cycles if self.topic else 4,
            from_cycle=self.topic.from_cycle if self.topic else "2025-Q1",
            to_cycle=self.topic.to_cycle if self.topic else "2026-Q1",
            date_from=self.topic.date_from if self.topic else "",
            date_to=self.topic.date_to if self.topic else "",
            rules=rules,
        )
        self.accept()


class TopicDossierDialog(QDialog):
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
        self.scope_mode = "pre_meeting"  # "pre_meeting", "post_meeting", "full_dossier"
        self._worker: Optional[TopicDossierExporterThread] = None

        self._init_ui()
        self._load_topics_list()

    def _init_ui(self):
        root_layout = QHBoxLayout(self)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        # ==========================================
        # LEFT PANE: Topics Navigation & Management
        # ==========================================
        left_frame = QFrame()
        left_frame.setFixedWidth(280)
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        lbl_topics = QLabel("<b>🗂️ Meta-Topics</b>")
        lbl_topics.setStyleSheet(f"font-size: 13px; color: {C.TEXT_STRONG};")
        left_layout.addWidget(lbl_topics)

        self.list_topics = QListWidget()
        self.list_topics.setStyleSheet(f"""
            QListWidget {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px;
            }}
            QListWidget::item {{
                padding: 8px 10px;
                border-radius: 4px;
                margin-bottom: 2px;
            }}
            QListWidget::item:selected {{
                background-color: {C.INFO_SOFT};
                color: {C.INFO};
                font-weight: bold;
                border: 1px solid {C.INFO_BORDER};
            }}
        """)
        self.list_topics.currentRowChanged.connect(self._on_topic_selected)
        left_layout.addWidget(self.list_topics)

        # Toolbar under topics list
        t_btn_layout = QGridLayout()
        t_btn_layout.setSpacing(4)
        self.btn_new = QPushButton("➕ New")
        self.btn_new.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_new.clicked.connect(self._on_new_topic)
        t_btn_layout.addWidget(self.btn_new, 0, 0)

        self.btn_edit = QPushButton("✏️ Edit")
        self.btn_edit.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_edit.clicked.connect(self._on_edit_topic)
        t_btn_layout.addWidget(self.btn_edit, 0, 1)

        self.btn_clone = QPushButton("📋 Clone")
        self.btn_clone.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_clone.clicked.connect(self._on_clone_topic)
        t_btn_layout.addWidget(self.btn_clone, 1, 0)

        self.btn_del = QPushButton("🗑️ Delete")
        self.btn_del.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_del.clicked.connect(self._on_delete_topic)
        t_btn_layout.addWidget(self.btn_del, 1, 1)
        left_layout.addLayout(t_btn_layout)

        self.btn_restore = QPushButton("🔄 Restore Built-in Topics...")
        self.btn_restore.setToolTip("Restore standard 3GPP topics (Ambient IoT, AI/ML Phase 2, 6G Architecture, Network Energy Savings, Data Framework) without overwriting custom topics.")
        self.btn_restore.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_restore.clicked.connect(self._on_restore_defaults)
        left_layout.addWidget(self.btn_restore)

        root_layout.addWidget(left_frame)

        # ==========================================
        # RIGHT PANE: Topic Runner & Options
        # ==========================================
        right_frame = QFrame()
        right_frame.setObjectName("CardFrame")
        right_frame.setStyleSheet(STYLE_CARD)
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        # 1. Topic Header
        self.lbl_topic_title = QLabel("Select a Topic")
        self.lbl_topic_title.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        right_layout.addWidget(self.lbl_topic_title)

        self.lbl_topic_desc = QLabel("Description...")
        self.lbl_topic_desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        self.lbl_topic_desc.setWordWrap(True)
        right_layout.addWidget(self.lbl_topic_desc)

        # 2. Scope Mode Segmented Bar
        mode_box = QGroupBox("Temporal Perspective / Scope Mode")
        mode_layout = QHBoxLayout(mode_box)
        self.btn_mode_pre = QPushButton("🧭 Pre-Meeting Briefing\n(Where are we? + What is coming?)")
        self.btn_mode_pre.clicked.connect(lambda: self._set_scope_mode("pre_meeting"))
        self.btn_mode_post = QPushButton("📑 Post-Meeting Report\n(What was agreed/decided?)")
        self.btn_mode_post.clicked.connect(lambda: self._set_scope_mode("post_meeting"))
        self.btn_mode_full = QPushButton("📚 Complete Feature Dossier\n(Full baseline + all documents)")
        self.btn_mode_full.clicked.connect(lambda: self._set_scope_mode("full_dossier"))
        mode_layout.addWidget(self.btn_mode_pre)
        mode_layout.addWidget(self.btn_mode_post)
        mode_layout.addWidget(self.btn_mode_full)
        right_layout.addWidget(mode_box)

        # 3. Meeting Scope & Timeframe Picker
        cycle_box = QGroupBox("Meeting Scope && Timeframe")
        cycle_layout = QGridLayout(cycle_box)

        cycle_layout.addWidget(QLabel("<b>Meeting Timeframe:</b>"), 0, 0)
        self.combo_cycle_window = QComboBox()
        self.combo_cycle_window.addItems([
            "Past 4 Meeting Rounds (~1 Year)",
            "Past 6 Meeting Rounds (~1.5 Years)",
            "Past 8 Meeting Rounds (~2 Years)",
            "Release 20 Timeframe (2025–2027)",
            "Release 19 Timeframe (2023–2025)",
            "Release 18 Timeframe (2021–2023)",
            "📅 Custom Date Range...",
            "All Available Meetings",
        ])
        self.combo_cycle_window.currentIndexChanged.connect(self._on_timeframe_preset_changed)
        cycle_layout.addWidget(self.combo_cycle_window, 0, 1)

        # Date range picker container (shown when "📅 Custom Date Range..." is selected)
        self.date_range_widget = QWidget()
        dr_layout = QHBoxLayout(self.date_range_widget)
        dr_layout.setContentsMargins(0, 2, 0, 2)
        dr_layout.setSpacing(8)

        dr_layout.addWidget(QLabel("From:"))
        self.edit_date_from = QDateEdit()
        self.edit_date_from.setCalendarPopup(True)
        self.edit_date_from.setDisplayFormat("yyyy-MM-dd")
        self.edit_date_from.setDate(QDate.currentDate().addYears(-1))
        self.edit_date_from.dateChanged.connect(self._update_cohorts_preview)
        dr_layout.addWidget(self.edit_date_from)

        dr_layout.addWidget(QLabel("To:"))
        self.edit_date_to = QDateEdit()
        self.edit_date_to.setCalendarPopup(True)
        self.edit_date_to.setDisplayFormat("yyyy-MM-dd")
        self.edit_date_to.setDate(QDate.currentDate().addYears(1))
        self.edit_date_to.dateChanged.connect(self._update_cohorts_preview)
        dr_layout.addWidget(self.edit_date_to)
        dr_layout.addStretch()

        self.date_range_widget.setVisible(False)
        cycle_layout.addWidget(self.date_range_widget, 1, 1)

        self.lbl_cohorts_summary = QLabel("Covering: SA2, RAN2, CT1 across 4 meeting rounds")
        self.lbl_cohorts_summary.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        cycle_layout.addWidget(self.lbl_cohorts_summary, 2, 0, 1, 2)
        right_layout.addWidget(cycle_box)

        # 4. Target Directory & Incremental State
        dest_box = QGroupBox("Target Folder && Incremental Updating")
        dest_layout = QGridLayout(dest_box)

        dest_layout.addWidget(QLabel("<b>Destination:</b>"), 0, 0)
        self.edit_dest_folder = QLineEdit()
        self.edit_dest_folder.textChanged.connect(self._check_incremental_status)
        dest_layout.addWidget(self.edit_dest_folder, 0, 1)

        btn_browse = QPushButton("Browse...")
        btn_browse.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_browse.clicked.connect(self._on_browse_dest)
        dest_layout.addWidget(btn_browse, 0, 2)

        self.lbl_inc_status = QLabel("🟢 New Folder")
        self.lbl_inc_status.setStyleSheet("font-size: 11px; font-weight: bold;")
        dest_layout.addWidget(self.lbl_inc_status, 1, 0, 1, 3)
        right_layout.addWidget(dest_box)

        # 5. Conversion Options
        opt_box = QGroupBox("Options")
        opt_layout = QHBoxLayout(opt_box)
        self.chk_discussions = QCheckBox("Include Discussion Papers && WFs")
        self.chk_discussions.setChecked(True)
        opt_layout.addWidget(self.chk_discussions)

        self.chk_dedup = QCheckBox("Deduplicate Revisions")
        self.chk_dedup.setChecked(True)
        opt_layout.addWidget(self.chk_dedup)

        self.chk_consolidated = QCheckBox("Generate Consolidated Single .md")
        self.chk_consolidated.setChecked(True)
        opt_layout.addWidget(self.chk_consolidated)

        self.chk_download = QCheckBox("Download missing documents from 3GPP FTP")
        self.chk_download.setChecked(True)
        self.chk_download.setToolTip("Always downloads missing TDoc documents and meeting archives from 3GPP FTP to ensure complete dossier export.")
        opt_layout.addWidget(self.chk_download)
        right_layout.addWidget(opt_box)

        # 6. Progress & Actions
        self.prog_bar = QProgressBar()
        self.prog_bar.setVisible(False)
        self.prog_bar.setFixedHeight(12)
        right_layout.addWidget(self.prog_bar)

        self.lbl_status = QLabel("Ready")
        self.lbl_status.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        right_layout.addWidget(self.lbl_status)

        action_layout = QHBoxLayout()
        action_layout.addStretch()

        self.btn_view_inputs = QPushButton("👁️ View Considered Inputs...")
        self.btn_view_inputs.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_view_inputs.setToolTip("Open interactive visual inspector of all candidate decisions, pCRs, CRs, agreements, and liaison statements")
        self.btn_view_inputs.clicked.connect(self._on_view_considered_inputs)
        action_layout.addWidget(self.btn_view_inputs)

        self.btn_open_folder = QPushButton("📂 Open Output Folder")
        self.btn_open_folder.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self._on_open_output)
        action_layout.addWidget(self.btn_open_folder)

        self.btn_export = QPushButton("🚀 Export / Update Topic Dossier")
        self.btn_export.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_export.clicked.connect(self._on_start_export)
        action_layout.addWidget(self.btn_export)

        right_layout.addLayout(action_layout)
        root_layout.addWidget(right_frame)

        self._set_scope_mode("pre_meeting")

    def _load_topics_list(self, preferred_id: Optional[str] = None, preferred_index: Optional[int] = None):
        self.list_topics.blockSignals(True)
        self.list_topics.clear()
        topics = self.topic_manager.list_topics()
        selected_row = -1
        for idx, t in enumerate(topics):
            item = QListWidgetItem(f"{t.name}")
            item.setToolTip(f"{t.description}\nWGs: {', '.join(t.target_wgs)}")
            self.list_topics.addItem(item)
            if preferred_id and t.id == preferred_id:
                selected_row = idx

        self.list_topics.blockSignals(False)

        if topics:
            self.btn_edit.setEnabled(True)
            self.btn_clone.setEnabled(True)
            self.btn_del.setEnabled(True)
            self.btn_export.setEnabled(True)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(True)
            if selected_row >= 0:
                self.list_topics.setCurrentRow(selected_row)
            elif preferred_index is not None and 0 <= preferred_index < len(topics):
                self.list_topics.setCurrentRow(preferred_index)
            else:
                self.list_topics.setCurrentRow(0)
        else:
            self.active_topic = None
            self.lbl_topic_title.setText("No Topics Defined")
            self.lbl_topic_desc.setText("Click '➕ New' or '🔄 Restore Built-in Topics...' to get started.")
            self.edit_dest_folder.clear()
            self.selected_cohorts = []
            self.btn_edit.setEnabled(False)
            self.btn_clone.setEnabled(False)
            self.btn_del.setEnabled(False)
            self.btn_export.setEnabled(False)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(False)
            self.lbl_cohorts_summary.setText("No active topic selected.")

    def _on_topic_selected(self, row: int):
        topics = self.topic_manager.list_topics()
        if 0 <= row < len(topics):
            self.active_topic = topics[row]
            self.lbl_topic_title.setText(f"🌐 {self.active_topic.name}")
            wgs_str = ", ".join(self.active_topic.target_wgs)
            self.lbl_topic_desc.setText(f"<b>WGs:</b> {wgs_str} | {self.active_topic.description}")

            # Update default destination in user meeting cache
            base_dir = get_default_topic_export_base()
            default_name = f"Topic_{sanitize_folder_name(self.active_topic.name)}"
            self.edit_dest_folder.setText(str(base_dir / default_name))

            self._update_cohorts_preview()
            self._check_incremental_status()
            self.btn_edit.setEnabled(True)
            self.btn_clone.setEnabled(True)
            self.btn_del.setEnabled(True)
            self.btn_export.setEnabled(True)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(True)
        else:
            self.active_topic = None
            self.btn_edit.setEnabled(False)
            self.btn_clone.setEnabled(False)
            self.btn_del.setEnabled(False)
            self.btn_export.setEnabled(False)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(False)

    def _set_scope_mode(self, mode: str):
        self.scope_mode = mode
        self.btn_mode_pre.setStyleSheet(STYLE_MODE_ACTIVE if mode == "pre_meeting" else STYLE_MODE_INACTIVE)
        self.btn_mode_post.setStyleSheet(STYLE_MODE_ACTIVE if mode == "post_meeting" else STYLE_MODE_INACTIVE)
        self.btn_mode_full.setStyleSheet(STYLE_MODE_ACTIVE if mode == "full_dossier" else STYLE_MODE_INACTIVE)

    def _on_timeframe_preset_changed(self, idx: int):
        is_custom = (idx == 6)  # "📅 Custom Date Range..."
        self.date_range_widget.setVisible(is_custom)
        self._update_cohorts_preview()

    def _update_cohorts_preview(self):
        if not self.active_topic:
            return

        idx = self.combo_cycle_window.currentIndex()
        topic_copy = MetaTopicDefinition.from_dict(self.active_topic.to_dict())

        if idx == 0:  # Past 4 Meeting Rounds (~1 Year)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 4
        elif idx == 1:  # Past 6 Meeting Rounds (~1.5 Years)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 6
        elif idx == 2:  # Past 8 Meeting Rounds (~2 Years)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 8
        elif idx == 3:  # Release 20 Timeframe (2025–2027)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2025-01-01"
            topic_copy.date_to = "2027-12-31"
        elif idx == 4:  # Release 19 Timeframe (2023–2025)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2023-01-01"
            topic_copy.date_to = "2025-12-31"
        elif idx == 5:  # Release 18 Timeframe (2021–2023)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2021-01-01"
            topic_copy.date_to = "2023-12-31"
        elif idx == 6:  # 📅 Custom Date Range...
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = self.edit_date_from.date().toString("yyyy-MM-dd")
            topic_copy.date_to = self.edit_date_to.date().toString("yyyy-MM-dd")
        else:  # All Available Meetings
            topic_copy.cycle_mode = "cycle_range"
            topic_copy.from_cycle = "2018-Q1"
            topic_copy.to_cycle = "2035-Q4"

        self.selected_cohorts = MeetingCohortResolver.resolve_cohorts_for_topic(self.meetings_db, topic_copy)
        if self.selected_cohorts:
            first_c = self.selected_cohorts[0].label
            last_c = self.selected_cohorts[-1].label
            total_mtgs = sum(len(c.meetings) for c in self.selected_cohorts)
            mtg_names = []
            for c in self.selected_cohorts[-2:]:
                for m in c.meetings:
                    mtg_names.append(f"{m.get('wg_name')} #{m.get('meeting_number')}")
            preview_str = ", ".join(mtg_names[:6])
            if len(mtg_names) > 6:
                preview_str += "..."
            self.lbl_cohorts_summary.setText(
                f"Resolved: <b>{len(self.selected_cohorts)} Synchronized Meeting Round(s)</b> ({first_c} → {last_c}) with <b>{total_mtgs} meetings</b> across {', '.join(sorted(self.active_topic.target_wgs))}.<br/>"
                f"<span style='color: {C.TEXT_MUTED};'>Recent meetings: {preview_str}</span>"
            )
        else:
            self.lbl_cohorts_summary.setText("No meetings found matching this timeframe and target WGs.")

    def _check_incremental_status(self):
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            self.lbl_inc_status.setText("⚠️ Please specify a destination folder.")
            self.btn_open_folder.setEnabled(False)
            self.btn_open_folder.setToolTip("Destination folder is empty.")
            return

        folder_exists = os.path.isdir(to_extended_path(folder_str))
        self.btn_open_folder.setEnabled(folder_exists)
        self.btn_open_folder.setToolTip(
            f"Open in File Explorer: {folder_str}"
            if folder_exists
            else "Destination folder does not exist yet. Run an export first."
        )

        p = Path(folder_str)
        manifest_p = p / ".export_manifest.json"
        if manifest_p.is_file():
            self.lbl_inc_status.setText(
                "🔄 <b>Existing Corpus Detected:</b> Incremental update mode active (skips already exported TDocs, writes delta feed)."
            )
            self.lbl_inc_status.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px;")
        else:
            self.lbl_inc_status.setText("🆕 <b>New Directory:</b> Initial complete export.")
            self.lbl_inc_status.setStyleSheet(f"color: {C.ACCENT_ALT}; font-size: 11px;")

    def _on_browse_dest(self):
        cur = self.edit_dest_folder.text().strip() or str(get_default_topic_export_base())
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory", cur)
        if folder:
            self.edit_dest_folder.setText(folder)

    def _on_new_topic(self):
        dlg = TopicEditorDialog(meetings_db=self.meetings_db, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list(preferred_id=dlg.result_topic.id)

    def _on_edit_topic(self):
        if not self.active_topic:
            return
        dlg = TopicEditorDialog(topic=self.active_topic, meetings_db=self.meetings_db, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list(preferred_id=dlg.result_topic.id)

    def _on_clone_topic(self):
        if not self.active_topic:
            return
        new_topic = self.topic_manager.clone_topic(self.active_topic.id)
        preferred_id = new_topic.id if new_topic else None
        self._load_topics_list(preferred_id=preferred_id)

    def _on_delete_topic(self):
        if not self.active_topic:
            return
        cur_row = self.list_topics.currentRow()
        ret = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete topic '{self.active_topic.name}'?\n\nThis will permanently remove it from your saved topics.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            self.topic_manager.delete_topic(self.active_topic.id)
            remaining = self.topic_manager.list_topics()
            next_idx = max(0, min(cur_row, len(remaining) - 1)) if remaining else None
            self._load_topics_list(preferred_index=next_idx)

    def _on_restore_defaults(self):
        ret = QMessageBox.question(
            self,
            "Restore Built-in Topics",
            "Would you like to restore standard built-in 3GPP topics (Ambient IoT, AI/ML Phase 2, 6G Architecture, Network Energy Savings, Data Framework) to your list?\n\nAny custom topics you created will be preserved.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            count = self.topic_manager.restore_default_topics(overwrite=False)
            self._load_topics_list()
            if count > 0:
                QMessageBox.information(self, "Topics Restored", f"{count} built-in topic(s) were restored to your list.")
            else:
                QMessageBox.information(self, "Topics Restored", "All built-in topics are already present in your list.")

    def _on_view_considered_inputs(self):
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Meetings", "No meetings found in the selected timeframe.")
            return

        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        tf_label = self.combo_cycle_window.currentText()
        if "Custom" in tf_label:
            tf_label = f"{self.edit_date_from.date().toString('yyyy-MM-dd')} to {self.edit_date_to.date().toString('yyyy-MM-dd')}"

        dlg = TopicInputsViewerDialog(
            meetings_db=self.meetings_db,
            topic=self.active_topic,
            selected_rounds=self.selected_cohorts,
            scope_mode=self.scope_mode,
            timeframe_label=tf_label,
            parent=None,
        )
        dlg.export_requested.connect(self._on_start_export)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        self._inputs_dlg = dlg

    def _on_start_export(self):
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Meetings", "No meetings found in the selected timeframe.")
            return

        dest_path = Path(self.edit_dest_folder.text().strip())
        parent_dir = dest_path.parent
        target_name = dest_path.name

        self.btn_export.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.prog_bar.setValue(0)
        self.lbl_status.setText("Starting export...")

        self._worker = TopicDossierExporterThread(
            meetings_db=self.meetings_db,
            topic=self.active_topic,
            selected_cohorts=self.selected_cohorts,
            export_scope_mode=self.scope_mode,
            output_base_dir=parent_dir,
            target_folder_name=target_name,
            include_discussions=self.chk_discussions.isChecked(),
            deduplicate_revisions=self.chk_dedup.isChecked(),
            generate_consolidated_md=self.chk_consolidated.isChecked(),
            allow_network_download=self.chk_download.isChecked(),
        )
        self._worker.progress_msg.connect(self.lbl_status.setText)
        self._worker.progress_step.connect(self.prog_bar.setValue)
        self._worker.session_finished.connect(self._on_export_finished)
        self._worker.start()

    def _on_export_finished(self, success: bool, target_dir: str, msg: str):
        self.btn_export.setEnabled(True)
        self.prog_bar.setVisible(False)
        if success:
            self.lbl_status.setText("Export completed successfully!")
            self.last_target_dir = target_dir
            self._check_incremental_status()
            if target_dir and os.path.isdir(to_extended_path(target_dir)):
                self.btn_open_folder.setEnabled(True)
            QMessageBox.information(self, "Export Complete", msg)
        else:
            self.lbl_status.setText("Export failed.")
            QMessageBox.critical(self, "Export Error", msg)

    def _on_open_output(self):
        dest_text = self.edit_dest_folder.text().strip()
        last_dir = str(getattr(self, "last_target_dir", "") or "").strip()
        target = dest_text if (dest_text and os.path.isdir(to_extended_path(dest_text))) else (last_dir or dest_text)
        if not target:
            QMessageBox.warning(self, "No Folder", "No destination folder is selected.")
            return
        clean_target = os.path.normpath(strip_extended_path(target))
        if os.path.isdir(to_extended_path(clean_target)):
            _open_folder(Path(clean_target))
        else:
            self.btn_open_folder.setEnabled(False)
            QMessageBox.information(
                self,
                "Folder Not Found",
                f"Output folder does not exist yet:\n{clean_target}\n\nPlease run 'Export / Update Topic Dossier' first.",
            )
