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
    QDateEdit, QMenu, QInputDialog,
)

from core.ai.cli_detector import CLI_AGENT_SPECS, detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import (
    add_model_selector_submenu,
    launch_cli_in_terminal,
    open_mcp_agent_setup_dialog,
    write_cli_context_file,
)
from core.ui.cli_context_dialog import CliContextEditorDialog
from core.ui.flow_layout import FlowLayout
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.paths import get_project_root, to_extended_path, strip_extended_path
from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    MeetingCohortResolver,
    format_cohort_label,
)
from modules.meetings.core.exporters.export_manifest import load_manifest
from modules.meetings.core.exporters.topic_digest_builder import STAGE2_SYNTHESIS_FILENAME
from modules.meetings.core.exporters.topic_dossier_exporter import (
    TopicDossierExporterThread,
    check_and_refresh_dossier,
    discover_stage2_packages,
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

        # 2. Meeting Scope & Timeframe Picker
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
        self.lbl_inc_status.setWordWrap(True)
        self.lbl_inc_status.setStyleSheet("font-size: 11px; font-weight: bold;")
        dest_layout.addWidget(self.lbl_inc_status, 1, 0, 1, 3)
        right_layout.addWidget(dest_box)

        # 4. Conversion Options
        opt_box = QGroupBox("Dossier Content && Options")
        opt_layout = QHBoxLayout(opt_box)
        self.chk_discussions = QCheckBox("Include Discussion Papers && WFs")
        self.chk_discussions.setChecked(True)
        self.chk_discussions.setToolTip(
            "Include non-agreed Discussion papers and Way Forwards alongside binding decisions and LSs "
            "(recommended for E2E company position tracking; uncheck for a lightweight Decisions/CRs/LSs baseline)."
        )
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

        self.btn_check_status = QPushButton("🔍 Check Status && Refresh Index")
        self.btn_check_status.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_check_status.setEnabled(False)
        self.btn_check_status.setToolTip(
            "Fast offline scan of the destination folder to link _company_positions_analysis.md and "
            "company_positions_e2e_synthesis.md, flag stale deltas, and refresh index.md without re-exporting documents."
        )
        self.btn_check_status.clicked.connect(lambda: self._on_check_status(silent=False))
        action_layout.addWidget(self.btn_check_status)

        self.btn_open_folder = QPushButton("📂 Open Output Folder")
        self.btn_open_folder.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self._on_open_output)
        action_layout.addWidget(self.btn_open_folder)

        self.btn_launch_cli = QPushButton("💬 Launch CLI Agent")
        self.btn_launch_cli.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_launch_cli.setEnabled(False)
        self.btn_launch_cli.setToolTip("Open Codex CLI, Claude CLI, or Gemini CLI directly inside the exported Cross-WG Topic Dossier.")
        self.cli_menu = QMenu(self)
        self.cli_menu.aboutToShow.connect(self._populate_cli_menu)
        self.btn_launch_cli.setMenu(self.cli_menu)
        action_layout.addWidget(self.btn_launch_cli)

        self.btn_export = QPushButton("🚀 Export / Update Topic Dossier")
        self.btn_export.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_export.clicked.connect(self._on_start_export)
        action_layout.addWidget(self.btn_export)

        right_layout.addLayout(action_layout)
        root_layout.addWidget(right_frame)

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
            self.btn_launch_cli.setEnabled(False)
            self.btn_check_status.setEnabled(False)
            return

        folder_exists = os.path.isdir(to_extended_path(folder_str))
        self.btn_open_folder.setEnabled(folder_exists)
        self.btn_launch_cli.setEnabled(folder_exists)
        self.btn_open_folder.setToolTip(
            f"Open in File Explorer: {folder_str}"
            if folder_exists
            else "Destination folder does not exist yet. Run an export first."
        )

        p = Path(folder_str)
        manifest_p = p / ".export_manifest.json"
        has_manifest = manifest_p.is_file()
        self.btn_check_status.setEnabled(has_manifest)
        if not has_manifest:
            self.lbl_inc_status.setText("🆕 <b>New Directory:</b> Initial complete export.")
            self.lbl_inc_status.setStyleSheet(f"color: {C.ACCENT_ALT}; font-size: 11px;")
            return

        manifest = load_manifest(p)
        if not manifest or not manifest.exported_meetings:
            self.lbl_inc_status.setText(
                "🔄 <b>Existing Corpus Detected:</b> Incremental update mode active (skips already exported TDocs, writes delta feed)."
            )
            self.lbl_inc_status.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px;")
            return

        analyzed_cnt = 0
        stale_cnt = 0
        pending_cnt = 0
        no_act_cnt = 0
        dec_cnt = 0
        doc_cnt = 0
        stub_cnt = 0
        max_input_mtime = 0.0
        excluded_wgs: Set[str] = set()

        for m_key, m_state in manifest.exported_meetings.items():
            if not isinstance(m_state, dict):
                continue
            if m_state.get("wg"):
                excluded_wgs.add(str(m_state["wg"]).strip())
            d_c = int(m_state.get("decision_count", 0) or 0)
            t_c = int(m_state.get("tdoc_count", 0) or 0)
            s_c = int(m_state.get("stub_count", 0) or 0)
            dec_cnt += d_c
            doc_cnt += t_c
            stub_cnt += s_c

            lc = str(m_state.get("lifecycle", "") or "").strip().lower()
            if lc == "no_activity" or (d_c == 0 and t_c == 0):
                no_act_cnt += 1
                continue

            cur_snap = m_state.get("current_snapshot") or {}
            pkt_mtime = float(cur_snap.get("packet_mtime", 0.0) or 0.0)
            max_input_mtime = max(max_input_mtime, pkt_mtime)

            a_file = Path(to_extended_path(p / m_key / "_company_positions_analysis.md"))
            if not a_file.is_file():
                pending_cnt += 1
                continue

            st = m_state.get("analysis_status")
            ana_snap = m_state.get("analyzed_snapshot") or {}
            a_mtime = float(os.path.getmtime(a_file))
            max_input_mtime = max(max_input_mtime, a_mtime)

            if st == "stale" and a_mtime > pkt_mtime > 0.0:
                analyzed_cnt += 1
            elif st in ("analyzed", "stale"):
                if st == "analyzed" and ana_snap and cur_snap and ana_snap.get("fingerprint") != cur_snap.get("fingerprint") and a_mtime <= pkt_mtime:
                    stale_cnt += 1
                elif st == "stale":
                    stale_cnt += 1
                else:
                    analyzed_cnt += 1
            else:
                # Legacy manifest without snapshot fields
                if pkt_mtime > 0.0 and a_mtime < pkt_mtime - 5.0:
                    stale_cnt += 1
                else:
                    analyzed_cnt += 1

        packages = discover_stage2_packages(p, excluded_dirs=excluded_wgs)
        completed_pkgs = [pkg for pkg in packages if pkg.get("status") == "completed"]
        in_prog_pkgs = [pkg for pkg in packages if pkg.get("status") == "in_progress"]
        syn_state_dict = manifest.synthesis_state or {}
        syn_state = str(syn_state_dict.get("status") or "pending").lower()

        if not completed_pkgs:
            if in_prog_pkgs:
                syn_badge = f"🏗️ Stage-2 In Progress [{in_prog_pkgs[-1]['folder_name']}]"
            else:
                syn_badge = "⏳ Stage-2 Pending"
        else:
            latest_pkg = completed_pkgs[-1]
            dated_lbl = f"{latest_pkg.get('as_of') or 'dated'} ({latest_pkg.get('revision') or 'r1'})"
            folder_tag = f" [{latest_pkg['folder_name']}]" if latest_pkg.get("folder_name") else " [root]"
            runs_tag = f" · {len(completed_pkgs)} runs" if len(completed_pkgs) > 1 else ""
            latest_mtime = float(latest_pkg.get("report_mtime", 0.0) or 0.0)
            is_stale = (
                syn_state == "stale"
                or stale_cnt > 0
                or (latest_mtime > 0.0 and max_input_mtime > latest_mtime + 0.5)
            )
            if is_stale:
                chg_cnt = len(syn_state_dict.get("changed_meetings") or [])
                if chg_cnt:
                    syn_badge = f"🔄 Stage-2 Stale: {chg_cnt} updated since {dated_lbl}{folder_tag}{runs_tag}"
                else:
                    syn_badge = f"🔄 Stage-2 Stale (since {dated_lbl}{folder_tag}){runs_tag}"
            else:
                syn_badge = f"✅ Stage-2: {dated_lbl}{folder_tag}{runs_tag}"

        total_mtgs = len(manifest.exported_meetings)
        stage1_parts = [f"✅ {analyzed_cnt} Analyzed"]
        if stale_cnt:
            stage1_parts.append(f"🔄 {stale_cnt} Stale")
        stage1_parts.append(f"⏳ {pending_cnt} Pending")
        if no_act_cnt:
            stage1_parts.append(f"⚪ {no_act_cnt} No Activity")

        stub_note = f" ({stub_cnt} stubs)" if stub_cnt else ""
        self.lbl_inc_status.setText(
            f"🔄 <b>Existing Corpus Detected:</b> {total_mtgs} meetings · {dec_cnt} decisions · "
            f"{doc_cnt} docs{stub_note}<br>"
            f"<b>LLM Stage-1:</b> {' · '.join(stage1_parts)} &nbsp;|&nbsp; <b>{syn_badge}</b>"
        )
        self.lbl_inc_status.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px;")

    def _on_check_status(self, silent: bool = False) -> Dict[str, object]:
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            if not silent:
                QMessageBox.warning(self, "No Destination", "Please specify a destination folder first.")
            return {"ok": False, "error": "Empty destination"}

        res = check_and_refresh_dossier(Path(folder_str), topic=self.active_topic)
        self._check_incremental_status()

        if not res.get("ok"):
            if not silent:
                QMessageBox.warning(
                    self,
                    "Checkup Failed",
                    str(res.get("error") or "Could not inspect destination corpus."),
                )
            return res

        syn = str(res.get("synthesis_status") or "pending").lower()
        s2_state = res.get("synthesis_state") or {}
        latest_pkg = s2_state.get("latest_package") or {}
        all_pkgs = list(s2_state.get("packages") or [])
        completed_cnt = sum(1 for pk in all_pkgs if pk.get("status") == "completed")

        if latest_pkg:
            dated_s2 = f"{latest_pkg.get('as_of') or 'dated'} ({latest_pkg.get('revision') or 'r1'})"
            s2_folder = latest_pkg.get("folder_name") or "(root)"
            if syn == "stale":
                chg_cnt = len(s2_state.get("changed_meetings") or [])
                syn_short = f"🔄 stale ({chg_cnt} updated since {dated_s2})" if chg_cnt else f"🔄 stale (since {dated_s2})"
            else:
                syn_short = f"✅ {dated_s2} [{s2_folder}]"
        elif syn == "in_progress":
            syn_short = f"🏗️ in-progress [{s2_state.get('folder_name') or 'stage2'}]"
        else:
            syn_short = "⏳ pending"

        self.lbl_status.setText(
            f"Index refreshed: ✅ {res['analyzed']} analyzed · 🔄 {res['stale']} stale · "
            f"⏳ {res['pending']} pending · ⚪ {res['no_activity']} no-activity · Stage-2: {syn_short}"
        )

        if not silent:
            promo = res.get("newly_promoted", 0)
            promo_line = f"\n• Newly Linked / Promoted Analyses: +{promo}" if promo else ""
            if latest_pkg:
                val_note = (
                    f" | Validation: {'PASSED' if latest_pkg.get('validation_passed') else 'ISSUES'}"
                    if latest_pkg.get("validation_passed") is not None
                    else ""
                )
                claims_note = f" ({latest_pkg['claims']} claims)" if latest_pkg.get("claims") is not None else ""
                s2_detail = (
                    f"• Stage-2 Synthesis: {syn.upper()} — {s2_folder} "
                    f"[{latest_pkg.get('package_id', '')}, cut-off {latest_pkg.get('as_of', 'N/A')} {latest_pkg.get('revision', '')}]"
                    f"{val_note}{claims_note}"
                    + (f" ({completed_cnt} package(s) total)" if completed_cnt > 1 else "")
                )
                chg_list = list(s2_state.get("changed_meetings") or [])
                if chg_list:
                    preview_chg = ", ".join(chg_list[:5]) + ("..." if len(chg_list) > 5 else "")
                    s2_detail += f"\n  ↳ Updated since baseline ({len(chg_list)}): {preview_chg}"
            else:
                s2_detail = f"• Stage-2 E2E Synthesis ({STAGE2_SYNTHESIS_FILENAME} / stage2_*): {syn.upper()}"

            QMessageBox.information(
                self,
                "Corpus Status & Index Refreshed",
                f"Checked '{res.get('topic_name', '')}' ({res['total_meetings']} meetings):\n"
                f"• Decisions: {res['total_decisions']} | Active Docs: {res['total_active_docs']} (Stubs: {res['total_stubs']})\n"
                f"• Stage-1 Analyses: ✅ {res['analyzed']} Analyzed | 🔄 {res['stale']} Stale (Delta Ready) | "
                f"⏳ {res['pending']} Pending | ⚪ {res['no_activity']} No Activity"
                f"{promo_line}\n"
                f"{s2_detail}\n\n"
                f"Updated index.md, company_positions_e2e_prompt.md, per-meeting briefs, and .export_manifest.json.",
            )
        return res

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
            scope_mode="full_dossier",
            timeframe_label=tf_label,
            parent=None,
        )
        dlg.export_requested.connect(self._on_start_export)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        self._inputs_dlg = dlg

    def _on_start_export(self):
        if self._worker and self._worker.isRunning():
            self.raise_()
            self.activateWindow()
            return
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Meetings", "No meetings found in the selected timeframe.")
            return

        if getattr(self, "_inputs_dlg", None) is not None:
            try:
                if self._inputs_dlg.isVisible():
                    self._inputs_dlg.close()
            except Exception:
                pass

        dest_path = Path(self.edit_dest_folder.text().strip())
        parent_dir = dest_path.parent
        target_name = dest_path.name

        self.btn_export.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.prog_bar.setValue(0)
        self.lbl_status.setText("Starting export...")
        self.raise_()
        self.activateWindow()

        self._worker = TopicDossierExporterThread(
            meetings_db=self.meetings_db,
            topic=self.active_topic,
            selected_cohorts=self.selected_cohorts,
            export_scope_mode="full_dossier",
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

    def closeEvent(self, event):
        if getattr(self, "_inputs_dlg", None) is not None:
            try:
                self._inputs_dlg.close()
            except Exception:
                pass
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait(5000)
        super().closeEvent(event)

    def _on_export_finished(self, success: bool, target_dir: str, msg: str):
        self.btn_export.setEnabled(True)
        self.prog_bar.setVisible(False)
        if success:
            self.lbl_status.setText("Export completed successfully!")
            self.last_target_dir = target_dir
            self._check_incremental_status()
            if target_dir and os.path.isdir(to_extended_path(target_dir)):
                self.btn_open_folder.setEnabled(True)
                self.btn_launch_cli.setEnabled(True)
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
            self.btn_launch_cli.setEnabled(False)
            QMessageBox.information(
                self,
                "Folder Not Found",
                f"Output folder does not exist yet:\n{clean_target}\n\nPlease run 'Export / Update Topic Dossier' first.",
            )

    def _collect_topic_manifest_summary(self, folder: Path) -> Dict[str, object]:
        """Inspect `.export_manifest.json` in `folder` for Stage-1 and Stage-2 status."""
        summary: Dict[str, object] = {
            "analyzed_meetings": [],
            "stale_meetings": [],
            "pending_meetings": [],
            "synthesis_status": "pending",
        }
        manifest = load_manifest(folder)
        if not manifest or not manifest.exported_meetings:
            return summary

        for m_key, m_state in manifest.exported_meetings.items():
            if not isinstance(m_state, dict):
                continue
            lc = str(m_state.get("lifecycle", "") or "").strip().lower()
            d_c = int(m_state.get("decision_count", 0) or 0)
            t_c = int(m_state.get("tdoc_count", 0) or 0)
            if lc == "no_activity" or (d_c == 0 and t_c == 0):
                continue

            a_file = Path(to_extended_path(folder / m_key / "_company_positions_analysis.md"))
            st = str(m_state.get("analysis_status") or "").strip().lower()
            if not a_file.is_file():
                summary["pending_meetings"].append(m_key)
            elif st == "stale":
                summary["stale_meetings"].append(m_key)
            else:
                summary["analyzed_meetings"].append(m_key)

        syn_state = manifest.synthesis_state or {}
        summary["synthesis_status"] = str(syn_state.get("status") or "pending").lower()
        return summary

    def _populate_cli_menu(self):
        self.cli_menu.clear()
        self.cli_menu.setToolTipsVisible(True)

        folder_str = self.edit_dest_folder.text().strip()
        summary = self._collect_topic_manifest_summary(Path(folder_str)) if folder_str else {}
        pending_cnt = len(summary.get("pending_meetings", []))
        stale_cnt = len(summary.get("stale_meetings", []))
        stage1_hint = f" ({pending_cnt} pending, {stale_cnt} stale)" if (pending_cnt or stale_cnt) else ""

        statuses = detect_cli_agents()
        installed = [st for st in statuses.values() if st.installed]
        missing = [st for st in statuses.values() if not st.installed]

        if installed:
            for st in installed:
                badge = format_active_model_badge(st.agent_id)
                sub = self.cli_menu.addMenu(f"{st.icon} {st.label}  [{badge}]")
                sub.setToolTipsVisible(True)
                sub.setToolTip(f"{st.executable_path}\nModel: {badge}")

                act_qa = sub.addAction("💬 Interactive Q&&A on Topic Corpus")
                act_qa.setToolTip(f"Launch {st.label} inside the topic dossier folder.")
                act_qa.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "topic_qa"))

                act_s1 = sub.addAction(f"🧪 Run Stage-1 Meeting Analysis{stage1_hint}")
                act_s1.setToolTip(f"Launch {st.label} focused on pending/stale per-meeting analysis files.")
                act_s1.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "topic_stage1"))

                act_s2 = sub.addAction("🌐 Run Stage-2 E2E Cross-WG Synthesis")
                act_s2.setToolTip(f"Launch {st.label} to run or update company_positions_e2e_synthesis.md.")
                act_s2.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "topic_stage2"))

                sub.addSeparator()
                act_custom = sub.addAction("✏️ Custom Question on Topic Corpus...")
                act_custom.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "custom"))

                act_ctx = sub.addAction("📝 Edit Topic Context / Notes && Launch...")
                act_ctx.setToolTip(
                    "Edit persistent topic _USER_NOTES.md / standing company guidance and select Model / Reasoning Effort before launching."
                )
                act_ctx.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "edit_context"))

                sub.addSeparator()
                add_model_selector_submenu(sub, st.agent_id)
        else:
            no_cli = self.cli_menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
            no_cli.setEnabled(False)

        if missing:
            self.cli_menu.addSeparator()
            for st in missing:
                act_miss = self.cli_menu.addAction(f"⚪ {st.label} (Not Found)")
                act_miss.setEnabled(False)
                act_miss.setToolTip(f"Not detected.\nInstall: {st.install_hint}")

        self.cli_menu.addSeparator()
        act_cfg = self.cli_menu.addAction("⚙️ Configure CLI Agents && MCP...")
        act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(self))

    def _launch_cli_for_topic(self, agent_id: str, task_mode: str):
        dest_text = self.edit_dest_folder.text().strip()
        if not dest_text or not os.path.isdir(to_extended_path(dest_text)):
            QMessageBox.warning(self, "Folder Not Found", "Please export the topic dossier before launching a CLI agent.")
            return

        target_dir = Path(strip_extended_path(dest_text))
        topic_name = self.active_topic.name if self.active_topic else target_dir.name
        custom_question = ""
        model_override: Optional[str] = None
        effort_override: Optional[str] = None

        if task_mode == "edit_context":
            topic_tasks = [
                ("topic_qa", "💬 Interactive Q&A on Topic Corpus"),
                ("topic_stage1", "🧪 Run Stage-1 Meeting Analysis"),
                ("topic_stage2", "🌐 Run Stage-2 E2E Cross-WG Synthesis"),
                ("custom", "✏️ Custom Question on Topic Corpus"),
            ]
            dlg = CliContextEditorDialog(
                agent_id=agent_id,
                scope_label=f"Topic: {topic_name}",
                target_dir=target_dir,
                task_options=topic_tasks,
                initial_task_mode="topic_qa",
                parent=self,
            )
            if dlg.exec_() != QDialog.Accepted:
                return
            res = dlg.get_result()
            if not res.launch_requested:
                return
            task_mode = res.task_mode
            custom_question = res.custom_question
            model_override = res.model
            effort_override = res.reasoning_effort

        if task_mode == "custom" and not custom_question.strip():
            spec = CLI_AGENT_SPECS.get((agent_id or "").strip().lower())
            agent_label = spec.label if spec else agent_id
            q_text, ok = QInputDialog.getMultiLineText(
                self,
                f"Ask {agent_label}",
                f"Enter your question about topic '{topic_name}':",
                "",
            )
            if not ok or not q_text.strip():
                return
            custom_question = q_text.strip()

        summary = self._collect_topic_manifest_summary(target_dir)
        write_cli_context_file(
            target_dir,
            task_mode=task_mode,
            custom_question=custom_question,
            topic_name=topic_name,
            manifest_summary=summary,
        )
        launch_cli_in_terminal(
            agent_id,
            target_dir,
            parent=self,
            model=model_override,
            reasoning_effort=effort_override,
        )
