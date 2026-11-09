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

from PyQt5.QtCore import Qt, QSize
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QGroupBox,
    QRadioButton, QCheckBox, QPushButton, QLineEdit, QFileDialog,
    QListWidget, QListWidgetItem, QStackedWidget, QFrame, QScrollArea,
    QWidget, QButtonGroup, QMessageBox, QProgressBar, QComboBox, QTextEdit,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.paths import get_project_root, to_extended_path
from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    MeetingCohortResolver,
    format_cohort_label,
)
from modules.meetings.core.exporters.topic_dossier_exporter import (
    TopicDossierExporterThread,
    sanitize_folder_name,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import (
    MetaTopicDefinition,
    MetaTopicManager,
    MetaTopicRules,
)

ALL_COMMON_WGS = ["SA2", "RAN2", "RAN3", "CT1", "RAN1", "SA3", "CT3", "CT4", "SA1", "SA4", "SA5", "RAN4"]

STYLE_CARD = """
    QFrame#CardFrame {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
    }
"""

STYLE_PRIMARY_BTN = """
    QPushButton {
        background-color: #10B981;
        color: #FFFFFF;
        font-weight: bold;
        border: 1px solid #059669;
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 12px;
    }
    QPushButton:hover {
        background-color: #059669;
    }
    QPushButton:disabled {
        background-color: #94A3B8;
        border-color: #94A3B8;
    }
"""

STYLE_SECONDARY_BTN = """
    QPushButton {
        background-color: #F8FAFC;
        color: #1E293B;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
    }
"""

STYLE_MODE_ACTIVE = """
    QPushButton {
        background-color: #2563EB;
        color: #FFFFFF;
        font-weight: bold;
        border: 1px solid #1D4ED8;
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 11px;
    }
"""

STYLE_MODE_INACTIVE = """
    QPushButton {
        background-color: #FFFFFF;
        color: #475569;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 6px 12px;
        font-size: 11px;
    }
    QPushButton:hover {
        background-color: #F8FAFC;
        color: #0F172A;
    }
"""


class TopicEditorDialog(QDialog):
    """Sub-dialog to create or edit a MetaTopicDefinition."""

    def __init__(self, topic: Optional[MetaTopicDefinition] = None, parent=None):
        super().__init__(parent)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle("Edit Meta-Topic Definition" if topic else "New Meta-Topic")
        self.resize(580, 520)
        self.topic = topic
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

        # Target WGs
        wg_box = QGroupBox("Target Working Groups")
        wg_layout = QHBoxLayout(wg_box)
        self.wg_checks = {}
        target_set = set(self.topic.target_wgs) if self.topic else {"SA2", "RAN2", "CT1"}
        for wg in ALL_COMMON_WGS:
            cb = QCheckBox(wg)
            cb.setChecked(wg in target_set)
            self.wg_checks[wg] = cb
            wg_layout.addWidget(cb)
        layout.addWidget(wg_box)

        # Rules Group
        rules_box = QGroupBox("Matching Rules (Auto-Discovery)")
        r_layout = QVBoxLayout(rules_box)

        r_layout.addWidget(QLabel("Work Item / Study Item Tokens (comma-separated):"))
        cur_wi = ", ".join(self.topic.rules.wi_tokens) if self.topic else ""
        self.edit_wi = QLineEdit(cur_wi)
        self.edit_wi.setPlaceholderText("e.g. ambientiot, ambientiot_ph2, fs_ambientiot")
        r_layout.addWidget(self.edit_wi)

        r_layout.addWidget(QLabel("Agenda Item Patterns per WG (format: WG: patterns, comma-separated):"))
        ai_lines = []
        if self.topic and self.topic.rules.ai_patterns:
            for wg, pats in self.topic.rules.ai_patterns.items():
                ai_lines.append(f"{wg}: {', '.join(pats)}")
        self.edit_ai = QLineEdit("; ".join(ai_lines))
        self.edit_ai.setPlaceholderText("e.g. SA2: 5.14*, 5.12*; RAN2: 8.7*; CT1: 6.18*")
        r_layout.addWidget(self.edit_ai)

        r_layout.addWidget(QLabel("Keywords in Title / Abstract (comma-separated):"))
        cur_kw = ", ".join(self.topic.rules.keywords) if self.topic else ""
        self.edit_kw = QLineEdit(cur_kw)
        self.edit_kw.setPlaceholderText("e.g. ambient iot, zero energy, passive iot")
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
            target_wgs=wgs,
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
        lbl_topics.setStyleSheet("font-size: 13px; color: #1E293B;")
        left_layout.addWidget(lbl_topics)

        self.list_topics = QListWidget()
        self.list_topics.setStyleSheet("""
            QListWidget {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 8px 10px;
                border-radius: 4px;
                margin-bottom: 2px;
            }
            QListWidget::item:selected {
                background-color: #EFF6FF;
                color: #1D4ED8;
                font-weight: bold;
                border: 1px solid #BFDBFE;
            }
        """)
        self.list_topics.currentRowChanged.connect(self._on_topic_selected)
        left_layout.addWidget(self.list_topics)

        # Toolbar under topics list
        t_btn_layout = QGridLayout()
        t_btn_layout.setSpacing(4)
        btn_new = QPushButton("➕ New")
        btn_new.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_new.clicked.connect(self._on_new_topic)
        t_btn_layout.addWidget(btn_new, 0, 0)

        btn_edit = QPushButton("✏️ Edit")
        btn_edit.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_edit.clicked.connect(self._on_edit_topic)
        t_btn_layout.addWidget(btn_edit, 0, 1)

        btn_clone = QPushButton("📋 Clone")
        btn_clone.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_clone.clicked.connect(self._on_clone_topic)
        t_btn_layout.addWidget(btn_clone, 1, 0)

        btn_del = QPushButton("🗑️ Delete")
        btn_del.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_del.clicked.connect(self._on_delete_topic)
        t_btn_layout.addWidget(btn_del, 1, 1)
        left_layout.addLayout(t_btn_layout)

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
        self.lbl_topic_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #0F172A;")
        right_layout.addWidget(self.lbl_topic_title)

        self.lbl_topic_desc = QLabel("Description...")
        self.lbl_topic_desc.setStyleSheet("color: #64748B; font-size: 11px;")
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

        # 3. Meeting Cohorts / Cycles Picker
        cycle_box = QGroupBox("Meeting Cycles Scope")
        cycle_layout = QGridLayout(cycle_box)

        cycle_layout.addWidget(QLabel("<b>Cohort Window:</b>"), 0, 0)
        self.combo_cycle_window = QComboBox()
        self.combo_cycle_window.addItems([
            "Past 4 Quarters (1 Year)",
            "Past 8 Quarters (2 Years)",
            "Release 19 Window (2024–2026)",
            "All Available Cohorts",
        ])
        self.combo_cycle_window.currentIndexChanged.connect(self._update_cohorts_preview)
        cycle_layout.addWidget(self.combo_cycle_window, 0, 1)

        self.lbl_cohorts_summary = QLabel("Covering: SA2, RAN2, CT1 across 4 meeting cycles")
        self.lbl_cohorts_summary.setStyleSheet("color: #475569; font-size: 11px;")
        cycle_layout.addWidget(self.lbl_cohorts_summary, 1, 0, 1, 2)
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
        right_layout.addWidget(opt_box)

        # 6. Progress & Actions
        self.prog_bar = QProgressBar()
        self.prog_bar.setVisible(False)
        self.prog_bar.setFixedHeight(12)
        right_layout.addWidget(self.prog_bar)

        self.lbl_status = QLabel("Ready")
        self.lbl_status.setStyleSheet("color: #64748B; font-size: 11px;")
        right_layout.addWidget(self.lbl_status)

        action_layout = QHBoxLayout()
        action_layout.addStretch()

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

    def _load_topics_list(self):
        self.list_topics.clear()
        topics = self.topic_manager.list_topics()
        for t in topics:
            item = QListWidgetItem(f"{t.name}")
            item.setToolTip(f"{t.description}\nWGs: {', '.join(t.target_wgs)}")
            self.list_topics.addItem(item)

        if topics:
            self.list_topics.setCurrentRow(0)

    def _on_topic_selected(self, row: int):
        topics = self.topic_manager.list_topics()
        if 0 <= row < len(topics):
            self.active_topic = topics[row]
            self.lbl_topic_title.setText(f"🌐 {self.active_topic.name}")
            wgs_str = ", ".join(self.active_topic.target_wgs)
            self.lbl_topic_desc.setText(f"<b>WGs:</b> {wgs_str} | {self.active_topic.description}")

            # Update default destination
            root = get_project_root()
            default_name = f"Topic_{sanitize_folder_name(self.active_topic.name)}"
            self.edit_dest_folder.setText(str(root / "Export" / default_name))

            self._update_cohorts_preview()
            self._check_incremental_status()

    def _set_scope_mode(self, mode: str):
        self.scope_mode = mode
        self.btn_mode_pre.setStyleSheet(STYLE_MODE_ACTIVE if mode == "pre_meeting" else STYLE_MODE_INACTIVE)
        self.btn_mode_post.setStyleSheet(STYLE_MODE_ACTIVE if mode == "post_meeting" else STYLE_MODE_INACTIVE)
        self.btn_mode_full.setStyleSheet(STYLE_MODE_ACTIVE if mode == "full_dossier" else STYLE_MODE_INACTIVE)

    def _update_cohorts_preview(self):
        if not self.active_topic:
            return

        idx = self.combo_cycle_window.currentIndex()
        topic_copy = MetaTopicDefinition.from_dict(self.active_topic.to_dict())

        if idx == 0:  # Past 4 Quarters
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 4
        elif idx == 1:  # Past 8 Quarters
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 8
        elif idx == 2:  # Rel-19 Window
            topic_copy.cycle_mode = "cycle_range"
            topic_copy.from_cycle = "2024-Q1"
            topic_copy.to_cycle = "2026-Q4"
        else:  # All
            topic_copy.cycle_mode = "cycle_range"
            topic_copy.from_cycle = "2020-Q1"
            topic_copy.to_cycle = "2030-Q4"

        self.selected_cohorts = MeetingCohortResolver.resolve_cohorts_for_topic(self.meetings_db, topic_copy)
        if self.selected_cohorts:
            first_c = self.selected_cohorts[0].key
            last_c = self.selected_cohorts[-1].key
            total_mtgs = sum(len(c.meetings) for c in self.selected_cohorts)
            mtg_names = []
            for c in self.selected_cohorts[-2:]:
                for m in c.meetings:
                    mtg_names.append(f"{m.get('wg_name')} #{m.get('meeting_number')}")
            preview_str = ", ".join(mtg_names[:6])
            if len(mtg_names) > 6:
                preview_str += "..."
            self.lbl_cohorts_summary.setText(
                f"Resolved: <b>{len(self.selected_cohorts)} Cohort(s)</b> ({first_c} → {last_c}) with <b>{total_mtgs} meetings</b> across {', '.join(sorted(self.active_topic.target_wgs))}.<br/>"
                f"<span style='color: #64748B;'>Recent meetings: {preview_str}</span>"
            )
        else:
            self.lbl_cohorts_summary.setText("No meetings found matching this cohort window and target WGs.")

    def _check_incremental_status(self):
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            self.lbl_inc_status.setText("⚠️ Please specify a destination folder.")
            return

        p = Path(folder_str)
        manifest_p = p / ".export_manifest.json"
        if manifest_p.is_file():
            self.lbl_inc_status.setText(
                "🔄 <b>Existing Corpus Detected:</b> Incremental update mode active (skips already exported TDocs, writes delta feed)."
            )
            self.lbl_inc_status.setStyleSheet("color: #047857; font-size: 11px;")
        else:
            self.lbl_inc_status.setText("🆕 <b>New Directory:</b> Initial complete export.")
            self.lbl_inc_status.setStyleSheet("color: #2563EB; font-size: 11px;")

    def _on_browse_dest(self):
        cur = self.edit_dest_folder.text().strip() or str(get_project_root() / "Export")
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory", cur)
        if folder:
            self.edit_dest_folder.setText(folder)

    def _on_new_topic(self):
        dlg = TopicEditorDialog(parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list()

    def _on_edit_topic(self):
        if not self.active_topic:
            return
        dlg = TopicEditorDialog(topic=self.active_topic, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list()

    def _on_clone_topic(self):
        if not self.active_topic:
            return
        self.topic_manager.clone_topic(self.active_topic.id)
        self._load_topics_list()

    def _on_delete_topic(self):
        if not self.active_topic:
            return
        ret = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete topic '{self.active_topic.name}'?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            self.topic_manager.delete_topic(self.active_topic.id)
            self._load_topics_list()

    def _on_start_export(self):
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Cohorts", "No meetings found in the selected cohort window.")
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
            self.btn_open_folder.setEnabled(True)
            self.last_target_dir = target_dir
            QMessageBox.information(self, "Export Complete", msg)
        else:
            self.lbl_status.setText("Export failed.")
            QMessageBox.critical(self, "Export Error", msg)

    def _on_open_output(self):
        target = getattr(self, "last_target_dir", self.edit_dest_folder.text().strip())
        if target and os.path.isdir(target):
            os.startfile(target)
