"""TopicEditorDialog for creating and editing Cross-WG MetaTopicDefinition rules."""

import re
from typing import Dict, List, Optional

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QDateEdit,
    QDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from core.utils.paths import get_project_root
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import MetaTopicDefinition, MetaTopicRules
from modules.meetings.ui.dialogs.topic_dossier.styles import (
    ALL_COMMON_WGS,
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
)
from modules.meetings.ui.topic_pills import (
    AgendaItemPillsWidget,
    TagCloudWidget,
    WgTogglePill,
)


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
        self.resize(720, 750)
        self.topic = topic
        self.meetings_db = meetings_db
        self._init_ui()

    def _init_ui(self):
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(12, 12, 12, 12)
        outer_layout.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        body = QWidget()
        body.setObjectName("TopicEditorScrollBody")
        body.setStyleSheet("QWidget#TopicEditorScrollBody { background: transparent; }")
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 0, 4, 0)
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

        grid.addWidget(QLabel("<b>Default Start Date:</b>"), 2, 0)
        start_row = QHBoxLayout()
        start_row.setContentsMargins(0, 0, 0, 0)
        start_row.setSpacing(8)
        self.chk_open_start = QCheckBox("Anchor start date (Since Start Date → Present)")
        is_since = bool(self.topic and self.topic.cycle_mode == "since_date")
        self.chk_open_start.setChecked(is_since)
        self.edit_start_date = QDateEdit()
        self.edit_start_date.setCalendarPopup(True)
        self.edit_start_date.setDisplayFormat("yyyy-MM-dd")
        q_df = QDate.fromString(str(self.topic.date_from or "")[:10], "yyyy-MM-dd") if self.topic else QDate()
        self.edit_start_date.setDate(q_df if q_df.isValid() else QDate.currentDate().addYears(-1))
        self.edit_start_date.setEnabled(is_since)
        self.chk_open_start.toggled.connect(self.edit_start_date.setEnabled)
        start_row.addWidget(self.chk_open_start)
        start_row.addWidget(self.edit_start_date)
        start_row.addStretch()
        grid.addLayout(start_row, 2, 1)
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
        self.edit_wi = TagCloudWidget(placeholder="Add WI/SI token (Enter or ,)...")
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
        self.edit_kw = TagCloudWidget(placeholder="Add keyword (Enter or ,)...")
        if self.topic and self.topic.rules.keywords:
            self.edit_kw.setText(", ".join(self.topic.rules.keywords))
        r_layout.addWidget(self.edit_kw)

        layout.addWidget(rules_box)
        layout.addStretch()

        scroll.setWidget(body)
        outer_layout.addWidget(scroll, 1)

        # Fixed Dialog Footer Buttons
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
        outer_layout.addLayout(btn_layout)

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

        if self.chk_open_start.isChecked():
            cycle_mode = "since_date"
            date_from = self.edit_start_date.date().toString("yyyy-MM-dd")
            date_to = ""
        else:
            prev_mode = self.topic.cycle_mode if self.topic else "past_n_cycles"
            cycle_mode = "past_n_cycles" if prev_mode == "since_date" else prev_mode
            date_from = self.topic.date_from if self.topic and prev_mode != "since_date" else ""
            date_to = self.topic.date_to if self.topic and prev_mode != "since_date" else ""

        topic_id = self.topic.id if self.topic else re.sub(r'[^a-zA-Z0-9_-]+', '_', name.lower()).strip('_')
        self.result_topic = MetaTopicDefinition(
            id=topic_id,
            name=name,
            description=self.edit_desc.text().strip(),
            color=self.topic.color if self.topic else C.EMERALD_ACCENT,
            target_wgs=wgs,
            cycle_mode=cycle_mode,
            past_n_cycles=self.topic.past_n_cycles if self.topic else 4,
            from_cycle=self.topic.from_cycle if self.topic else "2025-Q1",
            to_cycle=self.topic.to_cycle if self.topic else "2026-Q1",
            date_from=date_from,
            date_to=date_to,
            rules=rules,
        )
        self.accept()
