# --- File: src/modules/meetings/ui/dialogs/ai_suggestion_dialog.py ---
"""
Agenda Item (AI) Auto-Suggestion Dialog (`AiSuggestionDialog`).

Interactive assistant helping delegates discover and auto-select relevant
Agenda Items across 3GPP Working Groups (SA, RAN, CT) based on topic keywords
and Work Item correlations.
"""

from typing import Dict, List, Optional, Set, Tuple
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
    QFrame, QMessageBox, QWidget, QAbstractItemView,
)

from modules.meetings.core.ai_suggester import AgendaItemSuggester, DiscoveredAi
from modules.meetings.core.meetings_db import MeetingsDatabase

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

STYLE_TABLE = """
    QTableWidget {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        gridline-color: #F1F5F9;
    }
    QHeaderView::section {
        background-color: #F8FAFC;
        color: #475569;
        font-weight: bold;
        font-size: 11px;
        border: none;
        border-bottom: 1px solid #E2E8F0;
        padding: 5px;
    }
    QTableWidget::item {
        padding: 4px;
        font-size: 11px;
    }
    QTableWidget::item:selected {
        background-color: #EFF6FF;
        color: #1D4ED8;
    }
"""


class AiSuggestionDialog(QDialog):
    """Interactive multi-WG Agenda Item auto-discovery and suggestion dialog."""

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        target_wgs: List[str],
        initial_keywords: List[str],
        initial_wis: Optional[List[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle("🔍 Discover & Suggest Agenda Items across WGs")
        self.resize(880, 560)

        self.meetings_db = meetings_db
        self.target_wgs = [w.strip().upper() for w in target_wgs if w.strip()]
        self.initial_keywords = list(initial_keywords or [])
        self.initial_wis = list(initial_wis or [])
        self.suggester = AgendaItemSuggester(self.meetings_db)

        self.discovered_results: Dict[str, List[DiscoveredAi]] = {}
        self.table_widgets: Dict[str, QTableWidget] = {}
        self.checkbox_map: Dict[str, List[Tuple[QCheckBox, DiscoveredAi]]] = {}

        self.selected_patterns_per_wg: Dict[str, List[str]] = {}

        self._init_ui()
        self._trigger_scan()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # Header instructions
        lbl_info = QLabel(
            "<b>Auto-Discover Agenda Items (AIs):</b> Scans recent meeting agendas across your target WGs "
            "for matching Agenda Items based on topic keywords and Work Item codes."
        )
        lbl_info.setStyleSheet("color: #334155; font-size: 12px;")
        lbl_info.setWordWrap(True)
        layout.addWidget(lbl_info)

        # Query Parameter Strip
        param_frame = QFrame()
        param_frame.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;")
        p_layout = QHBoxLayout(param_frame)
        p_layout.setContentsMargins(10, 8, 10, 8)
        p_layout.setSpacing(10)

        p_layout.addWidget(QLabel("<b>Keywords:</b>"))
        self.edit_keywords = QLineEdit(", ".join(self.initial_keywords))
        self.edit_keywords.setPlaceholderText("e.g. data framework, data transfer")
        p_layout.addWidget(self.edit_keywords, 2)

        p_layout.addWidget(QLabel("<b>WI Tokens:</b>"))
        self.edit_wis = QLineEdit(", ".join(self.initial_wis))
        self.edit_wis.setPlaceholderText("e.g. fs_6g_arc, 6g_df")
        p_layout.addWidget(self.edit_wis, 2)

        btn_rescan = QPushButton("🔄 Re-Scan")
        btn_rescan.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_rescan.clicked.connect(self._trigger_scan)
        p_layout.addWidget(btn_rescan)

        layout.addWidget(param_frame)

        # WG Tabs
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #CBD5E1; border-radius: 6px; background-color: #FFFFFF; }
            QTabBar::tab { padding: 6px 14px; font-weight: bold; font-size: 11px; }
            QTabBar::tab:selected { background-color: #FFFFFF; color: #10B981; border-bottom: 2px solid #10B981; }
        """)
        layout.addWidget(self.tab_widget, 1)

        # Action shortcuts strip
        shortcut_layout = QHBoxLayout()
        btn_check_high = QPushButton("✅ Select High Confidence")
        btn_check_high.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_check_high.clicked.connect(lambda: self._set_checks_by_confidence({"High"}))
        shortcut_layout.addWidget(btn_check_high)

        btn_check_all = QPushButton("Select All")
        btn_check_all.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_check_all.clicked.connect(lambda: self._set_all_checks(True))
        shortcut_layout.addWidget(btn_check_all)

        btn_clear = QPushButton("Clear Selection")
        btn_clear.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_clear.clicked.connect(lambda: self._set_all_checks(False))
        shortcut_layout.addWidget(btn_clear)

        shortcut_layout.addStretch()
        layout.addLayout(shortcut_layout)

        # Bottom Preview Bar
        preview_frame = QFrame()
        preview_frame.setStyleSheet("background-color: #F1F5F9; border-radius: 6px; padding: 6px 10px;")
        pv_layout = QVBoxLayout(preview_frame)
        pv_layout.setContentsMargins(6, 4, 6, 4)
        self.lbl_preview = QLabel("<b>Generated Rule:</b> <i>Select AIs above to preview...</i>")
        self.lbl_preview.setStyleSheet("color: #1E293B; font-size: 11px;")
        self.lbl_preview.setWordWrap(True)
        pv_layout.addWidget(self.lbl_preview)
        layout.addWidget(preview_frame)

        # Dialog Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_apply = QPushButton("Apply Selected AIs")
        btn_apply.setStyleSheet(STYLE_PRIMARY_BTN)
        btn_apply.clicked.connect(self._on_apply)
        btn_layout.addWidget(btn_apply)

        layout.addLayout(btn_layout)

    def _trigger_scan(self):
        kws = [k.strip().lower() for k in self.edit_keywords.text().split(",") if k.strip()]
        wis = [w.strip().lower() for w in self.edit_wis.text().split(",") if w.strip()]

        if not kws and not wis:
            QMessageBox.information(self, "No Search Terms", "Please provide at least one keyword or Work Item token.")
            return

        self.tab_widget.clear()
        self.table_widgets.clear()
        self.checkbox_map.clear()

        # Run discovery engine
        self.discovered_results = self.suggester.suggest_ais_for_topic(
            target_wgs=self.target_wgs,
            keywords=kws,
            wi_tokens=wis,
            max_meetings_per_wg=6,
        )

        for wg in self.target_wgs:
            items = self.discovered_results.get(wg, [])
            tab_widget = self._create_wg_tab(wg, items)
            tab_title = f"{wg} ({len(items)})"
            self.tab_widget.addTab(tab_widget, tab_title)

        self._update_preview()

    def _create_wg_tab(self, wg: str, items: List[DiscoveredAi]) -> QWidget:
        container = QWidget()
        v_layout = QVBoxLayout(container)
        v_layout.setContentsMargins(6, 6, 6, 6)

        if not items:
            lbl_empty = QLabel(f"<i>No matching Agenda Items found in recent {wg} meetings.</i>")
            lbl_empty.setAlignment(Qt.AlignCenter)
            lbl_empty.setStyleSheet("color: #94A3B8; font-size: 12px; margin-top: 40px;")
            v_layout.addWidget(lbl_empty)
            return container

        table = QTableWidget(len(items), 5)
        table.setStyleSheet(STYLE_TABLE)
        table.setHorizontalHeaderLabels(["Include", "Confidence", "Agenda Item", "Description", "Activity & Reason"])
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)

        self.checkbox_map[wg] = []

        for row, item in enumerate(items):
            # Column 0: Checkbox (auto-checked for High confidence)
            cb = QCheckBox()
            cb.setChecked(item.confidence == "High")
            cb.stateChanged.connect(self._update_preview)
            self.checkbox_map[wg].append((cb, item))

            cb_widget = QWidget()
            cb_layout = QHBoxLayout(cb_widget)
            cb_layout.setContentsMargins(6, 2, 6, 2)
            cb_layout.setAlignment(Qt.AlignCenter)
            cb_layout.addWidget(cb)
            table.setCellWidget(row, 0, cb_widget)

            # Column 1: Confidence badge
            badge_text = "🟢 High" if item.confidence == "High" else "🟡 Medium"
            item_conf = QTableWidgetItem(badge_text)
            item_conf.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 1, item_conf)

            # Column 2: AI Number / Suggested Pattern
            item_num = QTableWidgetItem(item.suggested_pattern)
            item_num.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 2, item_num)

            # Column 3: Description & Sample Tooltip
            item_desc = QTableWidgetItem(item.ai_description or "—")
            tip_lines = [
                f"<b>Agenda Item:</b> {item.ai_number}",
                f"<b>Description:</b> {item.ai_description}",
                f"<b>Reason:</b> {item.matching_reason}",
                f"<b>Meetings:</b> {', '.join(item.meetings[:4])}",
            ]
            if item.sample_titles:
                tip_lines.append("<b>Sample Contributions:</b>")
                for st in item.sample_titles:
                    tip_lines.append(f"• {st}")
            item_desc.setToolTip("<br/>".join(tip_lines))
            table.setItem(row, 3, item_desc)

            # Column 4: Activity summary
            act_text = f"{item.matching_tdocs}/{item.tdoc_count} TDocs ({len(item.meetings)} mtgs)"
            item_act = QTableWidgetItem(act_text)
            item_act.setTextAlignment(Qt.AlignCenter)
            item_act.setToolTip(item.matching_reason)
            table.setItem(row, 4, item_act)

        v_layout.addWidget(table)
        self.table_widgets[wg] = table
        return container

    def _set_checks_by_confidence(self, allowed_confidences: Set[str]):
        for wg, cb_list in self.checkbox_map.items():
            for cb, item in cb_list:
                cb.setChecked(item.confidence in allowed_confidences)
        self._update_preview()

    def _set_all_checks(self, checked: bool):
        for wg, cb_list in self.checkbox_map.items():
            for cb, item in cb_list:
                cb.setChecked(checked)
        self._update_preview()

    def _update_preview(self):
        rule_parts = []
        self.selected_patterns_per_wg.clear()

        for wg in self.target_wgs:
            pats = []
            for cb, item in self.checkbox_map.get(wg, []):
                if cb.isChecked():
                    pats.append(item.suggested_pattern)
            if pats:
                self.selected_patterns_per_wg[wg] = pats
                rule_parts.append(f"{wg}: {', '.join(pats)}")

        if rule_parts:
            combined = "; ".join(rule_parts)
            self.lbl_preview.setText(f"<b>Generated Rule:</b> <code>{combined}</code>")
        else:
            self.lbl_preview.setText("<b>Generated Rule:</b> <i>No AIs currently selected.</i>")

    def _on_apply(self):
        self._update_preview()
        if not self.selected_patterns_per_wg:
            QMessageBox.information(self, "No AIs Selected", "Please select at least one suggested Agenda Item.")
            return
        self.accept()

    def get_formatted_rule_string(self) -> str:
        """Returns canonical WG: pat1, pat2; WG2: pat... rule string."""
        chunks = []
        for wg, pats in self.selected_patterns_per_wg.items():
            if pats:
                chunks.append(f"{wg}: {', '.join(pats)}")
        return "; ".join(chunks)
