# --- File: src/modules/meetings/ui/dialogs/ai_suggestion_dialog.py ---
"""
Agenda Item (AI) Auto-Suggestion Dialog (`AiSuggestionDialog`).

Interactive assistant helping delegates discover and auto-select relevant
Agenda Items across 3GPP Working Groups (SA, RAN, CT) based on topic keywords,
known WG Agenda Items, and Work Item correlations.
"""

from typing import Any, Dict, List, Optional, Set, Tuple
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
    QFrame, QMessageBox, QWidget, QAbstractItemView, QProgressBar,
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
    QPushButton:disabled {
        color: #94A3B8;
        background-color: #F8FAFC;
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


class AiSuggestionScanWorker(QThread):
    """Background scanner resolving candidate Agenda Items across target Working Groups."""
    progress_status = pyqtSignal(str, int, int)  # status_msg, current_wg_idx, total_wgs
    scan_finished = pyqtSignal(dict)  # results dict: WG -> List[DiscoveredAi]
    scan_failed = pyqtSignal(str)

    def __init__(
        self,
        suggester: AgendaItemSuggester,
        target_wgs: List[str],
        keywords: List[str],
        wi_tokens: List[str],
        max_meetings_per_wg: int = 6,
        parent=None,
    ):
        super().__init__(parent)
        self.suggester = suggester
        self.target_wgs = list(target_wgs)
        self.keywords = list(keywords)
        self.wi_tokens = list(wi_tokens)
        self.max_meetings_per_wg = max_meetings_per_wg
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()

    def run(self):
        try:
            def on_progress(wg_name: str, cur_idx: int, total_count: int):
                if not self._is_cancelled and not self.isInterruptionRequested():
                    self.progress_status.emit(f"Scanning recent {wg_name} agendas ({cur_idx}/{total_count})...", cur_idx, total_count)

            results = self.suggester.suggest_ais_for_topic(
                target_wgs=self.target_wgs,
                keywords=self.keywords,
                wi_tokens=self.wi_tokens,
                max_meetings_per_wg=self.max_meetings_per_wg,
                progress_callback=on_progress,
            )
            if not self._is_cancelled:
                self.scan_finished.emit(results)
        except Exception as e:
            if not self._is_cancelled:
                self.scan_failed.emit(str(e))


class AiSuggestionDialog(QDialog):
    """Interactive multi-WG Agenda Item auto-discovery and suggestion dialog."""

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        target_wgs: List[str],
        initial_keywords: List[str],
        initial_wis: Optional[List[str]] = None,
        inferred_summary: Optional[str] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle("🔍 Discover & Suggest Agenda Items across WGs")
        self.resize(920, 600)

        self.meetings_db = meetings_db
        self.target_wgs = [w.strip().upper() for w in target_wgs if w.strip()]
        self.initial_keywords = list(initial_keywords or [])
        self.initial_wis = list(initial_wis or [])
        self.inferred_summary = str(inferred_summary or "").strip()
        self.suggester = AgendaItemSuggester(self.meetings_db)

        self.discovered_results: Dict[str, List[DiscoveredAi]] = {}
        self.table_widgets: Dict[str, QTableWidget] = {}
        self.checkbox_map: Dict[str, List[Tuple[QCheckBox, DiscoveredAi]]] = {}

        self.selected_patterns_per_wg: Dict[str, List[str]] = {}
        self._scan_worker: Optional[AiSuggestionScanWorker] = None

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

        # Inferred Topic Banner (if seeded from known AI)
        if self.inferred_summary:
            banner_frame = QFrame()
            banner_frame.setStyleSheet(
                "background-color: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 6px; padding: 6px 10px;"
            )
            b_layout = QHBoxLayout(banner_frame)
            b_layout.setContentsMargins(6, 4, 6, 4)
            lbl_banner = QLabel(f"💡 <b>Inferred from known AI:</b> {self.inferred_summary}")
            lbl_banner.setStyleSheet("color: #065F46; font-size: 11px;")
            lbl_banner.setWordWrap(True)
            b_layout.addWidget(lbl_banner)
            layout.addWidget(banner_frame)

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

        self.btn_rescan = QPushButton("🔄 Re-Scan")
        self.btn_rescan.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_rescan.clicked.connect(self._trigger_scan)
        p_layout.addWidget(self.btn_rescan)

        layout.addWidget(param_frame)

        # Progress bar strip (asynchronous scanning)
        self.prog_frame = QFrame()
        self.prog_frame.setStyleSheet("background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 6px;")
        prog_layout = QHBoxLayout(self.prog_frame)
        prog_layout.setContentsMargins(10, 6, 10, 6)
        prog_layout.setSpacing(10)
        self.lbl_scan_status = QLabel("Ready to scan.")
        self.lbl_scan_status.setStyleSheet("color: #1E40AF; font-size: 11px; font-weight: bold;")
        prog_layout.addWidget(self.lbl_scan_status, 1)
        self.prog_bar = QProgressBar()
        self.prog_bar.setFixedHeight(14)
        self.prog_bar.setRange(0, 100)
        self.prog_bar.setValue(0)
        self.prog_bar.setTextVisible(False)
        self.prog_bar.setStyleSheet("""
            QProgressBar { border: 1px solid #93C5FD; border-radius: 3px; background-color: #DBEAFE; }
            QProgressBar::chunk { background-color: #3B82F6; border-radius: 2px; }
        """)
        prog_layout.addWidget(self.prog_bar, 1)
        layout.addWidget(self.prog_frame)
        self.prog_frame.setVisible(False)

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

        self.btn_apply = QPushButton("Apply Selected AIs")
        self.btn_apply.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_apply.clicked.connect(self._on_apply)
        btn_layout.addWidget(self.btn_apply)

        layout.addLayout(btn_layout)

    def _trigger_scan(self):
        kws = [k.strip().lower() for k in self.edit_keywords.text().split(",") if k.strip()]
        wis = [w.strip().lower() for w in self.edit_wis.text().split(",") if w.strip()]

        if not kws and not wis:
            QMessageBox.information(self, "No Search Terms", "Please provide at least one keyword or Work Item token.")
            return

        if self._scan_worker and self._scan_worker.isRunning():
            self._scan_worker.cancel()
            self._scan_worker.wait(500)

        self.tab_widget.clear()
        self.table_widgets.clear()
        self.checkbox_map.clear()
        self.btn_apply.setEnabled(False)
        self.btn_rescan.setEnabled(False)

        self.prog_frame.setVisible(True)
        self.prog_bar.setValue(0)
        wgs_str = ", ".join(self.target_wgs)
        self.lbl_scan_status.setText(f"Scanning agendas across {wgs_str}...")

        self._scan_worker = AiSuggestionScanWorker(
            suggester=self.suggester,
            target_wgs=self.target_wgs,
            keywords=kws,
            wi_tokens=wis,
            max_meetings_per_wg=6,
            parent=self,
        )
        self._scan_worker.progress_status.connect(self._on_scan_progress)
        self._scan_worker.scan_finished.connect(self._on_scan_finished)
        self._scan_worker.scan_failed.connect(self._on_scan_failed)
        self._scan_worker.start()

    def _on_scan_progress(self, msg: str, cur_idx: int, total_count: int):
        self.lbl_scan_status.setText(msg)
        pct = int((cur_idx / max(1, total_count)) * 100)
        self.prog_bar.setValue(pct)

    def _on_scan_finished(self, results: Dict[str, List[DiscoveredAi]]):
        self.discovered_results = results
        self.prog_frame.setVisible(False)
        self.btn_rescan.setEnabled(True)
        self.btn_apply.setEnabled(True)

        for wg in self.target_wgs:
            items = self.discovered_results.get(wg, [])
            tab_widget = self._create_wg_tab(wg, items)
            high_count = sum(1 for it in items if it.confidence == "High")
            tab_title = f"{wg} ({len(items)})"
            if high_count > 0:
                tab_title = f"{wg} ({high_count}★)"
            self.tab_widget.addTab(tab_widget, tab_title)

        self._update_preview()

    def _on_scan_failed(self, error_msg: str):
        self.prog_frame.setVisible(False)
        self.btn_rescan.setEnabled(True)
        self.btn_apply.setEnabled(True)
        QMessageBox.warning(self, "Scan Warning", f"AI scan completed with notice:\n{error_msg}")

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
            if item.confidence == "Low":
                badge_text = "⚪ Low"
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

    def get_active_keywords(self) -> List[str]:
        return [k.strip() for k in self.edit_keywords.text().split(",") if k.strip()]

    def get_active_wis(self) -> List[str]:
        return [w.strip() for w in self.edit_wis.text().split(",") if w.strip()]

    def closeEvent(self, event):
        if self._scan_worker and self._scan_worker.isRunning():
            self._scan_worker.cancel()
            self._scan_worker.wait(500)
        super().closeEvent(event)
