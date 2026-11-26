"""Tab 1 (Live Activity Stream) setup, polling, filtering, and payload inspection mixin."""

from __future__ import annotations

import json
from typing import Any, Dict, List

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ai.mcp_activity import (
    clear_mcp_activity_log,
    get_active_mcp_sessions,
    get_mcp_activity_stats,
    get_recent_mcp_events,
)
from core.ui.theme import C, T
from modules.assistant.core.tools.external_agent_tools import PROFILE_TOOLS


class McpActivityStreamMixin:
    """Renders live MCP session telemetry, activity table filters, and JSON inspector."""

    def _setup_activity_tab(self, parent: QWidget):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        # Controls row (Search filter, Tool filter, Status filter)
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        self.search_filter_edit = QLineEdit()
        self.search_filter_edit.setPlaceholderText("🔍 Filter by query, tool, action, or client...")
        self.search_filter_edit.textChanged.connect(self._apply_table_filters)
        filter_row.addWidget(self.search_filter_edit, stretch=2)

        self.tool_filter_combo = QComboBox()
        self.tool_filter_combo.addItem("All Tools", "")
        for tname in PROFILE_TOOLS.get("all", []):
            self.tool_filter_combo.addItem(tname, tname)
        self.tool_filter_combo.currentIndexChanged.connect(self._apply_table_filters)
        filter_row.addWidget(self.tool_filter_combo, stretch=1)

        self.status_filter_combo = QComboBox()
        self.status_filter_combo.addItem("All Statuses", "all")
        self.status_filter_combo.addItem("🟢 Success Only", "success")
        self.status_filter_combo.addItem("🔴 Errors Only", "error")
        self.status_filter_combo.currentIndexChanged.connect(self._apply_table_filters)
        filter_row.addWidget(self.status_filter_combo, stretch=1)

        self.autoscroll_chk = QCheckBox("Auto-scroll")
        self.autoscroll_chk.setChecked(True)
        filter_row.addWidget(self.autoscroll_chk)

        layout.addLayout(filter_row)

        # Splitter: Table on top, Payload inspector on bottom
        splitter = QSplitter(Qt.Vertical)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Time", "Client", "PID", "Tool", "Action", "Summary", "Latency", "Status"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {C.BORDER_SUBTLE};
                gridline-color: {C.BG_CANVAS};
                font-size: 12px;
            }}
            QTableWidget::item:selected {{
                background: {C.INFO_SOFT};
                color: {C.INFO_NAVY};
            }}
        """)
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)
        splitter.addWidget(self.table)

        # Bottom Inspector Frame
        inspector_frame = QFrame()
        inspector_frame.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px;")
        insp_layout = QVBoxLayout(inspector_frame)
        insp_layout.setContentsMargins(8, 8, 8, 8)
        insp_layout.setSpacing(6)

        insp_header = QHBoxLayout()
        self.insp_title = QLabel("Inspection Pane: Select an event row above to inspect parameters and payload")
        self.insp_title.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {C.TEXT_BODY};")
        insp_header.addWidget(self.insp_title)
        insp_header.addStretch()

        self.copy_args_btn = QPushButton("📋 Copy Arguments")
        self.copy_args_btn.setEnabled(False)
        self.copy_args_btn.setStyleSheet("padding: 2px 8px; font-size: 11px;")
        self.copy_args_btn.clicked.connect(self._copy_selected_args)
        insp_header.addWidget(self.copy_args_btn)

        self.copy_resp_btn = QPushButton("📋 Copy Response")
        self.copy_resp_btn.setEnabled(False)
        self.copy_resp_btn.setStyleSheet("padding: 2px 8px; font-size: 11px;")
        self.copy_resp_btn.clicked.connect(self._copy_selected_response)
        insp_header.addWidget(self.copy_resp_btn)

        insp_layout.addLayout(insp_header)

        # Split text views: Arguments (Left) and Response/Error Preview (Right)
        text_splitter = QSplitter(Qt.Horizontal)

        left_pane = QWidget()
        lp_layout = QVBoxLayout(left_pane)
        lp_layout.setContentsMargins(0, 0, 0, 0)
        lp_layout.addWidget(QLabel("Request Arguments:"))
        self.args_view = QTextEdit()
        self.args_view.setReadOnly(True)
        self.args_view.setStyleSheet(f"font-family: Consolas, 'Courier New', monospace; font-size: {T.SIZE_BODY}px;")
        self.args_view.setPlaceholderText("No arguments")
        lp_layout.addWidget(self.args_view)
        text_splitter.addWidget(left_pane)

        right_pane = QWidget()
        rp_layout = QVBoxLayout(right_pane)
        rp_layout.setContentsMargins(0, 0, 0, 0)
        rp_layout.addWidget(QLabel("Response Payload / Preview:"))
        self.resp_view = QTextEdit()
        self.resp_view.setReadOnly(True)
        self.resp_view.setStyleSheet(f"font-family: Consolas, 'Courier New', monospace; font-size: {T.SIZE_BODY}px;")
        self.resp_view.setPlaceholderText("No response preview")
        rp_layout.addWidget(self.resp_view)
        text_splitter.addWidget(right_pane)

        text_splitter.setSizes([380, 560])
        insp_layout.addWidget(text_splitter)

        splitter.addWidget(inspector_frame)
        splitter.setSizes([380, 220])
        layout.addWidget(splitter)

    def _on_timer_refresh(self):
        if self.isVisible() and self.auto_refresh_chk.isChecked():
            self.refresh_all()

    def _on_auto_refresh_toggled(self, checked: bool):
        if checked:
            self.refresh_timer.start(2000)
        else:
            self.refresh_timer.stop()

    def refresh_all(self):
        """Fetch active sessions, recent events, and update the UI."""
        # 1. Update active sessions status
        sessions = get_active_mcp_sessions()
        if sessions:
            desc_items = []
            for s in sessions:
                client = s.get("client", "agent")
                pid = s.get("pid", "?")
                profile = s.get("profile", "all")
                desc_items.append(f"🟢 <b>{client}</b> (PID: {pid}, Profile: {profile})")
            self.conn_status_label.setText(" | ".join(desc_items))
        else:
            self.conn_status_label.setText("⚪ Idle — No external agent process currently connected")

        # 2. Update aggregate metrics
        stats = get_mcp_activity_stats()
        self.metric_calls_pill.setText(f"Total Calls: {stats['total_calls']}")
        self.metric_rate_pill.setText(f"Success: {stats['success_rate']}%")
        self.metric_lat_pill.setText(f"Avg Latency: {stats['avg_latency_ms']} ms")

        # 3. Read events & populate table
        events = get_recent_mcp_events(limit=300)
        self._cached_events = events
        self._populate_table(events)

    def _populate_table(self, events: List[Dict[str, Any]]):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)

        query = self.search_filter_edit.text().strip().lower()
        tool_flt = self.tool_filter_combo.currentData() or ""
        status_flt = self.status_filter_combo.currentData() or "all"

        for _row_idx, ev in enumerate(events):
            tool = ev.get("tool", "")
            action = ev.get("action", "")
            client = ev.get("client", "")
            summary = ev.get("summary", "")
            is_ok = ev.get("ok", True)
            is_err = ev.get("is_error", False)

            # Filtering
            if tool_flt and tool != tool_flt:
                continue
            if status_flt == "success" and (not is_ok or is_err):
                continue
            if status_flt == "error" and (is_ok and not is_err):
                continue
            if query:
                combined_text = f"{tool} {action} {client} {summary} {json.dumps(ev.get('args', {}))}".lower()
                if query not in combined_text:
                    continue

            cur_row = self.table.rowCount()
            self.table.insertRow(cur_row)

            # Status Badge
            if not is_ok:
                st_text = "🔴 Error"
                st_color = QColor(C.DANGER_ALT)
            elif is_err:
                st_text = "🟠 Logic"
                st_color = QColor(C.WARNING_AMBER)
            else:
                st_text = "🟢 OK"
                st_color = QColor(C.EMERALD_PRIMARY)

            items = [
                QTableWidgetItem(ev.get("time_short", "")),
                QTableWidgetItem(client),
                QTableWidgetItem(str(ev.get("pid", ""))),
                QTableWidgetItem(tool),
                QTableWidgetItem(action),
                QTableWidgetItem(summary),
                QTableWidgetItem(f"{ev.get('elapsed_ms', 0.0)} ms"),
                QTableWidgetItem(st_text),
            ]

            items[7].setForeground(QBrush(st_color))
            f_st = items[7].font()
            f_st.setBold(True)
            items[7].setFont(f_st)

            # Store full event dictionary on the first item for easy retrieval
            items[0].setData(Qt.UserRole, ev)

            for col_idx, item in enumerate(items):
                self.table.setItem(cur_row, col_idx, item)

        if self.autoscroll_chk.isChecked() and self.table.rowCount() > 0:
            self.table.scrollToItem(self.table.item(0, 0))

    def _apply_table_filters(self):
        self._populate_table(self._cached_events)

    def _on_table_selection_changed(self):
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            self.copy_args_btn.setEnabled(False)
            self.copy_resp_btn.setEnabled(False)
            return

        row = selected_rows[0].row()
        item_0 = self.table.item(row, 0)
        if not item_0:
            return

        ev = item_0.data(Qt.UserRole)
        if not ev:
            return

        self._current_selected_event = ev
        self.copy_args_btn.setEnabled(True)
        self.copy_resp_btn.setEnabled(True)

        tool = ev.get("tool") or "General"
        action = ev.get("action") or ""
        elapsed = ev.get("elapsed_ms", 0.0)
        bytes_count = ev.get("response_chars", 0)

        self.insp_title.setText(
            f"Event: {tool}:{action} | Client: {ev.get('client')} (PID {ev.get('pid')}) | Latency: {elapsed} ms | Size: {bytes_count} chars"
        )

        args = ev.get("args") or {}
        self.args_view.setText(json.dumps(args, indent=2, ensure_ascii=False))

        resp_preview = ev.get("response_preview") or ""
        if not resp_preview and ev.get("error"):
            resp_preview = f"Error: {ev.get('error')}"
        self.resp_view.setText(resp_preview)

    def _copy_selected_args(self):
        if self._current_selected_event:
            args = self._current_selected_event.get("args") or {}
            QApplication.clipboard().setText(json.dumps(args, indent=2, ensure_ascii=False))
            self._copy_snippet_with_feedback(json.dumps(args, indent=2, ensure_ascii=False), self.copy_args_btn)

    def _copy_selected_response(self):
        if self._current_selected_event:
            preview = self._current_selected_event.get("response_preview") or self._current_selected_event.get("error") or ""
            QApplication.clipboard().setText(preview)
            self._copy_snippet_with_feedback(preview, self.copy_resp_btn)

    def _on_clear_history_clicked(self):
        reply = QMessageBox.question(
            self,
            "Clear MCP Activity History",
            "Are you sure you want to clear all logged MCP events?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            clear_mcp_activity_log()
            self.refresh_all()
            self.args_view.clear()
            self.resp_view.clear()
            self.insp_title.setText("Inspection Pane: History cleared")
