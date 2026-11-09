# --- File: src/core/ui/mcp_monitor_dialog.py ---
"""
MCP Server & External AI Agent Monitor Dialog.

Provides real-time visibility into MCP server tool calls, queries, and document
access by external AI coding and standards agents (Claude Desktop, Codex CLI,
Claude CLI, Cursor, Windsurf), ready-to-copy client configurations, and an
interactive domain tool playground.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QDateTime, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
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
    get_mcp_client_configurations,
    get_recent_mcp_events,
)
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    create_app_icon,
)
from modules.assistant.core.tools.external_agent_tools import (
    PROFILE_TOOLS,
    build_external_agent_tool_registry,
)

logger = logging.getLogger(__name__)

# Example arguments for interactive tool playground
DEFAULT_PLAYGROUND_ARGUMENTS: Dict[str, Dict[str, Dict[str, Any]]] = {
    "spec_text_and_clauses": {
        "search_text": {"spec_number": "23.501", "query": "URSP", "limit": 3},
        "get_clause": {"spec_number": "23.501", "clause_id": "5.15.1"},
        "diff_versions": {"spec_number": "23.501", "old_version": "19.0.0", "new_version": "20.2.0"},
        "evolution_matrix": {"spec_number": "23.501", "clause_id": "5.15.1"},
        "extract_clause": {"spec_number": "23.501", "clause_id": "5.15.1", "export_format": "markdown"},
    },
    "protocol_inspector": {
        "list_protocols": {},
        "find_message": {"protocol": "5GS_NAS", "query": "REGISTRATION REQUEST"},
        "find_ie": {"protocol": "5GS_NAS", "query": "5GS tracking area identity"},
        "diff_releases": {"protocol": "5GS_NAS", "old_version": "18.0.0", "new_version": "19.0.0"},
    },
    "meetings_and_tdocs": {
        "search_meetings": {"wg": "SA2", "limit": 3},
        "search_tdocs": {"wg": "SA2", "meeting_number": "174", "limit": 3},
        "read_tdoc_markdown": {"tdoc_id": "S2-2601720"},
        "meeting_stats": {"wg": "SA2", "meeting_number": "174"},
        "get_or_set_stance": {"tdoc_id": "S2-2601720"},
    },
    "onsite_meeting_cockpit": {
        "network_status": {"wg": "SA2"},
        "where_should_we_go_now": {"wg": "SA2"},
        "browse_lan_server": {"subfolder": "Docs"},
    },
    "spec_history_and_attribution": {
        "query_history": {"spec_number": "23.501", "limit": 3},
        "attribute_clause": {"spec_number": "23.501", "clause_id": "5.15.1", "limit": 3},
        "rank_origin_crs": {"spec_number": "23.501", "query": "URSP", "limit": 3},
        "tdoc_dossier": {"tdoc_id": "S2-2601720"},
    },
    "work_items_and_emails": {
        "search_work_items": {"query": "6G", "limit": 3},
    },
    "cr_pcr_generator": {
        "preview_filename": {
            "mode": "pCR",
            "spec_number": "23.501",
            "clauses": "5.15.1",
            "tdoc_id": "S2-2601720",
        },
    },
}


class McpMonitorDialog(QDialog):
    """Real-time Model Context Protocol (MCP) server & agent monitor cockpit."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(None)  # Top-level window without forced modal lock
        self._parent_ref = parent
        self.setWindowTitle("🤖 MCP Server & External AI Agent Cockpit")
        self.setWindowIcon(create_app_icon())
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.resize(1060, 720)

        self._cached_events: List[Dict[str, Any]] = []
        self._current_selected_event: Optional[Dict[str, Any]] = None
        self._tool_registry = None

        self._setup_ui()

        # Polling timer for live updates
        self.refresh_timer = QTimer(self)
        self.refresh_timer.setInterval(2000)
        self.refresh_timer.timeout.connect(self._on_timer_refresh)
        self.refresh_timer.start()

        # Initial data load
        self.refresh_all()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # 1. Top System & Connection Banner
        banner_frame = QFrame()
        banner_frame.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F8FAFC, stop:1 #EFF6FF);
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 4px;
            }
        """)
        banner_layout = QHBoxLayout(banner_frame)
        banner_layout.setContentsMargins(12, 10, 12, 10)

        # Left: Title, Description & Connection Status
        left_layout = QVBoxLayout()
        left_layout.setSpacing(4)

        header_title = QLabel("🤖 Model Context Protocol (MCP) & AI Agent Monitor")
        header_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #0F172A;")
        left_layout.addWidget(header_title)

        self.conn_status_label = QLabel("⚪ Detecting active agent sessions...")
        self.conn_status_label.setStyleSheet("font-size: 12px; font-weight: 600; color: #475569;")
        left_layout.addWidget(self.conn_status_label)

        banner_layout.addLayout(left_layout, stretch=1)

        # Right: Aggregate Metric Pills & Quick Actions
        right_layout = QHBoxLayout()
        right_layout.setSpacing(8)

        self.metric_calls_pill = QLabel("Calls: 0")
        self.metric_calls_pill.setStyleSheet("""
            background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: #1E293B;
        """)
        self.metric_rate_pill = QLabel("Success: 100%")
        self.metric_rate_pill.setStyleSheet("""
            background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: #059669;
        """)
        self.metric_lat_pill = QLabel("Avg Latency: 0 ms")
        self.metric_lat_pill.setStyleSheet("""
            background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: #2563EB;
        """)

        right_layout.addWidget(self.metric_calls_pill)
        right_layout.addWidget(self.metric_rate_pill)
        right_layout.addWidget(self.metric_lat_pill)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background: #FFFFFF; border: 1px solid #94A3B8; border-radius: 5px;
                padding: 4px 12px; font-weight: 600; color: #1E293B;
            }
            QPushButton:hover { background: #F1F5F9; border-color: #64748B; }
        """)
        self.refresh_btn.clicked.connect(self.refresh_all)
        right_layout.addWidget(self.refresh_btn)

        banner_layout.addLayout(right_layout)
        main_layout.addWidget(banner_frame)

        # 2. Main Tab Widget
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                background: #FFFFFF;
                top: -1px;
            }
            QTabBar::tab {
                background: #F1F5F9;
                border: 1px solid #CBD5E1;
                border-bottom: none;
                padding: 8px 18px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-size: 12px;
                font-weight: 600;
                color: #475569;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-bottom: 2px solid #2563EB;
                color: #1E40AF;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background: #E2E8F0;
            }
        """)

        # Tab 1: Live Activity Stream
        self.activity_tab = QWidget()
        self._setup_activity_tab(self.activity_tab)
        self.tabs.addTab(self.activity_tab, "📊 Live Activity Stream")

        # Tab 2: Client Setup & Configurations
        self.setup_tab = QWidget()
        self._setup_config_tab(self.setup_tab)
        self.tabs.addTab(self.setup_tab, "⚙️ AI Agent Client Setup")

        # Tab 3: Interactive Tool Playground
        self.playground_tab = QWidget()
        self._setup_playground_tab(self.playground_tab)
        self.tabs.addTab(self.playground_tab, "🧪 Tool Playground")

        main_layout.addWidget(self.tabs, stretch=1)

        # Bottom status row
        bottom_row = QHBoxLayout()
        self.auto_refresh_chk = QCheckBox("Auto-refresh stream (2s)")
        self.auto_refresh_chk.setChecked(True)
        self.auto_refresh_chk.toggled.connect(self._on_auto_refresh_toggled)
        bottom_row.addWidget(self.auto_refresh_chk)

        bottom_row.addStretch()

        self.clear_btn = QPushButton("🗑️ Clear Activity History")
        self.clear_btn.setStyleSheet("padding: 3px 10px; font-size: 11px;")
        self.clear_btn.clicked.connect(self._on_clear_history_clicked)
        bottom_row.addWidget(self.clear_btn)

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(80)
        close_btn.clicked.connect(self.close)
        bottom_row.addWidget(close_btn)

        main_layout.addLayout(bottom_row)

    # =====================================================================
    # Tab 1: Activity Stream Setup
    # =====================================================================

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
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E2E8F0;
                gridline-color: #F1F5F9;
                font-size: 12px;
            }
            QTableWidget::item:selected {
                background: #EFF6FF;
                color: #1E3A8A;
            }
        """)
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)
        splitter.addWidget(self.table)

        # Bottom Inspector Frame
        inspector_frame = QFrame()
        inspector_frame.setStyleSheet("background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px;")
        insp_layout = QVBoxLayout(inspector_frame)
        insp_layout.setContentsMargins(8, 8, 8, 8)
        insp_layout.setSpacing(6)

        insp_header = QHBoxLayout()
        self.insp_title = QLabel("Inspection Pane: Select an event row above to inspect parameters and payload")
        self.insp_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #334155;")
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
        self.args_view.setFont(QFont("Consolas", 10))
        self.args_view.setPlaceholderText("No arguments")
        lp_layout.addWidget(self.args_view)
        text_splitter.addWidget(left_pane)

        right_pane = QWidget()
        rp_layout = QVBoxLayout(right_pane)
        rp_layout.setContentsMargins(0, 0, 0, 0)
        rp_layout.addWidget(QLabel("Response Payload / Preview:"))
        self.resp_view = QTextEdit()
        self.resp_view.setReadOnly(True)
        self.resp_view.setFont(QFont("Consolas", 10))
        self.resp_view.setPlaceholderText("No response preview")
        rp_layout.addWidget(self.resp_view)
        text_splitter.addWidget(right_pane)

        text_splitter.setSizes([380, 560])
        insp_layout.addWidget(text_splitter)

        splitter.addWidget(inspector_frame)
        splitter.setSizes([380, 220])
        layout.addWidget(splitter)

    # =====================================================================
    # Tab 2: Agent Client Setup
    # =====================================================================

    def _setup_config_tab(self, parent: QWidget):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        info_box = QLabel(
            "External AI agents communicate with 3GPP Tools via standard input/output (stdio) JSON-RPC. "
            "Copy the pre-configured snippets below into your AI client settings. All absolute Python "
            "and application paths are automatically resolved."
        )
        info_box.setWordWrap(True)
        info_box.setStyleSheet("background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 5px; padding: 8px;")
        layout.addWidget(info_box)

        cfg_tabs = QTabWidget()

        # Generate configuration dictionary
        cfgs = get_mcp_client_configurations()
        self._client_cfgs = cfgs

        # 1. Claude Desktop Card
        claude_card = self._create_client_card(
            title="Claude Desktop",
            snippet=cfgs["claude_desktop"]["config_json"],
            instructions=cfgs["claude_desktop"]["instructions"],
            target_path=cfgs["claude_desktop"]["target_file"],
            is_json=True,
        )
        cfg_tabs.addTab(claude_card, "Claude Desktop")

        # 2. Codex CLI Card
        codex_card = self._create_client_card(
            title="OpenAI Codex CLI",
            snippet=cfgs["codex_cli"]["config_json"],
            instructions=cfgs["codex_cli"]["instructions"],
            command_line=cfgs["codex_cli"]["cli_command"],
            is_json=True,
        )
        cfg_tabs.addTab(codex_card, "Codex CLI")

        # 3. Claude CLI Card
        claude_cli_card = self._create_client_card(
            title="Anthropic Claude CLI",
            snippet=cfgs["claude_cli"]["cli_command"],
            instructions=cfgs["claude_cli"]["instructions"],
            is_json=False,
        )
        cfg_tabs.addTab(claude_cli_card, "Claude CLI")

        # 4. Cursor / Windsurf Card
        cursor_card = self._create_client_card(
            title="Cursor / Windsurf",
            snippet=cfgs["cursor"]["command_line"],
            instructions=cfgs["cursor"]["instructions"],
            is_json=False,
        )
        cfg_tabs.addTab(cursor_card, "Cursor / Windsurf")

        # 5. Windows Batch Launcher Card
        batch_card = self._create_client_card(
            title="Zero-Config Batch Launcher (Launch_MCP.cmd)",
            snippet=cfgs["claude_desktop"]["cmd_config_json"],
            instructions="Use this variant if you prefer running via the bundled Launch_MCP.cmd launcher script.",
            target_path=cfgs["paths"]["cmd_launcher"],
            is_json=True,
        )
        cfg_tabs.addTab(batch_card, "Batch Launcher (.cmd)")

        layout.addWidget(cfg_tabs, stretch=1)

    def _create_client_card(
        self,
        title: str,
        snippet: str,
        instructions: str,
        target_path: Optional[str] = None,
        command_line: Optional[str] = None,
        is_json: bool = True,
    ) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        inst_lbl = QLabel(instructions)
        inst_lbl.setWordWrap(True)
        inst_lbl.setStyleSheet("font-size: 12px; color: #334155; font-weight: 500;")
        layout.addWidget(inst_lbl)

        if target_path:
            path_row = QHBoxLayout()
            p_lbl = QLabel(f"<b>Target Config File:</b> <code>{target_path}</code>")
            p_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            path_row.addWidget(p_lbl)
            path_row.addStretch()

            open_dir_btn = QPushButton("📂 Show in Explorer")
            open_dir_btn.setStyleSheet("padding: 2px 8px; font-size: 11px;")
            open_dir_btn.clicked.connect(lambda: self._open_in_explorer(target_path))
            path_row.addWidget(open_dir_btn)
            layout.addLayout(path_row)

        snippet_edit = QTextEdit()
        snippet_edit.setReadOnly(True)
        snippet_edit.setFont(QFont("Consolas", 10))
        snippet_edit.setText(snippet)
        snippet_edit.setStyleSheet("background: #0F172A; color: #E2E8F0; border-radius: 4px; padding: 6px;")
        layout.addWidget(snippet_edit, stretch=1)

        btn_row = QHBoxLayout()
        copy_btn = QPushButton("📋 Copy Configuration Snippet")
        copy_btn.setStyleSheet("""
            QPushButton {
                background: #2563EB; color: white; border: none; border-radius: 4px;
                padding: 6px 16px; font-weight: bold;
            }
            QPushButton:hover { background: #1D4ED8; }
        """)
        copy_btn.clicked.connect(lambda: self._copy_snippet_with_feedback(snippet, copy_btn))
        btn_row.addWidget(copy_btn)

        if command_line:
            copy_cmd_btn = QPushButton("📋 Copy CLI Command")
            copy_cmd_btn.setStyleSheet("padding: 6px 12px; font-weight: 600;")
            copy_cmd_btn.clicked.connect(lambda: self._copy_snippet_with_feedback(command_line, copy_cmd_btn))
            btn_row.addWidget(copy_cmd_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)
        return widget

    def _open_in_explorer(self, target_path: str):
        try:
            p = Path(target_path)
            target_dir = p.parent if p.is_file() or not p.exists() else p
            if not target_dir.exists():
                target_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(target_dir))
        except Exception as exc:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{exc}")

    def _copy_snippet_with_feedback(self, text: str, btn: QPushButton):
        QApplication.clipboard().setText(text)
        orig_text = btn.text()
        btn.setText("✓ Copied to Clipboard!")
        btn.setStyleSheet("background: #059669; color: white; border: none; border-radius: 4px; padding: 6px 16px; font-weight: bold;")
        QTimer.singleShot(1500, lambda: self._reset_copy_btn(btn, orig_text))

    def _reset_copy_btn(self, btn: QPushButton, orig_text: str):
        btn.setText(orig_text)
        btn.setStyleSheet("background: #2563EB; color: white; border: none; border-radius: 4px; padding: 6px 16px; font-weight: bold;")

    # =====================================================================
    # Tab 3: Tool Playground Setup
    # =====================================================================

    def _setup_playground_tab(self, parent: QWidget):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        intro = QLabel(
            "Test-drive any of the 12 domain tools and 58 actions directly within the desktop application. "
            "Select a tool, modify arguments as needed, and hit 'Execute Test Call' to inspect live output."
        )
        intro.setStyleSheet("color: #475569; font-size: 12px;")
        layout.addWidget(intro)

        # Tool & Action Selector
        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("Tool:"))
        self.play_tool_combo = QComboBox()
        for tname in PROFILE_TOOLS.get("all", []):
            self.play_tool_combo.addItem(tname, tname)
        self.play_tool_combo.currentIndexChanged.connect(self._on_play_tool_changed)
        sel_row.addWidget(self.play_tool_combo, stretch=1)

        sel_row.addWidget(QLabel("Action:"))
        self.play_action_combo = QComboBox()
        self.play_action_combo.currentIndexChanged.connect(self._on_play_action_changed)
        sel_row.addWidget(self.play_action_combo, stretch=1)

        self.play_exec_btn = QPushButton("▶️ Execute Test Call")
        self.play_exec_btn.setStyleSheet("""
            QPushButton {
                background: #059669; color: white; border: none; border-radius: 4px;
                padding: 6px 18px; font-weight: bold; font-size: 12px;
            }
            QPushButton:hover { background: #047857; }
        """)
        self.play_exec_btn.clicked.connect(self._execute_playground_call)
        sel_row.addWidget(self.play_exec_btn)

        layout.addLayout(sel_row)

        # Splitter: Input Arguments JSON (Left) and Output JSON (Right)
        play_splitter = QSplitter(Qt.Horizontal)

        in_widget = QWidget()
        in_layout = QVBoxLayout(in_widget)
        in_layout.setContentsMargins(0, 0, 0, 0)
        in_layout.addWidget(QLabel("Arguments JSON:"))
        self.play_args_edit = QTextEdit()
        self.play_args_edit.setFont(QFont("Consolas", 10))
        self.play_args_edit.setText("{\n  \"action\": \"search_text\"\n}")
        in_layout.addWidget(self.play_args_edit)
        play_splitter.addWidget(in_widget)

        out_widget = QWidget()
        out_layout = QVBoxLayout(out_widget)
        out_layout.setContentsMargins(0, 0, 0, 0)
        self.play_out_title = QLabel("Execution Output & Telemetry:")
        out_layout.addWidget(self.play_out_title)
        self.play_output_edit = QTextEdit()
        self.play_output_edit.setReadOnly(True)
        self.play_output_edit.setFont(QFont("Consolas", 10))
        self.play_output_edit.setPlaceholderText("Execution result will appear here...")
        out_layout.addWidget(self.play_output_edit)
        play_splitter.addWidget(out_widget)

        play_splitter.setSizes([380, 560])
        layout.addWidget(play_splitter, stretch=1)

        # Trigger initial tool populating
        self._on_play_tool_changed(0)

    def _ensure_registry(self):
        if self._tool_registry is None:
            self._tool_registry = build_external_agent_tool_registry(profile="all", read_only=False)
        return self._tool_registry

    def _on_play_tool_changed(self, _index: int):
        tool_name = self.play_tool_combo.currentText()
        registry = self._ensure_registry()
        tool = registry.get(tool_name)
        self.play_action_combo.blockSignals(True)
        self.play_action_combo.clear()
        if tool:
            for arg in tool.arguments:
                if arg.name == "action" and arg.enum:
                    for act in arg.enum:
                        self.play_action_combo.addItem(act, act)
                    break
        self.play_action_combo.blockSignals(False)
        self._on_play_action_changed(0)

    def _on_play_action_changed(self, _index: int):
        tool_name = self.play_tool_combo.currentText()
        action = self.play_action_combo.currentText()
        preset = DEFAULT_PLAYGROUND_ARGUMENTS.get(tool_name, {}).get(action, {})
        full_args = {"action": action, **preset}
        self.play_args_edit.setText(json.dumps(full_args, indent=2))

    def _execute_playground_call(self):
        tool_name = self.play_tool_combo.currentText()
        raw_text = self.play_args_edit.toPlainText().strip()
        try:
            args = json.loads(raw_text) if raw_text else {}
        except Exception as exc:
            QMessageBox.critical(self, "Invalid JSON", f"Arguments must be valid JSON:\n{exc}")
            return

        self.play_exec_btn.setEnabled(False)
        self.play_out_title.setText("⏳ Executing tool...")
        QApplication.processEvents()

        t0 = time.perf_counter()
        registry = self._ensure_registry()
        res = registry.execute(tool_name, args)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        self.play_exec_btn.setEnabled(True)
        if res.ok:
            is_err = isinstance(res.data, dict) and res.data.get("ok") is False
            status_emoji = "🟠 Logic Note" if is_err else "🟢 Success"
            self.play_out_title.setText(f"{status_emoji} | Finished in {elapsed_ms:.1f} ms")
            self.play_output_edit.setText(json.dumps(res.data, ensure_ascii=False, indent=2, default=str))
        else:
            err_msg = getattr(res, "error", None) or getattr(res, "message", "Tool execution failed")
            self.play_out_title.setText(f"🔴 Failed in {elapsed_ms:.1f} ms")
            self.play_output_edit.setText(f"Error:\n{err_msg}")

        # Trigger activity refresh so this test call appears in the table
        QTimer.singleShot(200, self.refresh_all)

    # =====================================================================
    # Data Refresh & Live Updates
    # =====================================================================

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

        for row_idx, ev in enumerate(events):
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
                st_color = QColor("#DC2626")
            elif is_err:
                st_text = "🟠 Logic"
                st_color = QColor("#D97706")
            else:
                st_text = "🟢 OK"
                st_color = QColor("#059669")

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
            items[7].setFont(QFont("Segoe UI", 9, QFont.Bold))

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
