"""McpMonitorDialog shell combining Activity Stream, Client Setup, and Playground mixins."""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.mcp_monitor.mixins import (
    McpActivityStreamMixin,
    McpClientSetupMixin,
    McpToolPlaygroundMixin,
)
from core.ui.ui_components import create_app_icon
from core.ui.theme import C

logger = logging.getLogger("core.ui.mcp_monitor_dialog")


class McpMonitorDialog(
    McpActivityStreamMixin,
    McpClientSetupMixin,
    McpToolPlaygroundMixin,
    QDialog,
):
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
        banner_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {C.BG_SURFACE}, stop:1 {C.INFO_SOFT});
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 8px;
                padding: 4px;
            }}
        """)
        banner_layout = QHBoxLayout(banner_frame)
        banner_layout.setContentsMargins(12, 10, 12, 10)

        # Left: Title, Description & Connection Status
        left_layout = QVBoxLayout()
        left_layout.setSpacing(4)

        header_title = QLabel("🤖 Model Context Protocol (MCP) & AI Agent Monitor")
        header_title.setStyleSheet(f"font-size: 15px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        left_layout.addWidget(header_title)

        self.conn_status_label = QLabel("⚪ Detecting active agent sessions...")
        self.conn_status_label.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {C.TEXT_SECONDARY};")
        left_layout.addWidget(self.conn_status_label)

        banner_layout.addLayout(left_layout, stretch=1)

        # Right: Aggregate Metric Pills & Quick Actions
        right_layout = QHBoxLayout()
        right_layout.setSpacing(8)

        self.metric_calls_pill = QLabel("Calls: 0")
        self.metric_calls_pill.setStyleSheet(f"""
            background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: {C.TEXT_STRONG};
        """)
        self.metric_rate_pill = QLabel("Success: 100%")
        self.metric_rate_pill.setStyleSheet(f"""
            background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: {C.EMERALD_PRIMARY};
        """)
        self.metric_lat_pill = QLabel("Avg Latency: 0 ms")
        self.metric_lat_pill.setStyleSheet(f"""
            background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px;
            padding: 4px 10px; font-size: 12px; font-weight: bold; color: {C.ACCENT_ALT};
        """)

        right_layout.addWidget(self.metric_calls_pill)
        right_layout.addWidget(self.metric_rate_pill)
        right_layout.addWidget(self.metric_lat_pill)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background: {C.BG_CARD}; border: 1px solid {C.TEXT_FAINT}; border-radius: 5px;
                padding: 4px 12px; font-weight: 600; color: {C.TEXT_STRONG};
            }}
            QPushButton:hover {{ background: {C.BG_CANVAS}; border-color: {C.TEXT_MUTED}; }}
        """)
        self.refresh_btn.clicked.connect(self.refresh_all)
        right_layout.addWidget(self.refresh_btn)

        banner_layout.addLayout(right_layout)
        main_layout.addWidget(banner_frame)

        # 2. Main Tab Widget
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                background: {C.BG_CARD};
                top: -1px;
            }}
            QTabBar::tab {{
                background: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_STRONG};
                border-bottom: none;
                padding: 8px 18px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-size: 12px;
                font-weight: 600;
                color: {C.TEXT_SECONDARY};
            }}
            QTabBar::tab:selected {{
                background: {C.BG_CARD};
                border-bottom: 2px solid {C.ACCENT_ALT};
                color: {C.INFO_DARK};
                font-weight: bold;
            }}
            QTabBar::tab:hover:!selected {{
                background: {C.BORDER_SUBTLE};
            }}
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
