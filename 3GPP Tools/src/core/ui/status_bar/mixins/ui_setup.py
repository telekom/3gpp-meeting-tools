# --- File: src/core/ui/status_bar/mixins/ui_setup.py ---
"""Status bar pill construction, system tools menu, update badge, and drawer/dock badges."""

from PyQt5.QtCore import QSize, Qt, QTimer
from PyQt5.QtWidgets import QLabel, QMenu, QPushButton

from core.ai.ollama_client import load_ollama_config
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, T

from ..styles import _status_pill_qss


class StatusBarUiSetupMixin:
    """Mixin creating permanent status bar pills, menus, and drawer/update badges."""

    def _setup_indicators(self):
        saved_cfg = load_ollama_config()
        saved_model = saved_cfg.get("selected_model", "").strip()
        self._ollama_initial_model = saved_model
        initial_ollama_text = f"🦙 {saved_model}" if saved_model else "🦙 Ollama"

        # 0. App Update Pill (shown only when a new commit is available on master)
        self.update_indicator = QPushButton("🆕 Update Available")
        self.update_indicator.setObjectName("statusUpdateAvailableBtn")
        self.update_indicator.setCursor(Qt.PointingHandCursor)
        self.update_indicator.setIcon(
            lucide_icon("download", color=C.ACCENT_PRIMARY, size=14)
        )
        self.update_indicator.setIconSize(QSize(14, 14))
        self.update_indicator.setStyleSheet(
            _status_pill_qss(
                bg=C.ACCENT_SOFT,
                fg=C.ACCENT_PRIMARY,
                bd=C.ACCENT_RING,
                hover_bg=C.ACCENT_SOFT_HOVER,
                hover_bd=C.ACCENT_PRIMARY,
                compact=self._minimal_mode,
            )
        )
        self.update_indicator.clicked.connect(self.update_check_requested.emit)
        self.update_indicator.setVisible(False)
        self.addPermanentWidget(self.update_indicator)

        # 1. Ollama Indicator Pill
        self.ollama_indicator = QPushButton(initial_ollama_text)
        self.ollama_indicator.setCursor(Qt.PointingHandCursor)
        self.ollama_indicator.setToolTip("Click to configure Ollama LLM connection")
        self.ollama_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))
        self.ollama_indicator.clicked.connect(self.open_ollama_settings)
        self.addPermanentWidget(self.ollama_indicator)

        # 2. MCP AI Agent Indicator Pill
        self.mcp_indicator = QPushButton("🤖 MCP: Idle")
        self.mcp_indicator.setCursor(Qt.PointingHandCursor)
        self.mcp_indicator.setToolTip("Click to inspect MCP Server activity & external AI agents (Claude, Codex)")
        self.mcp_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))
        self.mcp_indicator.clicked.connect(self.open_mcp_monitor)
        self.addPermanentWidget(self.mcp_indicator)

        # 3. Proxy Indicator Pill
        self.proxy_indicator = QPushButton("🌐 Direct")
        self.proxy_indicator.setCursor(Qt.PointingHandCursor)
        self.proxy_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))
        self.proxy_indicator.clicked.connect(self.open_proxy_settings)
        self.addPermanentWidget(self.proxy_indicator)

        # 4. Network Indicator
        self.network_indicator = QLabel("📶 Network")
        self.network_indicator.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; padding: 0 8px;"
        )
        self.addPermanentWidget(self.network_indicator)

        # 5. System Tools Menu Pill (Task Manager, Database, Proxy, Network, Java, MCP)
        self.system_btn = QPushButton("System")
        self.system_btn.setObjectName("statusSystemMenuBtn")
        self.system_btn.setCursor(Qt.PointingHandCursor)
        self.system_btn.setIcon(
            lucide_icon("wrench", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.system_btn.setIconSize(QSize(14, 14))
        self.system_btn.setToolTip("System diagnostics and maintenance tools (COM Task Manager, Database, Proxy, Network, Java, MCP)")
        self._render_system_btn()
        self.system_menu = QMenu(self)
        self.system_menu.setToolTipsVisible(True)
        self._populate_system_menu()
        self.system_btn.setMenu(self.system_menu)
        self.addPermanentWidget(self.system_btn)

        # 6. Right Side Panel Dock Toggles (Assistant & Quick Spec Search)
        self.assistant_dock_btn = QPushButton("Assistant")
        self.assistant_dock_btn.setObjectName("statusAssistantDockBtn")
        self.assistant_dock_btn.setCheckable(True)
        self.assistant_dock_btn.setCursor(Qt.PointingHandCursor)
        self.assistant_dock_btn.setIcon(
            lucide_icon("sparkles", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.assistant_dock_btn.setIconSize(QSize(14, 14))
        self.assistant_dock_btn.setToolTip("Show or hide Docked Assistant side panel (Ctrl+Shift+A)")
        self.assistant_dock_btn.clicked.connect(self.assistant_dock_requested.emit)
        self.addPermanentWidget(self.assistant_dock_btn)

        self.search_dock_btn = QPushButton("Quick Search")
        self.search_dock_btn.setObjectName("statusSearchDockBtn")
        self.search_dock_btn.setCheckable(True)
        self.search_dock_btn.setCursor(Qt.PointingHandCursor)
        self.search_dock_btn.setIcon(
            lucide_icon("search", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.search_dock_btn.setIconSize(QSize(14, 14))
        self.search_dock_btn.setToolTip("Show or hide Quick Spec Search side panel (Ctrl+Shift+F)")
        self.search_dock_btn.clicked.connect(self.search_dock_requested.emit)
        self.addPermanentWidget(self.search_dock_btn)
        self.update_right_panel_badges(False, "assistant")

        # 7. Collapsible Bottom Drawer Toggle (Console & Queue + Badges)
        self.drawer_toggle_btn = QPushButton("Console && Queue")
        self.drawer_toggle_btn.setObjectName("statusDrawerToggleBtn")
        self.drawer_toggle_btn.setCheckable(True)
        self.drawer_toggle_btn.setCursor(Qt.PointingHandCursor)
        self.drawer_toggle_btn.setIcon(
            lucide_icon("terminal", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.drawer_toggle_btn.setIconSize(QSize(14, 14))
        self.drawer_toggle_btn.setToolTip("Show or hide the Console & Queue drawer (Ctrl+J)")
        self.drawer_toggle_btn.clicked.connect(self.drawer_toggle_requested.emit)
        self.addPermanentWidget(self.drawer_toggle_btn)
        self.update_drawer_badge(False, 0, 0)

        # 8. Live World Clock & Active 3GPP Meeting Venues Pill
        self.clock_indicator = QPushButton("🕒 --:--")
        self.clock_indicator.setObjectName("statusClockBtn")
        self.clock_indicator.setCursor(Qt.PointingHandCursor)
        self.clock_indicator.setStyleSheet(_status_pill_qss(bg=C.BG_SURFACE, fg=C.TEXT_BODY))
        self.clock_indicator.clicked.connect(self.open_world_clock)
        self.addPermanentWidget(self.clock_indicator)

        self.clock_timer = QTimer(self)
        self.clock_timer.setInterval(1000)
        self.clock_timer.timeout.connect(self.update_clock_indicator)
        self.clock_timer.start()
        self.update_clock_indicator()

        # Synchronize proxy on boot
        self.sync_proxy_environment_from_session()
        self.update_proxy_indicator()

        # MCP status update timer
        self._mcp_dialog = None
        self.mcp_timer = QTimer(self)
        self.mcp_timer.setInterval(3000)
        self.mcp_timer.timeout.connect(self.update_mcp_indicator)
        self.mcp_timer.start()
        self.update_mcp_indicator()

    def _populate_system_menu(self) -> None:
        """Populate the status-bar System tools dropdown menu."""
        self.system_menu.clear()
        self.system_menu.addAction("🖥️ Task Manager...", self.task_manager_requested.emit)
        self.system_menu.addAction("🗄️ Database...", self.db_maintenance_requested.emit)
        self.system_menu.addAction("📡 Proxy...", self.open_proxy_settings)
        self.system_menu.addAction("⚙️ Network...", self.network_config_requested.emit)
        self.system_menu.addAction("☕ Java...", self.java_info_requested.emit)
        self.system_menu.addAction("🤖 MCP...", self.open_mcp_monitor)
        self.system_menu.addSeparator()
        restart_act = self.system_menu.addAction("🔁 Restart Application", self.restart_requested.emit)
        restart_act.setToolTip("Save session state and restart 3GPP Delegate Tools (Ctrl+Shift+R)")

    def _render_system_btn(self) -> None:
        self.system_btn.setText("" if self._minimal_mode else "System")
        self.system_btn.setStyleSheet(
            _status_pill_qss(bg=C.BG_SURFACE, fg=C.TEXT_BODY, compact=self._minimal_mode)
        )

    def _render_update_indicator(self) -> None:
        visible, _count, latest_short_sha, summary = self._update_state
        if not visible:
            self.update_indicator.setVisible(False)
            return
        if self._minimal_mode:
            label = f"{_count}" if _count > 0 else ""
        else:
            label = self._full_update_text()
        self.update_indicator.setText(label)
        tip = f"New update available on GitHub master ({latest_short_sha})"
        if summary:
            tip += f"\nLatest: {summary}"
        tip += "\n\nClick to view changes and update"
        self.update_indicator.setToolTip(tip)
        self.update_indicator.setStyleSheet(
            _status_pill_qss(
                bg=C.ACCENT_SOFT,
                fg=C.ACCENT_PRIMARY,
                bd=C.ACCENT_RING,
                hover_bg=C.ACCENT_SOFT_HOVER,
                hover_bd=C.ACCENT_PRIMARY,
                compact=self._minimal_mode,
            )
        )
        self.update_indicator.setVisible(True)

    def set_update_available(
        self, count: int, latest_short_sha: str = "", summary: str = ""
    ) -> None:
        """Display the clickable update badge in the status bar."""
        self._update_state = (True, int(count), str(latest_short_sha), str(summary))
        self._render_update_indicator()
        self._recheck_responsive_fit()

    def clear_update_available(self) -> None:
        """Hide the status-bar update badge."""
        self._update_state = (False, 0, "", "")
        self._render_update_indicator()
        self._recheck_responsive_fit()

    def update_drawer_badge(self, opened: bool, queue_count: int, unread_errors: int) -> None:
        """Update the status-bar Console & Queue pill text, active state, and error/queue badges."""
        self._drawer_open = bool(opened)
        self._drawer_queue_count = max(0, int(queue_count))
        self._drawer_unread_errors = max(0, int(unread_errors))

        if self._minimal_mode:
            if self._drawer_unread_errors > 0:
                text = f"{self._drawer_unread_errors}!"
            elif self._drawer_queue_count > 0:
                text = f"{self._drawer_queue_count}"
            else:
                text = ""
            pad = "1px 5px"
        else:
            text = self._full_drawer_text()
            pad = "1px 8px"

        self.drawer_toggle_btn.blockSignals(True)
        self.drawer_toggle_btn.setChecked(self._drawer_open)
        self.drawer_toggle_btn.blockSignals(False)
        self.drawer_toggle_btn.setText(text)

        if self._drawer_unread_errors > 0 and not self._drawer_open:
            bg, fg, bd = C.DANGER_SOFT, C.DANGER, C.DANGER_BORDER
            hover_bg = C.DANGER_SOFT_HOVER
        elif self._drawer_open:
            bg, fg, bd = C.ACCENT_SOFT, C.ACCENT_PRIMARY, C.ACCENT_RING
            hover_bg = C.ACCENT_SOFT_HOVER
        else:
            bg, fg, bd = "transparent", C.TEXT_SECONDARY, C.BORDER_STRONG
            hover_bg = C.BG_CANVAS

        self.drawer_toggle_btn.setStyleSheet(f"""
            QPushButton#statusDrawerToggleBtn {{
                background-color: {bg};
                color: {fg};
                border: 1px solid {bd};
                border-radius: {R.SM}px;
                padding: {pad};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#statusDrawerToggleBtn:hover {{
                background-color: {hover_bg};
                border-color: {C.BORDER_HOVER};
            }}
        """)
        self._recheck_responsive_fit()

    def update_right_panel_badges(self, opened: bool, mode: str) -> None:
        """Synchronize the status-bar Assistant & Quick Search dock pill styles."""
        self._right_panel_open = bool(opened)
        self._right_panel_mode = str(mode)
        pad = "1px 5px" if self._minimal_mode else "1px 8px"
        for btn, obj_name, target_mode, full_label in (
            (self.assistant_dock_btn, "statusAssistantDockBtn", "assistant", "Assistant"),
            (self.search_dock_btn, "statusSearchDockBtn", "spec_search", "Quick Search"),
        ):
            active = bool(opened and mode == target_mode)
            btn.blockSignals(True)
            btn.setChecked(active)
            btn.blockSignals(False)
            btn.setText("" if self._minimal_mode else full_label)
            if active:
                bg, fg, bd, hover_bg = C.ACCENT_SOFT, C.ACCENT_PRIMARY, C.ACCENT_RING, C.ACCENT_SOFT_HOVER
            else:
                bg, fg, bd, hover_bg = "transparent", C.TEXT_SECONDARY, C.BORDER_STRONG, C.BG_CANVAS
            btn.setStyleSheet(f"""
                QPushButton#{obj_name} {{
                    background-color: {bg};
                    color: {fg};
                    border: 1px solid {bd};
                    border-radius: {R.SM}px;
                    padding: {pad};
                    font-family: {T.FONT_FAMILY_UI};
                    font-size: {T.SIZE_CAPTION}px;
                    font-weight: 600;
                }}
                QPushButton#{obj_name}:hover {{
                    background-color: {hover_bg};
                    border-color: {C.BORDER_HOVER};
                }}
            """)
