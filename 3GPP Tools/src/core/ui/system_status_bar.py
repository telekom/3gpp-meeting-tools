# --- File: src/core/ui/system_status_bar.py ---
import logging
import os
import urllib.request
from typing import Optional

from PyQt5.QtCore import Qt, QSize, pyqtSignal, QTimer
from PyQt5.QtWidgets import (
    QDialog,
    QLabel,
    QPushButton,
    QStatusBar,
    QWidget,
)

from core.ai.ollama_client import load_ollama_config, OllamaMonitorThread
from core.ai.ollama_dialog import OllamaConfigDialog
from core.network.session import NetworkSession, ProxyProfileManager
from core.network.wifi_monitor import WifiMonitorThread
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, T
from core.ui.ui_components import ProxyDialog

logger = logging.getLogger(__name__)


class SystemStatusBar(QStatusBar):
    """
    Manages application status, background connectivity monitors,
    and interactive status indicator pills (Ollama, Proxy, Network).
    """

    proxy_changed = pyqtSignal()
    reinit_requested = pyqtSignal()
    drawer_toggle_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.wifi_monitor: Optional[WifiMonitorThread] = None
        self.ollama_monitor: Optional[OllamaMonitorThread] = None
        self._drawer_open: bool = False
        self._drawer_queue_count: int = 0
        self._drawer_unread_errors: int = 0

        self.showMessage("⏳ Initializing...")
        self._setup_indicators()

    def _setup_indicators(self):
        saved_cfg = load_ollama_config()
        saved_model = saved_cfg.get("selected_model", "").strip()
        initial_ollama_text = f"🦙 {saved_model}" if saved_model else "🦙 Ollama"

        # 1. Ollama Indicator Pill
        self.ollama_indicator = QPushButton(initial_ollama_text)
        self.ollama_indicator.setCursor(Qt.PointingHandCursor)
        self.ollama_indicator.setToolTip("Click to configure Ollama LLM connection")
        self.ollama_indicator.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 1px 8px;
                font-size: 11px;
                font-weight: bold;
                color: #64748B;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
        """)
        self.ollama_indicator.clicked.connect(self.open_ollama_settings)
        self.addPermanentWidget(self.ollama_indicator)

        # 2. MCP AI Agent Indicator Pill
        self.mcp_indicator = QPushButton("🤖 MCP: Idle")
        self.mcp_indicator.setCursor(Qt.PointingHandCursor)
        self.mcp_indicator.setToolTip("Click to inspect MCP Server activity & external AI agents (Claude, Codex)")
        self.mcp_indicator.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 1px 8px;
                font-size: 11px;
                font-weight: bold;
                color: #64748B;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
        """)
        self.mcp_indicator.clicked.connect(self.open_mcp_monitor)
        self.addPermanentWidget(self.mcp_indicator)

        # 3. Proxy Indicator Pill
        self.proxy_indicator = QPushButton("🌐 Direct")
        self.proxy_indicator.setCursor(Qt.PointingHandCursor)
        self.proxy_indicator.clicked.connect(self.open_proxy_settings)
        self.addPermanentWidget(self.proxy_indicator)

        # 3. Network Indicator
        self.network_indicator = QLabel("📶 Network")
        self.network_indicator.setStyleSheet("color: gray; padding: 0 10px;")
        self.addPermanentWidget(self.network_indicator)

        # 4. Collapsible Bottom Drawer Toggle (Console & Queue + Badges)
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

    def update_drawer_badge(self, opened: bool, queue_count: int, unread_errors: int) -> None:
        """Update the status-bar Console & Queue pill text, active state, and error/queue badges."""
        self._drawer_open = bool(opened)
        self._drawer_queue_count = max(0, int(queue_count))
        self._drawer_unread_errors = max(0, int(unread_errors))

        parts = ["Console && Queue"]
        if self._drawer_queue_count > 0:
            parts.append(f"{self._drawer_queue_count} queued")
        if self._drawer_unread_errors > 0:
            noun = "err" if self._drawer_unread_errors == 1 else "errs"
            parts.append(f"{self._drawer_unread_errors} {noun}")

        self.drawer_toggle_btn.blockSignals(True)
        self.drawer_toggle_btn.setChecked(self._drawer_open)
        self.drawer_toggle_btn.blockSignals(False)
        self.drawer_toggle_btn.setText(" · ".join(parts))

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
                padding: 1px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#statusDrawerToggleBtn:hover {{
                background-color: {hover_bg};
                border-color: {C.BORDER_HOVER};
            }}
        """)

    # ==========================================
    # --- PROXY ORCHESTRATION ---
    # ==========================================
    def sync_proxy_environment_from_session(self):
        """Ensures urllib and os.environ match NetworkSession active proxies."""
        session = NetworkSession.get_instance()
        proxies = {k: v for k, v in (session.proxies or {}).items() if v}

        if proxies.get("http"):
            os.environ["HTTP_PROXY"] = proxies["http"]
        else:
            os.environ.pop("HTTP_PROXY", None)

        if proxies.get("https"):
            os.environ["HTTPS_PROXY"] = proxies["https"]
        else:
            os.environ.pop("HTTPS_PROXY", None)

        proxy_handler = urllib.request.ProxyHandler(proxies)
        opener = urllib.request.build_opener(proxy_handler)
        urllib.request.install_opener(opener)

    def update_proxy_indicator(self, console_panel=None):
        """Updates both the bottom status pill and the optional ConsolePanel proxy button."""
        session = NetworkSession.get_instance()
        active_proxies = {
            proto: NetworkSession._sanitize_proxy_url(url)
            for proto, url in (session.proxies or {}).items()
            if url
        }

        if active_proxies:
            details = "\n".join(f"• {k.upper()}: {v}" for k, v in active_proxies.items())
            tooltip = f"Proxy Active:\n{details}\n\nClick to configure proxy settings"

            self.proxy_indicator.setText("🛡️ Proxy: ON")
            self.proxy_indicator.setToolTip(tooltip)
            self.proxy_indicator.setStyleSheet("""
                QPushButton {
                    background-color: #EBF8FF;
                    border: 1px solid #90CDF4;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #1E5C99;
                }
                QPushButton:hover {
                    background-color: #BEE3F8;
                    border-color: #63B3ED;
                }
            """)
            if console_panel:
                console_panel.set_proxy_status(True, tooltip)
        else:
            profile = ProxyProfileManager.load_profile()
            has_saved = bool(profile.get("http_host") or profile.get("https_host"))
            state_note = " (Saved profile inactive)" if has_saved else ""
            tooltip = f"Direct Connection — No proxy active{state_note}\nClick to configure proxy settings"

            self.proxy_indicator.setText("🌐 Direct")
            self.proxy_indicator.setToolTip(tooltip)
            self.proxy_indicator.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #64748B;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    border-color: #94A3B8;
                }
            """)
            if console_panel:
                console_panel.set_proxy_status(False, tooltip)

    def open_proxy_settings(self):
        """Presents proxy dialog and synchronizes active proxy state."""
        proxy_dialog = ProxyDialog(self.window())
        if proxy_dialog.exec_() == QDialog.Accepted:
            self.sync_proxy_environment_from_session()
            console_panel = getattr(self.window(), "console_panel", None)
            self.update_proxy_indicator(console_panel)
            self.proxy_changed.emit()
            self.reinit_requested.emit()

    # ==========================================
    # --- BACKGROUND SERVICE MONITORS ---
    # ==========================================
    def start_wifi_monitor(self):
        """Starts Wi-Fi monitor independently."""
        try:
            if self.wifi_monitor is None or not self.wifi_monitor.isRunning():
                logger.info("🏁 [STARTUP:SERVICES] Starting WifiMonitorThread...")
                self.wifi_monitor = WifiMonitorThread(parent=None)
                self.wifi_monitor.status_updated.connect(self._update_network_indicator)
                self.wifi_monitor.start()
        except Exception as e:
            logger.error(f"❌ Failed to start WifiMonitorThread: {e}")
            self.network_indicator.setText("🌐 Offline")

    def start_ollama_monitor(self):
        """Starts Ollama monitor independently."""
        try:
            if self.ollama_monitor is None or not self.ollama_monitor.isRunning():
                logger.info("🏁 [STARTUP:SERVICES] Starting OllamaMonitorThread...")
                self.ollama_monitor = OllamaMonitorThread(parent=None)
                self.ollama_monitor.status_updated.connect(self._update_ollama_indicator)
                self.ollama_monitor.start()
        except Exception as e:
            logger.error(f"❌ Failed to start OllamaMonitorThread: {e}")
            self.ollama_indicator.setText("🦙 Offline")

    def _update_network_indicator(self, ssid: str, is_3gpp: bool, server_reachable: bool):
        if is_3gpp and server_reachable:
            self.network_indicator.setText(f"🟢 {ssid} (Local Server Active)")
            self.network_indicator.setStyleSheet("color: #2e7d32; font-weight: bold; padding: 0 10px;")
        elif is_3gpp and not server_reachable:
            self.network_indicator.setText(f"🟡 {ssid} (No Local Server)")
            self.network_indicator.setStyleSheet("color: #b8860b; font-weight: bold; padding: 0 10px;")
        else:
            display_name = ssid if ssid else "Offline"
            self.network_indicator.setText(f"🌐 {display_name}")
            self.network_indicator.setStyleSheet("color: gray; padding: 0 10px;")

    def _update_ollama_indicator(self, is_online: bool, selected_model: str, models: list, error: str):
        if is_online:
            model_display = selected_model if selected_model else (models[0] if models else "Connected")
            self.ollama_indicator.setText(f"🦙 {model_display}")
            self.ollama_indicator.setToolTip(f"Ollama Online ({len(models)} models available)\nClick to configure")
            self.ollama_indicator.setStyleSheet("""
                QPushButton {
                    background-color: #F0FDF4;
                    border: 1px solid #BBF7D0;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #15803D;
                }
                QPushButton:hover {
                    background-color: #DCFCE7;
                    border-color: #86EFAC;
                }
            """)
        else:
            self.ollama_indicator.setText("🦙 Offline")
            self.ollama_indicator.setToolTip(f"Ollama Unreachable: {error}\nClick to configure")
            self.ollama_indicator.setStyleSheet("""
                QPushButton {
                    background-color: #FEF2F2;
                    border: 1px solid #FECACA;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #DC2626;
                }
                QPushButton:hover {
                    background-color: #FEE2E2;
                    border-color: #FCA5A5;
                }
            """)

    def open_ollama_settings(self):
        """Opens Ollama connection, model preferences, and service management dialog."""
        dialog = OllamaConfigDialog(self.window())
        dialog.exec_()

        if self.ollama_monitor and self.ollama_monitor.isRunning():
            cfg = load_ollama_config()
            self.ollama_monitor.client.reconfigure(cfg.get("host"), cfg.get("proxy_mode"))
            self.ollama_monitor.trigger_check()

    def open_mcp_monitor(self):
        """Opens the MCP Server & External AI Agent Cockpit dialog."""
        try:
            from core.ui.mcp_monitor_dialog import McpMonitorDialog
            if self._mcp_dialog is None:
                self._mcp_dialog = McpMonitorDialog(self.window())
            self._mcp_dialog.show()
            self._mcp_dialog.raise_()
            self._mcp_dialog.activateWindow()
            self._mcp_dialog.refresh_all()
        except Exception as exc:
            logger.exception("Failed to open MCP Monitor dialog: %s", exc)

    def update_mcp_indicator(self):
        """Polls active MCP sessions and recent activity stats to update the status bar pill."""
        try:
            from core.ai.mcp_activity import get_active_mcp_sessions, get_mcp_activity_stats
            sessions = get_active_mcp_sessions()
            stats = get_mcp_activity_stats()
            total_calls = stats.get("total_calls", 0)

            if sessions:
                clients = ", ".join(s.get("client", "agent") for s in sessions)
                pids = ", ".join(str(s.get("pid", "?")) for s in sessions)
                text = f"🤖 MCP: 🟢 Active ({total_calls})"
                tooltip = (
                    f"MCP Server: Active\n"
                    f"Connected Client: {clients} (PID: {pids})\n"
                    f"Total Calls: {total_calls}\n"
                    f"Success Rate: {stats.get('success_rate', 100)}%\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self.mcp_indicator.setText(text)
                self.mcp_indicator.setToolTip(tooltip)
                self.mcp_indicator.setStyleSheet("""
                    QPushButton {
                        background-color: #F0FDF4;
                        border: 1px solid #86EFAC;
                        border-radius: 4px;
                        padding: 1px 8px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #15803D;
                    }
                    QPushButton:hover {
                        background-color: #DCFCE7;
                        border-color: #4ADE80;
                    }
                """)
            else:
                text = f"🤖 MCP: {total_calls} calls" if total_calls > 0 else "🤖 MCP: Idle"
                tooltip = (
                    f"MCP Server: Idle (No active agent process connected)\n"
                    f"Total Calls Logged: {total_calls}\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self.mcp_indicator.setText(text)
                self.mcp_indicator.setToolTip(tooltip)
                self.mcp_indicator.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        border: 1px solid #CBD5E1;
                        border-radius: 4px;
                        padding: 1px 8px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #64748B;
                    }
                    QPushButton:hover {
                        background-color: #F1F5F9;
                        border-color: #94A3B8;
                    }
                """)
        except Exception:
            pass

    def stop_monitors(self):
        """Cleanly stops background thread monitors upon window shutdown."""
        if hasattr(self, "mcp_timer") and self.mcp_timer.isActive():
            self.mcp_timer.stop()
        if self.wifi_monitor and self.wifi_monitor.isRunning():
            self.wifi_monitor.requestInterruption()
            self.wifi_monitor.wait(300)
        if self.ollama_monitor and self.ollama_monitor.isRunning():
            self.ollama_monitor.requestInterruption()
            self.ollama_monitor.wait(300)
