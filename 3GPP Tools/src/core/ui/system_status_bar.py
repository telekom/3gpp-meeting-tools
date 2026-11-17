# --- File: src/core/ui/system_status_bar.py ---
import logging
import os
import urllib.request
from typing import Optional

from PyQt5.QtCore import Qt, QSize, pyqtSignal, QTimer
from PyQt5.QtWidgets import (
    QDialog,
    QLabel,
    QMenu,
    QPushButton,
    QStatusBar,
    QWidget,
)

from core.ai.ollama_client import load_ollama_config, OllamaMonitorThread
from core.ai.ollama_dialog import OllamaConfigDialog
from core.network.session import NetworkSession, ProxyProfileManager
from core.network.wifi_monitor import WifiMonitorThread
from core.ui.font_scale import get_font_scale_notifier
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, T
from core.ui.ui_components import ProxyDialog
from core.ui.world_clock_service import WebTimezoneSyncWorker, WorldClockService

logger = logging.getLogger(__name__)

_MINIMAL_HYSTERESIS_PX = 24


def _status_pill_qss(
    bg: str = "transparent",
    fg: str = C.TEXT_MUTED,
    bd: str = C.BORDER_STRONG,
    hover_bg: str = C.BG_CANVAS,
    hover_bd: str = C.BORDER_HOVER,
    compact: bool = False,
) -> str:
    """Build consistent DPI-friendly status bar pill QSS."""
    pad = "2px 5px" if compact else "2px 8px"
    return (
        f"QPushButton {{"
        f" background-color: {bg}; border: 1px solid {bd}; border-radius: {R.SM}px;"
        f" padding: {pad}; font-family: {T.FONT_FAMILY_UI}; font-size: {T.SIZE_CAPTION}px;"
        f" font-weight: 600; color: {fg};"
        f"}}"
        f"QPushButton:hover {{ background-color: {hover_bg}; border-color: {hover_bd}; }}"
        f"QPushButton::menu-indicator {{ image: none; width: 0px; }}"
    )


class SystemStatusBar(QStatusBar):
    """
    Manages application status, background connectivity monitors,
    System tools menu, Console & Queue drawer badge toggle,
    and interactive status indicator pills (Ollama, MCP, Proxy, Network).
    """

    proxy_changed = pyqtSignal()
    reinit_requested = pyqtSignal()
    drawer_toggle_requested = pyqtSignal()
    assistant_dock_requested = pyqtSignal()
    search_dock_requested = pyqtSignal()
    task_manager_requested = pyqtSignal()
    db_maintenance_requested = pyqtSignal()
    network_config_requested = pyqtSignal()
    java_info_requested = pyqtSignal()
    update_check_requested = pyqtSignal()
    restart_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.wifi_monitor: Optional[WifiMonitorThread] = None
        self.ollama_monitor: Optional[OllamaMonitorThread] = None
        self._minimal_mode: bool = False
        self._responsive_armed: bool = False
        self._in_responsive_check: bool = False
        self._last_available_width: Optional[int] = None
        self._full_pill_widths: dict = {}
        self._drawer_open: bool = False
        self._drawer_queue_count: int = 0
        self._drawer_unread_errors: int = 0
        self._right_panel_open: bool = False
        self._right_panel_mode: str = "assistant"
        self._update_state: tuple = (False, 0, "", "")
        self._ollama_initial_model: str = ""
        self._ollama_state: Optional[tuple] = None
        self._mcp_state: tuple = (False, 0)
        self._proxy_active: bool = False
        self._network_state: tuple = ("", False, False)
        self._last_full_clock_text: str = "🕒 --:--"
        self._world_clock_service = WorldClockService()
        self._clock_sync_workers: list = []
        self._clock_auto_synced: set = set()
        self._monitors_stopped: bool = False

        self.showMessage("⏳ Initializing...")
        self._setup_indicators()
        self.messageChanged.connect(self._on_status_message_changed)
        get_font_scale_notifier().scale_changed.connect(
            lambda _scale: self._recheck_responsive_fit()
        )

    @property
    def is_minimal_mode(self) -> bool:
        """Whether the status bar is currently in narrow-window minimalist icon mode."""
        return self._minimal_mode

    def _full_update_text(self) -> str:
        visible, count, _, _ = self._update_state
        if not visible:
            return ""
        return f"Update ({count})" if count > 0 else "Update Available"

    def _full_ollama_text(self) -> str:
        if self._ollama_state is None:
            return f"🦙 {self._ollama_initial_model}" if self._ollama_initial_model else "🦙 Ollama"
        is_online, model_display, _, _ = self._ollama_state
        return f"🦙 {model_display}" if is_online else "🦙 Offline"

    def _full_mcp_text(self) -> str:
        active, total_calls = self._mcp_state
        if active:
            return f"🤖 MCP: 🟢 Active ({total_calls})"
        return f"🤖 MCP: {total_calls} calls" if total_calls > 0 else "🤖 MCP: Idle"

    def _full_proxy_text(self) -> str:
        return "🛡️ Proxy: ON" if self._proxy_active else "🌐 Direct"

    def _full_network_text(self) -> str:
        ssid, is_3gpp, server_reachable = self._network_state
        if is_3gpp and server_reachable:
            return f"🟢 {ssid} (Local Server Active)"
        if is_3gpp and not server_reachable:
            return f"🟡 {ssid} (No Local Server)"
        display_name = ssid if ssid else ("Network" if self.wifi_monitor is None else "Offline")
        return "📶 Network" if display_name == "Network" else f"🌐 {display_name}"

    def _full_drawer_text(self) -> str:
        parts = ["Console && Queue"]
        if self._drawer_queue_count > 0:
            parts.append(f"{self._drawer_queue_count} queued")
        if self._drawer_unread_errors > 0:
            noun = "err" if self._drawer_unread_errors == 1 else "errs"
            parts.append(f"{self._drawer_unread_errors} {noun}")
        return " · ".join(parts)

    def _estimate_full_pill_width(
        self, widget: QWidget, full_text: str, is_label: bool = False, has_icon: bool = False
    ) -> int:
        """Return exact/estimated pixel width of a pill when rendered in full mode."""
        if not self._minimal_mode:
            w = widget.sizeHint().width()
            self._full_pill_widths[widget] = (full_text, w)
            return w
        cached = self._full_pill_widths.get(widget)
        if cached is not None and cached[0] == full_text:
            return cached[1]
        display_text = full_text.replace("&&", "&")
        fm = widget.fontMetrics()
        text_w = fm.horizontalAdvance(display_text) if display_text else 0
        if is_label:
            return text_w + 16
        icon_w = 18 if has_icon else 0
        return text_w + icon_w + 18

    def compute_full_required_width(self) -> int:
        """
        Compute the total width required to display the left status message
        plus all active status bar pills in full (non-minimalist) mode.
        """
        entries = []
        if self._update_state[0]:
            entries.append((self.update_indicator, self._full_update_text(), False, True))
        entries.extend([
            (self.ollama_indicator, self._full_ollama_text(), False, False),
            (self.mcp_indicator, self._full_mcp_text(), False, False),
            (self.proxy_indicator, self._full_proxy_text(), False, False),
            (self.network_indicator, self._full_network_text(), True, False),
            (self.system_btn, "System", False, True),
            (self.assistant_dock_btn, "Assistant", False, True),
            (self.search_dock_btn, "Quick Search", False, True),
            (self.drawer_toggle_btn, self._full_drawer_text(), False, True),
            (self.clock_indicator, self._last_full_clock_text, False, False),
        ])

        pills_total = sum(
            self._estimate_full_pill_width(w, txt, is_label=is_lbl, has_icon=has_ico)
            for w, txt, is_lbl, has_ico in entries
        )
        spacing_total = len(entries) * 6

        msg = self.currentMessage() or "🟢 System Idle."
        msg_w = self.fontMetrics().horizontalAdvance(msg)
        left_reserve = max(125, msg_w + 28)

        return left_reserve + pills_total + spacing_total + 24

    def _on_status_message_changed(self, _message: str) -> None:
        self._recheck_responsive_fit()

    def _recheck_responsive_fit(self) -> None:
        """Re-evaluate minimalist mode after pill content or left message changes."""
        if not self._responsive_armed or self._in_responsive_check:
            return
        self.apply_responsive_width(self._last_available_width)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        w = event.size().width()
        if w > 0:
            self._responsive_armed = True
            self.apply_responsive_width(w)

    def apply_responsive_width(self, available_width: Optional[int] = None) -> None:
        """Dynamically toggle minimalist mode based on actual full-mode pill width + hysteresis."""
        if self._in_responsive_check:
            return
        if available_width is not None:
            self._responsive_armed = True
            self._last_available_width = int(available_width)
            avail = int(available_width)
        elif self.width() > 200 and self.isVisible():
            avail = self.width()
            self._last_available_width = avail
        elif self._last_available_width is not None:
            avail = self._last_available_width
        else:
            win = self.window()
            avail = win.width() if win is not None else self.width()

        if avail <= 0:
            return

        self._in_responsive_check = True
        try:
            req = self.compute_full_required_width()
            if not self._minimal_mode and avail < req:
                self.set_minimal_mode(True)
            elif self._minimal_mode and avail >= req + _MINIMAL_HYSTERESIS_PX:
                self.set_minimal_mode(False)
        finally:
            self._in_responsive_check = False

    def set_minimal_mode(self, minimal: bool) -> None:
        """Switch between full pill labels and compact icon-only minimalist mode."""
        want = bool(minimal)
        if want == self._minimal_mode:
            return
        self._minimal_mode = want
        self._refresh_all_indicators()

    def _refresh_all_indicators(self) -> None:
        """Re-render every status bar pill for the current minimalist mode state."""
        self._render_update_indicator()
        self._render_ollama_indicator()
        self._render_mcp_indicator()
        self._render_proxy_indicator()
        self._render_network_indicator()
        self._render_system_btn()
        self.update_right_panel_badges(self._right_panel_open, self._right_panel_mode)
        self.update_drawer_badge(
            self._drawer_open,
            self._drawer_queue_count,
            self._drawer_unread_errors,
        )
        self.update_clock_indicator()

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

    def _render_proxy_indicator(self) -> None:
        if self._proxy_active:
            self.proxy_indicator.setText("🛡️" if self._minimal_mode else "🛡️ Proxy: ON")
            self.proxy_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.INFO_SOFT_ALT,
                    fg=C.ACCENT_PRIMARY,
                    bd=C.INFO_BORDER_SKY,
                    hover_bg=C.INFO_BORDER_ALT,
                    hover_bd=C.ACCENT_SKY,
                    compact=self._minimal_mode,
                )
            )
        else:
            self.proxy_indicator.setText("🌐" if self._minimal_mode else "🌐 Direct")
            self.proxy_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))

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
            self._proxy_active = True
            self.proxy_indicator.setToolTip(tooltip)
            self._render_proxy_indicator()
            if console_panel:
                console_panel.set_proxy_status(True, tooltip)
        else:
            profile = ProxyProfileManager.load_profile()
            has_saved = bool(profile.get("http_host") or profile.get("https_host"))
            state_note = " (Saved profile inactive)" if has_saved else ""
            tooltip = f"Direct Connection — No proxy active{state_note}\nClick to configure proxy settings"
            self._proxy_active = False
            self.proxy_indicator.setToolTip(tooltip)
            self._render_proxy_indicator()
            if console_panel:
                console_panel.set_proxy_status(False, tooltip)
        self._recheck_responsive_fit()

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
        self._monitors_stopped = False
        QTimer.singleShot(2500, self._schedule_active_venue_web_sync)
        try:
            if self.wifi_monitor is None or not self.wifi_monitor.isRunning():
                logger.info("🏁 [STARTUP:SERVICES] Starting WifiMonitorThread...")
                self.wifi_monitor = WifiMonitorThread(parent=None)
                self.wifi_monitor.status_updated.connect(self._update_network_indicator)
                self.wifi_monitor.start()
        except Exception as e:
            logger.error(f"❌ Failed to start WifiMonitorThread: {e}")
            self._network_state = ("Offline", False, False)
            self._render_network_indicator()

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
            self._ollama_state = (False, "", 0, str(e))
            self._render_ollama_indicator()

    def _render_network_indicator(self) -> None:
        ssid, is_3gpp, server_reachable = self._network_state
        if self._minimal_mode:
            if not is_3gpp:
                self.network_indicator.setVisible(False)
                return
            self.network_indicator.setVisible(True)
            if server_reachable:
                self.network_indicator.setText("🟢 3GPP")
                self.network_indicator.setToolTip(f"{ssid} (Local Server Active)")
                self.network_indicator.setStyleSheet(
                    f"color: {C.SUCCESS}; font-size: {T.SIZE_CAPTION}px; font-weight: 600; padding: 0 4px;"
                )
            else:
                self.network_indicator.setText("🟡 3GPP")
                self.network_indicator.setToolTip(f"{ssid} (No Local Server)")
                self.network_indicator.setStyleSheet(
                    f"color: {C.WARNING}; font-size: {T.SIZE_CAPTION}px; font-weight: 600; padding: 0 4px;"
                )
            return

        self.network_indicator.setVisible(True)
        if is_3gpp and server_reachable:
            self.network_indicator.setText(f"🟢 {ssid} (Local Server Active)")
            self.network_indicator.setToolTip(f"Connected to 3GPP meeting Wi-Fi: {ssid} (Local Server Active)")
            self.network_indicator.setStyleSheet(
                f"color: {C.SUCCESS}; font-size: {T.SIZE_CAPTION}px; font-weight: 600; padding: 0 8px;"
            )
        elif is_3gpp and not server_reachable:
            self.network_indicator.setText(f"🟡 {ssid} (No Local Server)")
            self.network_indicator.setToolTip(f"Connected to 3GPP meeting Wi-Fi: {ssid} (No Local Server)")
            self.network_indicator.setStyleSheet(
                f"color: {C.WARNING}; font-size: {T.SIZE_CAPTION}px; font-weight: 600; padding: 0 8px;"
            )
        else:
            display_name = ssid if ssid else ("Network" if ssid == "" and not self._monitors_stopped else "Offline")
            if ssid == "" and not self._network_state[0]:
                display_name = "Network" if self.wifi_monitor is None else "Offline"
            self.network_indicator.setText(f"🌐 {display_name}" if display_name != "Network" else "📶 Network")
            self.network_indicator.setToolTip(f"Network: {display_name}")
            self.network_indicator.setStyleSheet(
                f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; padding: 0 8px;"
            )

    def _update_network_indicator(self, ssid: str, is_3gpp: bool, server_reachable: bool):
        clean_ssid = str(ssid or "").strip() or "Offline"
        self._network_state = (clean_ssid, bool(is_3gpp), bool(server_reachable))
        self._render_network_indicator()
        self._recheck_responsive_fit()

    def _render_ollama_indicator(self) -> None:
        if self._ollama_state is None:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))
            return

        is_online, model_display, models_count, error = self._ollama_state
        if is_online:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setToolTip(
                f"Ollama Online ({model_display}, {models_count} models available)\nClick to configure"
            )
            self.ollama_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.SUCCESS_SOFT,
                    fg=C.SUCCESS,
                    bd=C.SUCCESS_BORDER,
                    hover_bg=C.SUCCESS_SOFT_ALT,
                    hover_bd=C.SUCCESS_BORDER_MINT,
                    compact=self._minimal_mode,
                )
            )
        else:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setToolTip(f"Ollama Unreachable: {error}\nClick to configure")
            self.ollama_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.DANGER_SOFT,
                    fg=C.DANGER_ALT,
                    bd=C.DANGER_BORDER,
                    hover_bg=C.DANGER_SOFT_HOVER,
                    hover_bd=C.DANGER_BORDER_ALT,
                    compact=self._minimal_mode,
                )
            )

    def _update_ollama_indicator(self, is_online: bool, selected_model: str, models: list, error: str):
        model_display = selected_model if selected_model else (models[0] if models else "Connected")
        self._ollama_state = (bool(is_online), str(model_display), len(models), str(error or ""))
        self._render_ollama_indicator()
        self._recheck_responsive_fit()

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

    def _render_mcp_indicator(self) -> None:
        active, total_calls = self._mcp_state
        if active:
            text = "🤖 🟢" if self._minimal_mode else f"🤖 MCP: 🟢 Active ({total_calls})"
            self.mcp_indicator.setText(text)
            self.mcp_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.SUCCESS_SOFT,
                    fg=C.SUCCESS,
                    bd=C.SUCCESS_BORDER,
                    hover_bg=C.SUCCESS_SOFT_ALT,
                    hover_bd=C.SUCCESS_BORDER_MINT,
                    compact=self._minimal_mode,
                )
            )
        else:
            if self._minimal_mode:
                text = "🤖"
            else:
                text = f"🤖 MCP: {total_calls} calls" if total_calls > 0 else "🤖 MCP: Idle"
            self.mcp_indicator.setText(text)
            self.mcp_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))

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
                tooltip = (
                    f"MCP Server: Active\n"
                    f"Connected Client: {clients} (PID: {pids})\n"
                    f"Total Calls: {total_calls}\n"
                    f"Success Rate: {stats.get('success_rate', 100)}%\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self._mcp_state = (True, total_calls)
                self.mcp_indicator.setToolTip(tooltip)
                self._render_mcp_indicator()
            else:
                tooltip = (
                    f"MCP Server: Idle (No active agent process connected)\n"
                    f"Total Calls Logged: {total_calls}\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self._mcp_state = (False, total_calls)
                self.mcp_indicator.setToolTip(tooltip)
                self._render_mcp_indicator()
            self._recheck_responsive_fit()
        except Exception:
            pass

    def update_clock_indicator(self) -> None:
        """Refreshes the status-bar World Clock pill text and rich tooltip."""
        try:
            snapshot = self._world_clock_service.get_snapshot()
            self._last_full_clock_text = self._world_clock_service.format_status_bar_text(
                snapshot, compact=False
            )
            self.clock_indicator.setText(
                self._world_clock_service.format_status_bar_text(
                    snapshot, compact=self._minimal_mode
                )
            )
            self.clock_indicator.setToolTip(self._world_clock_service.format_status_bar_tooltip(snapshot))
            self.clock_indicator.setStyleSheet(
                _status_pill_qss(bg=C.BG_SURFACE, fg=C.TEXT_BODY, compact=self._minimal_mode)
            )
            self._recheck_responsive_fit()
        except Exception as e:
            logger.debug(f"[StatusBarClock] Failed to update clock indicator: {e}")

    def refresh_world_clock(self) -> None:
        """Invalidates meeting DB cache and immediately refreshes the clock pill."""
        self._world_clock_service.invalidate_cache()
        self.update_clock_indicator()

    def open_world_clock(self) -> None:
        """Opens the interactive World Clock & Active 3GPP Venue Timezone dialog."""
        from core.ui.world_clock_dialog import WorldClockDialog

        dlg = WorldClockDialog(self._world_clock_service, parent=self.window())
        dlg.settings_changed.connect(self.refresh_world_clock)
        dlg.exec_()
        self.refresh_world_clock()

    def _schedule_active_venue_web_sync(self) -> None:
        """Background-verifies active meeting venue timezones not yet web-verified or manually overridden."""
        if self._monitors_stopped:
            return
        try:
            snapshot = self._world_clock_service.get_snapshot()
            for v in snapshot.primary_venues + snapshot.other_venues:
                loc_key = v.location.strip().lower()
                if (
                    v.source in ("builtin", "unresolved")
                    and loc_key
                    and loc_key not in self._clock_auto_synced
                ):
                    self._clock_auto_synced.add(loc_key)
                    worker = WebTimezoneSyncWorker(
                        mode="city",
                        location=v.location,
                        meetings=v.meetings,
                        parent=self,
                    )
                    worker.city_resolved.connect(self._on_auto_venue_web_resolved)
                    worker.finished.connect(
                        lambda w=worker: self._clock_sync_workers.remove(w)
                        if w in self._clock_sync_workers
                        else None
                    )
                    self._clock_sync_workers.append(worker)
                    worker.start()
        except Exception as e:
            logger.debug(f"[StatusBarClock] Auto web sync scheduling failed: {e}")

    def _on_auto_venue_web_resolved(
        self, location: str, offset_min: int, label: str, iana_tz: str, meetings: list
    ) -> None:
        self._world_clock_service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="web_api",
            iana_tz=iana_tz,
            matching_meetings=meetings,
        )
        self.refresh_world_clock()

    def stop_monitors(self):
        """Cleanly stops background thread monitors upon window shutdown."""
        self._monitors_stopped = True
        if hasattr(self, "clock_timer") and self.clock_timer.isActive():
            self.clock_timer.stop()
        if hasattr(self, "mcp_timer") and self.mcp_timer.isActive():
            self.mcp_timer.stop()
        for w in list(getattr(self, "_clock_sync_workers", [])):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        if hasattr(self, "_clock_sync_workers"):
            self._clock_sync_workers.clear()
        if self.wifi_monitor and self.wifi_monitor.isRunning():
            self.wifi_monitor.requestInterruption()
            self.wifi_monitor.wait(300)
        if self.ollama_monitor and self.ollama_monitor.isRunning():
            self.ollama_monitor.requestInterruption()
            self.ollama_monitor.wait(300)
