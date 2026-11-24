# --- File: src/core/ui/status_bar/bar.py ---
"""Composed SystemStatusBar widget combining UI setup, responsive sizing, network/proxy, and AI/clock monitors."""

from typing import Optional

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QStatusBar, QWidget

from core.ai.ollama_client import OllamaMonitorThread
from core.network.wifi_monitor import WifiMonitorThread
from core.ui.font_scale import get_font_scale_notifier
from core.ui.world_clock_service import WorldClockService

from .mixins import (
    StatusBarAiClockMixin,
    StatusBarNetworkProxyMixin,
    StatusBarResponsiveMixin,
    StatusBarUiSetupMixin,
)


class SystemStatusBar(
    StatusBarResponsiveMixin,
    StatusBarUiSetupMixin,
    StatusBarNetworkProxyMixin,
    StatusBarAiClockMixin,
    QStatusBar,
):
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
