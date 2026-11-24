# --- File: src/core/ui/status_bar/mixins/network_and_proxy.py ---
"""Proxy environment synchronization and Wi-Fi / 3GPP LAN status pill updates."""

import logging
import os
import urllib.request

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QDialog

from core.network.session import NetworkSession, ProxyProfileManager
from core.network.wifi_monitor import WifiMonitorThread
from core.ui.theme import C, T
from core.ui.ui_components import ProxyDialog

from ..styles import _status_pill_qss

logger = logging.getLogger(__name__)


class StatusBarNetworkProxyMixin:
    """Mixin managing proxy settings sync and Wi-Fi / 10.10.10.10 venue network status."""

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
