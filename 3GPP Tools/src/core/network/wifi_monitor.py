# --- File: src/core/network/wifi_monitor.py ---
"""
Network & 3GPP Wi-Fi Monitor Thread.
Monitors active Wi-Fi SSID, detects 3GPP meeting networks, and checks
local meeting server reachability without blocking other threads or the GUI.
"""

import logging
import socket
import subprocess
from typing import Optional, Tuple
from PyQt5.QtCore import QThread, pyqtSignal
from core.network.network_state import NetworkState


class WifiMonitorThread(QThread):
    # Emits: (network_name: str, is_3gpp: bool, server_reachable: bool)
    status_updated = pyqtSignal(str, bool, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self.target_keyword = "3GPPWIFI"
        self.target_server = "10.10.10.10"
        self.CREATE_NO_WINDOW = 0x08000000
        self._last_state: Tuple[Optional[str], Optional[bool], Optional[bool]] = (None, None, None)

    def run(self):
        logging.info("📶 [WiFi Monitor] Worker thread started running.")
        net_state = NetworkState.get_instance()

        while self._running:
            try:
                network_name = self._get_network_profile_name()

                # If netsh momentarily dropped, preserve the previous SSID to prevent flapping
                if not network_name and self._last_state[0]:
                    network_name = self._last_state[0]

                is_3gpp = bool(network_name and self.target_keyword in network_name.upper())
                server_reachable = False

                if is_3gpp and self._running:
                    server_reachable = self._check_server_reachable(self.target_server)

                if not self._running:
                    break

                # Update global network state singleton
                net_state.update_state(network_name, is_3gpp, server_reachable)

                current_state = (str(network_name), bool(is_3gpp), bool(server_reachable))

                # Log and emit ONLY when network status actually changes
                if current_state != self._last_state:
                    self._last_state = current_state
                    logging.info(
                        f"📶 [WiFi Monitor] Network changed: SSID='{network_name}', "
                        f"is_3gpp={is_3gpp}, server_reachable={server_reachable}"
                    )
                    self.status_updated.emit(str(network_name), bool(is_3gpp), bool(server_reachable))
                else:
                    logging.debug(f"📶 [WiFi Monitor] Heartbeat unchanged: SSID='{network_name}'")

            except Exception as e:
                logging.error(f"❌ [WiFi Monitor] Loop exception: {e}", exc_info=True)
                if self._running and self._last_state != ("", False, False):
                    self._last_state = ("", False, False)
                    self.status_updated.emit("", False, False)

            # Polling delay: checks exit flag every 100ms
            for _ in range(100):
                if not self._running:
                    break
                self.msleep(100)

        logging.info("📶 [WiFi Monitor] Worker thread cleanly exited run loop.")

    def _get_network_profile_name(self) -> str:
        """Queries Windows WLAN interface cleanly via netsh without heavy shell wrappers."""
        try:
            res = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                timeout=2.0,
                creationflags=self.CREATE_NO_WINDOW
            )
            if res.returncode == 0 and res.stdout:
                output = res.stdout.decode("cp1252", errors="replace")
                for line in output.splitlines():
                    if ":" in line:
                        parts = line.split(":", 1)
                        key = parts[0].strip()
                        val = parts[1].strip()
                        # Strict key match guarantees BSSID is never mistaken for SSID
                        if key == "SSID" and val:
                            return val
        except Exception:
            pass

        return ""

    def _check_server_reachable(self, ip: str) -> bool:
        """Fast TCP probe to port 80/443 with a 200ms timeout (no ping subprocesses)."""
        for port in (80, 443):
            try:
                with socket.create_connection((ip, port), timeout=0.2):
                    return True
            except Exception:
                continue
        return False

    def stop(self):
        """Cleanly signals thread to stop and joins."""
        self._running = False
        self.quit()
        if not self.wait(600):
            self.terminate()
            self.wait(200)