# --- File: src/core/network/wifi_monitor.py ---
import logging
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal
from core.network.network_state import NetworkState


class WifiMonitorThread(QThread):
    status_updated = pyqtSignal(str, bool, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self.target_keyword = "3GPPWIFI"
        self.target_server = "10.10.10.10"
        self.CREATE_NO_WINDOW = 0x08000000

    def run(self):
        logging.info("📶 [WiFi Monitor] Worker thread started running.")
        net_state = NetworkState.get_instance()

        while self._running:
            try:
                network_name = self._get_network_profile_name()
                is_3gpp = bool(network_name and self.target_keyword in network_name.upper())
                server_reachable = False

                if is_3gpp and self._running:
                    server_reachable = self._ping_server(self.target_server)

                if not self._running:
                    break

                net_state.update_state(network_name, is_3gpp, server_reachable)
                self.status_updated.emit(str(network_name), bool(is_3gpp), bool(server_reachable))

            except Exception as e:
                logging.error(f"❌ [WiFi Monitor] Loop exception: {e}", exc_info=True)
                if self._running:
                    self.status_updated.emit("", False, False)

            # Polling delay: checks exit flag every 100ms
            for _ in range(100):
                if not self._running:
                    break
                self.msleep(100)

        logging.info("📶 [WiFi Monitor] Worker thread cleanly exited run loop.")

    def _get_network_profile_name(self) -> str:
        """Attempts fast netsh query first, falls back to non-interactive PowerShell."""
        # Fast path: Native Windows netsh command
        try:
            output = subprocess.check_output(
                ["netsh", "wlan", "show", "interfaces"],
                creationflags=self.CREATE_NO_WINDOW,
                stdin=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                text=True,
                timeout=1.5
            )
            for line in output.splitlines():
                if "SSID" in line and "BSSID" not in line:
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        ssid = parts[1].strip()
                        if ssid:
                            return ssid
        except Exception:
            pass

        # Fallback: Non-interactive PowerShell with DEVNULL redirection
        try:
            output = subprocess.check_output(
                [
                    'powershell',
                    '-NoProfile',
                    '-NonInteractive',
                    '-ExecutionPolicy', 'Bypass',
                    '-Command', '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; (Get-NetConnectionProfile).Name'
                ],
                creationflags=self.CREATE_NO_WINDOW,
                stdin=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                text=True,
                timeout=2.0
            )
            lines = [line.strip() for line in output.splitlines() if line.strip()]
            return lines[0] if lines else ""
        except Exception as e:
            logging.debug(f"[WiFi Monitor] Fallback lookup failed: {e}")
            return ""

    def _ping_server(self, ip: str) -> bool:
        try:
            result = subprocess.run(
                ["ping", "-n", "1", "-w", "500", ip],
                creationflags=self.CREATE_NO_WINDOW,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=1.0
            )
            return result.returncode == 0
        except Exception:
            return False

    def stop(self):
        """Signals thread to stop and joins cleanly."""
        self._running = False
        self.quit()
        if not self.wait(1000):
            self.terminate()
            self.wait(300)