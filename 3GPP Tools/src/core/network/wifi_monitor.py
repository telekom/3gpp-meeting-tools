# --- File: src/core/network/wifi_monitor.py ---
import subprocess
import time
import logging
from PyQt5.QtCore import QThread, pyqtSignal
from core.network.network_state import NetworkState

class WifiMonitorThread(QThread):
    status_updated = pyqtSignal(str, bool, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        self.target_keyword = "3GPPWIFI"
        self.target_server = "10.10.10.10"
        self.CREATE_NO_WINDOW = 0x08000000

    def run(self):
        # Grab the singleton instance
        net_state = NetworkState.get_instance()

        while self.running:
            try:
                network_name = self._get_network_profile_name()
                is_3gpp = (self.target_keyword in network_name.upper())
                server_reachable = False

                if is_3gpp:
                    server_reachable = self._ping_server(self.target_server)

                # ---> NEW: Update the global state synchronously
                net_state.update_state(network_name, is_3gpp, server_reachable)

                self.status_updated.emit(network_name, is_3gpp, server_reachable)

            except Exception as e:
                logging.error(f"[WiFi Monitor] Loop error: {e}")

            time.sleep(10)  # Safe 10-second polling interval

    def _get_network_profile_name(self) -> str:
        try:
            output = subprocess.check_output(
                ['powershell', '-NoProfile', '-Command', '(Get-NetConnectionProfile).Name'],
                creationflags=self.CREATE_NO_WINDOW,
                text=True,
                timeout=5
            )
            lines = [line.strip() for line in output.split('\n') if line.strip()]
            if lines:
                return lines[0]
        except subprocess.CalledProcessError:
            pass
        except subprocess.TimeoutExpired:
            logging.warning("[WiFi Monitor] PowerShell command timed out.")
        except Exception as e:
            logging.error(f"[WiFi Monitor] Unexpected error getting network name: {e}")
        return ""

    def _ping_server(self, ip: str) -> bool:
        try:
            result = subprocess.run(
                ["ping", "-n", "1", "-w", "1000", ip],
                creationflags=self.CREATE_NO_WINDOW,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=3
            )
            return result.returncode == 0
        except Exception:
            return False

    def stop(self):
        self.running = False
        self.wait()