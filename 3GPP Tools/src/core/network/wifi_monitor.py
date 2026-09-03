import subprocess
import logging
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
        net_state = NetworkState.get_instance()

        while self._running:
            try:
                network_name = self._get_network_profile_name()
                is_3gpp = (self.target_keyword in network_name.upper())
                server_reachable = False

                if is_3gpp and self._running:
                    server_reachable = self._ping_server(self.target_server)

                if not self._running:
                    break

                net_state.update_state(network_name, is_3gpp, server_reachable)
                self.status_updated.emit(network_name, is_3gpp, server_reachable)

            except Exception as e:
                logging.error(f"[WiFi Monitor] Loop error: {e}")

            # Non-blocking 10-second polling interval (checks exit condition every 100ms)
            for _ in range(100):
                if not self._running:
                    break
                self.msleep(100)

    def _get_network_profile_name(self) -> str:
        try:
            output = subprocess.check_output(
                ['powershell', '-NoProfile', '-Command', '(Get-NetConnectionProfile).Name'],
                creationflags=self.CREATE_NO_WINDOW,
                text=True,
                timeout=3
            )
            lines = [line.strip() for line in output.splitlines() if line.strip()]
            return lines[0] if lines else ""
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            return ""
        except Exception as e:
            logging.debug(f"[WiFi Monitor] Error getting network name: {e}")
            return ""

    def _ping_server(self, ip: str) -> bool:
        try:
            result = subprocess.run(
                ["ping", "-n", "1", "-w", "800", ip],
                creationflags=self.CREATE_NO_WINDOW,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=2
            )
            return result.returncode == 0
        except Exception:
            return False

    def stop(self):
        """Immediately interrupts the loop and cleanly joins the thread."""
        self._running = False
        self.quit()
        if not self.wait(1500):
            self.terminate()
            self.wait(500)