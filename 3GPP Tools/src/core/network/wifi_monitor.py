import subprocess
import time
import logging
from PyQt5.QtCore import QThread, pyqtSignal


class WifiMonitorThread(QThread):
    status_updated = pyqtSignal(str, bool, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        # ---> UPDATED: Exact 3GPP WiFi identifier <---
        self.target_keyword = "3GPPWIFI"
        self.target_server = "10.10.10.10"
        self.CREATE_NO_WINDOW = 0x08000000

    def run(self):
        while self.running:
            try:
                network_name = self._get_network_profile_name()

                # Fuzzy match: Check if "3GPPWIFI" is anywhere in the Windows profile name
                is_3gpp = (self.target_keyword in network_name.upper())
                server_reachable = False

                if is_3gpp:
                    server_reachable = self._ping_server(self.target_server)

                self.status_updated.emit(network_name, is_3gpp, server_reachable)

            except Exception as e:
                logging.error(f"[WiFi Monitor] Loop error: {e}")

            time.sleep(10)  # Safe 10-second polling interval

    def _get_network_profile_name(self) -> str:
        """Uses PowerShell to get the active network profile name without requiring Admin/Location rights."""
        try:
            output = subprocess.check_output(
                ['powershell', '-NoProfile', '-Command', '(Get-NetConnectionProfile).Name'],
                creationflags=self.CREATE_NO_WINDOW,
                text=True,
                timeout=5
            )

            # PowerShell might return multiple lines if connected to Ethernet AND WiFi.
            # We take the first valid non-empty line.
            lines = [line.strip() for line in output.split('\n') if line.strip()]
            if lines:
                return lines[0]

        except subprocess.CalledProcessError:
            # Command failed, likely no network connection active
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
        """Safely kills the loop before application shutdown."""
        self.running = False
        self.wait()