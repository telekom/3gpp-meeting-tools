import socket
import subprocess
import threading
import time
import logging


def detect_active_network_profile_name(target_keyword: str = "3GPPWIFI") -> str:
    """
    Fast headless network profile / SSID detection on Windows without requiring PyQt.
    Uses Windows Network List Manager (NLM) COM API (~2ms) with netsh fallback.
    """
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        nlm = None
        connected_networks = None
        network = None
        try:
            nlm = win32com.client.Dispatch("{DCB00C01-570F-4A9B-8D69-199FDBA5723B}")
            connected_networks = nlm.GetNetworks(1)
            names = []
            for network in connected_networks:
                try:
                    net_name = network.GetName()
                    if net_name:
                        names.append(net_name.strip())
                except Exception:
                    continue
            if names:
                for name in names:
                    if target_keyword.upper() in name.upper():
                        return name
                return names[0]
        finally:
            network = None
            connected_networks = None
            nlm = None
            pythoncom.CoUninitialize()
    except Exception as e:
        logging.debug(f"[NetworkState] NLM COM query exception: {e}")

    try:
        res = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            timeout=2.0,
            creationflags=0x08000000,
        )
        if res.returncode == 0 and res.stdout:
            output = res.stdout.decode("cp1252", errors="replace")
            for line in output.splitlines():
                if ":" in line:
                    key, val = line.split(":", 1)
                    if key.strip() == "SSID" and val.strip():
                        return val.strip()
    except Exception:
        pass

    return ""


class NetworkState:
    """
    A thread-safe Singleton that holds the current status of the network connection.
    Guarantees that background probes cannot overlap and supports headless CLI/MCP auto-probing.
    """
    _instance = None
    _lock = threading.RLock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(NetworkState, cls).__new__(cls)
                cls._instance.network_name = ""
                cls._instance.is_3gpp_wifi = False
                cls._instance.is_local_reachable = False
                cls._instance._last_probe_time = 0.0
                cls._instance._is_probing = False
        return cls._instance

    @classmethod
    def get_instance(cls):
        return cls()

    def update_state(self, name: str, is_3gpp: bool, reachable: bool):
        with self._lock:
            self.network_name = name
            self.is_3gpp_wifi = is_3gpp
            self.is_local_reachable = reachable
            self._last_probe_time = time.time()

    def ensure_probed(self, force: bool = False, timeout: float = 0.25) -> dict:
        """
        Synchronously probes active SSID and 10.10.10.10 reachability if never probed
        (e.g. in headless CLI / MCP mode where WifiMonitorThread is not running) or if forced.
        """
        with self._lock:
            if not force and self._last_probe_time > 0.0 and (time.time() - self._last_probe_time < 15.0):
                return self.get_status_dict()

        name = detect_active_network_profile_name("3GPPWIFI")
        is_3gpp = bool(name and "3GPPWIFI" in name.upper())
        reachable = self.check_local_reachability_fast("10.10.10.10", port=80, timeout=timeout)
        with self._lock:
            if name or not self.network_name:
                self.network_name = name
            self.is_3gpp_wifi = is_3gpp or reachable
            self.is_local_reachable = reachable
            self._last_probe_time = time.time()
            return self.get_status_dict()

    def get_status_dict(self) -> dict:
        with self._lock:
            return {
                "network_name": self.network_name,
                "is_3gpp_wifi": self.is_3gpp_wifi,
                "is_local_reachable": self.is_local_reachable,
                "local_server": "http://10.10.10.10/ftp",
                "last_probe_time": self._last_probe_time,
            }

    def is_local_active(self) -> bool:
        """Non-blocking: returns cached state and schedules a background probe if stale."""
        with self._lock:
            now = time.time()
            if self.is_3gpp_wifi and (now - self._last_probe_time > 5.0) and not self._is_probing:
                self._is_probing = True
                self._trigger_background_probe()

            return (self.is_3gpp_wifi or self.is_local_reachable) and self.is_local_reachable

    def check_local_reachability_fast(self, host: str = "10.10.10.10", port: int = 80, timeout: float = 0.25) -> bool:
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except (socket.timeout, OSError):
            return False

    def _trigger_background_probe(self):
        def _worker():
            try:
                reachable = self.check_local_reachability_fast("10.10.10.10", port=80, timeout=0.3)
                with self._lock:
                    self.is_local_reachable = reachable
                    self._last_probe_time = time.time()
            except Exception as e:
                logging.debug(f"[NetworkState] Probe error: {e}")
                with self._lock:
                    self.is_local_reachable = False
            finally:
                with self._lock:
                    self._is_probing = False

        thread = threading.Thread(target=_worker, daemon=True, name="NetworkStateProbe")
        thread.start()
