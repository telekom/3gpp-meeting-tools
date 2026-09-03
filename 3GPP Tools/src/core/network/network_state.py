import socket
import threading
import time
import logging

class NetworkState:
    """
    A thread-safe Singleton that holds the current status of the network connection.
    Guarantees that background probes cannot overlap.
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

    def is_local_active(self) -> bool:
        """Non-blocking: returns cached state and schedules a background probe if stale."""
        with self._lock:
            now = time.time()
            if self.is_3gpp_wifi and (now - self._last_probe_time > 5.0) and not self._is_probing:
                self._is_probing = True
                self._trigger_background_probe()

            return self.is_3gpp_wifi and self.is_local_reachable

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