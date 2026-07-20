# --- File: src/core/network/network_state.py ---
import threading

class NetworkState:
    """
    A thread-safe Singleton that holds the current status of the user's network connection.
    This allows any part of the application (like the UI or download threads) to instantly
    check if the local 3GPP server is reachable without waiting for PyQt signals.
    """
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(NetworkState, cls).__new__(cls)
                cls._instance.network_name = ""
                cls._instance.is_3gpp_wifi = False
                cls._instance.is_local_reachable = False
        return cls._instance

    @classmethod
    def get_instance(cls):
        return cls()

    def update_state(self, name: str, is_3gpp: bool, reachable: bool):
        with self._lock:
            self.network_name = name
            self.is_3gpp_wifi = is_3gpp
            self.is_local_reachable = reachable

    def is_local_active(self) -> bool:
        with self._lock:
            return self.is_3gpp_wifi and self.is_local_reachable