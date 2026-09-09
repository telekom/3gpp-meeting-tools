# --- File: src/core/ai/ollama_client.py ---
"""
Ollama Core Client & Background Status Monitor.
Provides persistent configuration, REST communication, and non-blocking
heartbeat monitoring for local LLM integration.
"""

import json
import logging
from typing import Tuple, List, Dict, Any

from PyQt5.QtCore import QThread, pyqtSignal, QObject

from core.network.session import get_ai_session
from core.utils.paths import get_project_root

CONFIG_PATH = get_project_root() / "ollama_config.json"

DEFAULT_CONFIG: Dict[str, Any] = {
    "host": "http://127.0.0.1:11434",
    "selected_model": "",
    "proxy_mode": "direct",
    "timeout": 60,
    "check_interval": 10,
    "keep_alive": "5m"
}


def load_ollama_config() -> dict:
    """Loads Ollama configuration from disk with safe defaults."""
    cfg = DEFAULT_CONFIG.copy()
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            logging.warning(f"[Ollama] Error reading {CONFIG_PATH.name}: {e}")

    # Normalize localhost to direct IPv4 to prevent Windows IPv6 / corporate DNS stalls
    host = str(cfg.get("host", "")).strip()
    if host.startswith("http://localhost:"):
        cfg["host"] = host.replace("http://localhost:", "http://127.0.0.1:")
    elif not host:
        cfg["host"] = "http://127.0.0.1:11434"

    return cfg


def save_ollama_config(cfg_data: dict) -> None:
    """Persists Ollama configuration to disk."""
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg_data, f, indent=4)
    except Exception as e:
        logging.error(f"[Ollama] Failed to save config: {e}")


class OllamaClient:
    """Core communication client for the Ollama REST API."""

    def __init__(self, host: str = None, proxy_mode: str = None):
        cfg = load_ollama_config()
        raw_host = (host or cfg.get("host", "http://127.0.0.1:11434")).rstrip("/")
        if raw_host.startswith("http://localhost:"):
            raw_host = raw_host.replace("http://localhost:", "http://127.0.0.1:")
        self.host = raw_host
        self.proxy_mode = proxy_mode or cfg.get("proxy_mode", "direct")
        self.timeout = cfg.get("timeout", 60)
        self.session = get_ai_session(self.host, self.proxy_mode)

    def reconfigure(self, host: str, proxy_mode: str):
        """Re-initializes session when user changes host or proxy settings."""
        self.host = host.rstrip("/")
        if self.host.startswith("http://localhost:"):
            self.host = self.host.replace("http://localhost:", "http://127.0.0.1:")
        self.proxy_mode = proxy_mode
        self.session = get_ai_session(self.host, self.proxy_mode)

    def ping_and_get_models(self) -> Tuple[bool, List[str], str]:
        """
        Polls Ollama /api/tags to verify liveness and retrieve installed models.
        Uses a tight 2-second timeout so unresponsive connections fail fast.
        """
        endpoint = f"{self.host}/api/tags"
        try:
            resp = self.session.get(endpoint, timeout=(1.5, 2.0))
            if resp.status_code == 200:
                data = resp.json()
                models = [str(m.get("name", "")) for m in data.get("models", []) if m.get("name")]
                return True, models, ""
            return False, [], f"HTTP {resp.status_code}: {resp.reason}"
        except Exception as e:
            err_str = str(e)
            if "Connection refused" in err_str or "actively refused" in err_str or "10061" in err_str:
                return False, [], "Ollama server is not running."
            if "ConnectTimeout" in err_str or "timed out" in err_str:
                return False, [], "Connection timed out."
            return False, [], f"Connection failed ({type(e).__name__})"


class OllamaMonitorThread(QThread):
    """
    Non-blocking background thread that periodically monitors Ollama health
    and notifies the UI only when connectivity or model lists change.
    """
    status_updated = pyqtSignal(bool, str, list, str)

    def __init__(self, client: OllamaClient = None, parent: QObject = None):
        super().__init__(parent)
        self.client = client or OllamaClient()
        self._running = True
        self._last_state: Tuple[Optional[bool], str, int, str] = (None, "", -1, "")

    def run(self):
        logging.info("🦙 [Ollama Monitor] Worker thread started running.")

        while self._running:
            try:
                cfg = load_ollama_config()
                check_interval = max(5, int(cfg.get("check_interval", 10)))
                selected_model = str(cfg.get("selected_model") or "")

                is_online, models, err = self.client.ping_and_get_models()

                # If selected model is missing but models exist, auto-select the first
                if is_online and models and not selected_model:
                    selected_model = str(models[0])
                    cfg["selected_model"] = selected_model
                    save_ollama_config(cfg)

                # Current fingerprint: (online_bool, active_model, total_models, error_text)
                current_state = (bool(is_online), str(selected_model or ""), len(models), str(err or ""))

                if current_state != self._last_state:
                    self._last_state = current_state
                    logging.info(
                        f"🦙 [Ollama Monitor] Status changed: online={is_online}, "
                        f"models={len(models)}, active='{selected_model}', err='{err}'"
                    )
                else:
                    logging.debug(
                        f"🦙 [Ollama Monitor] Heartbeat unchanged: online={is_online}, active='{selected_model}'"
                    )

                # Emit signal to keep UI updated
                self.status_updated.emit(
                    bool(is_online),
                    str(selected_model or ""),
                    list(models or []),
                    str(err or "")
                )

            except Exception as e:
                logging.error(f"❌ [Ollama Monitor] Loop exception: {e}", exc_info=True)
                if self._running:
                    self.status_updated.emit(False, "", [], f"Monitor error: {e}")

            # Sleep in 250ms chunks to allow fast thread termination
            for _ in range(check_interval * 4):
                if not self._running:
                    break
                self.msleep(250)

        logging.info("🦙 [Ollama Monitor] Worker thread cleanly exited run loop.")

    def stop(self):
        """Stops the polling thread gracefully on application exit."""
        self._running = False
        self.quit()
        if not self.wait(1000):
            self.terminate()
            self.wait(300)