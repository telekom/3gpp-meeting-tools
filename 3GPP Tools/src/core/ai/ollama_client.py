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
    "host": "http://localhost:11434",
    "selected_model": "",
    "proxy_mode": "direct",  # "direct", "auto", "app_proxy"
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
        self.host = (host or cfg.get("host", "http://localhost:11434")).rstrip("/")
        self.proxy_mode = proxy_mode or cfg.get("proxy_mode", "direct")
        self.timeout = cfg.get("timeout", 60)
        self.session = get_ai_session(self.host, self.proxy_mode)

    def reconfigure(self, host: str, proxy_mode: str):
        """Re-initializes session when user changes host or proxy settings."""
        self.host = host.rstrip("/")
        self.proxy_mode = proxy_mode
        self.session = get_ai_session(self.host, self.proxy_mode)

    def ping_and_get_models(self) -> Tuple[bool, List[str], str]:
        """
        Polls Ollama /api/tags to verify liveness and retrieve installed models.
        Returns: (is_online: bool, installed_models: list[str], error_message: str)
        """
        endpoint = f"{self.host}/api/tags"
        try:
            resp = self.session.get(endpoint, timeout=4)
            if resp.status_code == 200:
                data = resp.json()
                models = [m.get("name", "") for m in data.get("models", []) if m.get("name")]
                return True, models, ""
            return False, [], f"HTTP {resp.status_code}: {resp.reason}"
        except Exception as e:
            err_str = str(e)
            if "Connection refused" in err_str or "actively refused" in err_str:
                return False, [], "Ollama server is not running."
            return False, [], "Connection failed."


class OllamaMonitorThread(QThread):
    """
    Non-blocking background thread that periodically monitors Ollama health
    and notifies the UI when connectivity or model lists change.
    """
    # Emits: (is_online, selected_model, available_models, error_message)
    status_updated = pyqtSignal(bool, str, list, str)

    def __init__(self, client: OllamaClient = None, parent: QObject = None):
        super().__init__(parent)
        self.client = client or OllamaClient()
        self._running = True
        self._check_interval = 10

    def run(self):
        while self._running:
            cfg = load_ollama_config()
            self._check_interval = max(5, cfg.get("check_interval", 10))
            selected_model = cfg.get("selected_model", "")

            is_online, models, err = self.client.ping_and_get_models()

            # If the selected model is not set but models exist, auto-select the first
            if is_online and models and not selected_model:
                selected_model = models[0]
                cfg["selected_model"] = selected_model
                save_ollama_config(cfg)

            self.status_updated.emit(is_online, selected_model, models, err)

            # Responsive sleep loop (checks _running every 500ms)
            for _ in range(self._check_interval * 2):
                if not self._running:
                    break
                self.msleep(500)

    def stop(self):
        """Stops the polling thread gracefully on application exit."""
        self._running = False
        self.quit()
        self.wait(2000)