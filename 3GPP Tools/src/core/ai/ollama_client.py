# --- File: src/core/ai/ollama_client.py ---
"""
Ollama Core Client, Background Status Monitor, and Local Service Controller.
Provides persistent configuration, REST communication, non-blocking
heartbeat monitoring, and daemon lifecycle management (start/stop/restart).
"""

import json
import logging
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Tuple, List, Dict, Any, Optional

from PyQt5.QtCore import QThread, pyqtSignal, QObject

from core.network.session import get_ai_session
from core.utils.paths import get_project_root

CONFIG_PATH = get_project_root() / "config" / "ollama_config.json"

DEFAULT_CONFIG: Dict[str, Any] = {
    "host": "http://127.0.0.1:11434",
    "selected_model": "",
    "proxy_mode": "direct",
    "timeout": 60,
    "check_interval": 10,
    "keep_alive": "5m",
    "custom_binary_path": ""
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

    # Enforce IPv4 loopback to prevent Windows IPv6 or corporate DNS timeouts
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


def is_local_host(host: str) -> bool:
    """Checks whether the given Ollama host URL points to the local machine."""
    try:
        parsed = urllib.parse.urlsplit(host)
        hostname = (parsed.hostname or "127.0.0.1").lower()
        return hostname in ("127.0.0.1", "localhost", "0.0.0.0", "::1")
    except Exception:
        return False


def is_ollama_port_open(host: str = None) -> bool:
    """Checks whether the Ollama server port is accepting TCP connections (< 500ms)."""
    cfg = load_ollama_config()
    target_host = host or cfg.get("host", "http://127.0.0.1:11434")
    try:
        parsed = urllib.parse.urlsplit(target_host)
        host_ip = parsed.hostname or "127.0.0.1"
        port = parsed.port or 11434
        with socket.create_connection((host_ip, port), timeout=0.5):
            return True
    except OSError:
        return False
    except Exception:
        return False


def find_ollama_binary(custom_path: str = "") -> Optional[str]:
    """
    Locates the Ollama executable on the system:
    1. Checks explicitly provided custom_path.
    2. Checks saved custom_binary_path from config.
    3. Checks system PATH via shutil.which.
    4. Scans standard platform-specific installation paths.
    """
    if custom_path and Path(custom_path).is_file():
        return str(Path(custom_path).resolve())

    cfg = load_ollama_config()
    saved_path = cfg.get("custom_binary_path", "").strip()
    if saved_path and Path(saved_path).is_file():
        return str(Path(saved_path).resolve())

    found = shutil.which("ollama")
    if found:
        return str(Path(found).resolve())

    if os.name == "nt":
        candidates = [
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Ollama" / "ollama.exe",
            Path(os.environ.get("ProgramFiles", "C:\\Program Files")) / "Ollama" / "ollama.exe",
            Path(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")) / "Ollama" / "ollama.exe",
        ]
    elif sys.platform == "darwin":
        candidates = [
            Path("/usr/local/bin/ollama"),
            Path("/opt/homebrew/bin/ollama"),
            Path("/Applications/Ollama.app/Contents/Resources/ollama"),
        ]
    else:  # Linux / Unix
        candidates = [
            Path("/usr/local/bin/ollama"),
            Path("/usr/bin/ollama"),
            Path(os.path.expanduser("~/.local/bin/ollama")),
        ]

    for p in candidates:
        if p.is_file() and os.access(p, os.X_OK if os.name != "nt" else os.F_OK):
            return str(p.resolve())

    return None


def start_ollama_service(custom_path: str = "", host: str = None) -> Tuple[bool, str]:
    """
    Starts the local Ollama daemon in the background without creating console windows.
    Polls until the service is active or times out.
    """
    cfg = load_ollama_config()
    target_host = host or cfg.get("host", "http://127.0.0.1:11434")

    if not is_local_host(target_host):
        return False, f"Host '{target_host}' is remote. Service control is only available for local instances."

    if is_ollama_port_open(target_host):
        return True, "Ollama service is already running."

    binary = find_ollama_binary(custom_path)
    if not binary:
        return False, "Ollama executable not found. Please install Ollama or specify its path."

    try:
        creationflags = 0
        if os.name == "nt":
            creationflags = 0x08000000  # CREATE_NO_WINDOW

        subprocess.Popen(
            [binary, "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags
        )
    except Exception as e:
        return False, f"Failed to launch Ollama: {e}"

    # Poll for up to 10 seconds for service availability
    for _ in range(25):
        time.sleep(0.4)
        if is_ollama_port_open(target_host):
            return True, "Ollama service started successfully."

    return False, "Ollama process launched, but port did not respond within 10 seconds."


def stop_ollama_service(host: str = None) -> Tuple[bool, str]:
    """
    Stops the local Ollama daemon by terminating running process trees.
    """
    cfg = load_ollama_config()
    target_host = host or cfg.get("host", "http://127.0.0.1:11434")

    if not is_local_host(target_host):
        return False, f"Host '{target_host}' is remote. Service control is only available for local instances."

    if not is_ollama_port_open(target_host):
        return True, "Ollama service is not running."

    try:
        if os.name == "nt":
            creationflags = 0x08000000
            for proc_name in ("ollama.exe", "ollama_llama_server.exe", "ollama app.exe"):
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/IM", proc_name],
                        capture_output=True,
                        creationflags=creationflags,
                        timeout=5.0
                    )
                except Exception:
                    pass
        else:
            try:
                subprocess.run(["pkill", "-f", "ollama"], capture_output=True, timeout=5.0)
            except Exception:
                pass
    except Exception as e:
        return False, f"Failed to terminate Ollama: {e}"

    # Verify shutdown
    for _ in range(15):
        time.sleep(0.2)
        if not is_ollama_port_open(target_host):
            return True, "Ollama service stopped successfully."

    return False, "Ollama service did not shut down within the expected time."


def restart_ollama_service(custom_path: str = "", host: str = None) -> Tuple[bool, str]:
    """Restarts the local Ollama service."""
    stop_ollama_service(host)
    time.sleep(0.5)
    return start_ollama_service(custom_path, host)


class OllamaServiceWorker(QThread):
    """
    Background worker thread to execute service start/stop/restart operations
    without stalling the Qt GUI event loop.
    """
    finished = pyqtSignal(bool, str)
    progress_status = pyqtSignal(str)

    def __init__(self, action: str, custom_path: str = "", host: str = "", parent: QObject = None):
        super().__init__(parent)
        self.action = action.lower()
        self.custom_path = custom_path
        self.host = host

    def run(self):
        if self.action == "start":
            self.progress_status.emit("⏳ Starting Ollama service...")
            success, msg = start_ollama_service(self.custom_path, self.host)
        elif self.action == "stop":
            self.progress_status.emit("⏳ Stopping Ollama service...")
            success, msg = stop_ollama_service(self.host)
        elif self.action == "restart":
            self.progress_status.emit("⏳ Stopping Ollama service...")
            stop_ollama_service(self.host)
            time.sleep(0.5)
            self.progress_status.emit("⏳ Starting Ollama service...")
            success, msg = start_ollama_service(self.custom_path, self.host)
        else:
            success, msg = False, f"Unknown action: {self.action}"

        self.finished.emit(success, msg)


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
        Fast non-blocking probe:
        - Raw socket pre-check rejects offline instances in < 1ms without HTTP overhead.
        - REST request queries /api/tags if the port is open.
        """
        parsed = urllib.parse.urlsplit(self.host)
        host_ip = parsed.hostname or "127.0.0.1"
        port = parsed.port or 11434

        # 1. Pure TCP pre-check (< 1ms when port is closed)
        try:
            with socket.create_connection((host_ip, port), timeout=0.5):
                pass
        except OSError:
            return False, [], "Ollama server is not running."
        except Exception as e:
            return False, [], f"Connection error: {e}"

        # 2. Query /api/tags
        endpoint = f"{self.host}/api/tags"
        try:
            resp = self.session.get(endpoint, timeout=2.0)
            if resp.status_code == 200:
                data = resp.json()
                models = [str(m.get("name", "")) for m in data.get("models", []) if m.get("name")]
                return True, models, ""
            return False, [], f"HTTP {resp.status_code}: {resp.reason}"
        except Exception as e:
            return False, [], f"API error: {e}"

    def stream_chat(self, model: str, messages: List[Dict[str, str]], options: Optional[Dict[str, Any]] = None):
        """
        Executes a streaming request against /api/chat.
        Yields text chunks as they arrive.
        """
        endpoint = f"{self.host}/api/chat"
        payload = {
            "model": model,
            "messages": messages,
            "stream": True,
            "options": options or {"temperature": 0.2}
        }

        with self.session.post(endpoint, json=payload, stream=True, timeout=(5.0, 300.0)) as resp:
            resp.raise_for_status()
            for line in resp.iter_lines():
                if not line:
                    continue
                try:
                    chunk = json.loads(line.decode("utf-8"))
                    delta = chunk.get("message", {}).get("content", "")
                    if delta:
                        yield delta
                    if chunk.get("done", False):
                        break
                except Exception as parse_err:
                    logging.warning(f"[Ollama] Stream chunk parse error: {parse_err}")


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
        self._force_check = False
        self._last_state: Tuple[Optional[bool], str, int, str] = (None, "", -1, "")

    def trigger_check(self):
        """Forces an immediate heartbeat check without waiting for the sleep interval."""
        self._force_check = True

    def run(self):
        logging.info("🦙 [Ollama Monitor] Worker thread started running.")

        while self._running:
            try:
                cfg = load_ollama_config()
                check_interval = max(5, int(cfg.get("check_interval", 10)))
                selected_model = str(cfg.get("selected_model") or "")

                is_online, models, err = self.client.ping_and_get_models()

                # If selected model is missing but models exist, pick the first
                if is_online and models and not selected_model:
                    selected_model = str(models[0])
                    cfg["selected_model"] = selected_model
                    save_ollama_config(cfg)

                current_state = (bool(is_online), str(selected_model or ""), len(models), str(err or ""))

                if current_state != self._last_state:
                    self._last_state = current_state
                    logging.info(
                        f"🦙 [Ollama Monitor] Status changed: online={is_online}, "
                        f"models={len(models)}, active='{selected_model}', err='{err}'"
                    )

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

            # Sleep in 250ms intervals to allow immediate shutdown or forced update
            for _ in range(check_interval * 4):
                if not self._running or self._force_check:
                    self._force_check = False
                    break
                self.msleep(250)

        logging.info("🦙 [Ollama Monitor] Worker thread cleanly exited run loop.")

    def stop(self):
        """Stops the polling thread gracefully on application exit."""
        self._running = False
        self.quit()
        if not self.wait(600):
            self.terminate()
            self.wait(200)