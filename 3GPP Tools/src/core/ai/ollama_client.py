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
    "custom_binary_path": "",
    "enable_vulkan": False,
    "enable_igpu": False
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



def build_ollama_service_environment(cfg: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """Builds the environment for an Ollama daemon started by 3GPP Tools.

    Hardware settings are opt-in application overrides. When disabled, the
    corresponding environment variable is left untouched so an existing user
    or system environment configuration remains authoritative.
    """
    cfg = cfg or load_ollama_config()
    env = os.environ.copy()

    if bool(cfg.get("enable_vulkan", False)):
        env["OLLAMA_VULKAN"] = "1"

    if bool(cfg.get("enable_igpu", False)):
        # Integrated-GPU scheduling currently relies on the Vulkan path.
        env["OLLAMA_IGPU_ENABLE"] = "1"
        env["OLLAMA_VULKAN"] = "1"

    return env


def describe_ollama_hardware_overrides(cfg: Optional[Dict[str, Any]] = None) -> str:
    """Human-readable summary for logs/UI diagnostics."""
    cfg = cfg or load_ollama_config()
    enabled = []
    if bool(cfg.get("enable_vulkan", False)) or bool(cfg.get("enable_igpu", False)):
        enabled.append("Vulkan")
    if bool(cfg.get("enable_igpu", False)):
        enabled.append("iGPU")
    return ", ".join(enabled) if enabled else "system/default"


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

        service_env = build_ollama_service_environment(cfg)
        logging.info(
            "[Ollama] Starting local service with hardware overrides: %s",
            describe_ollama_hardware_overrides(cfg),
        )
        subprocess.Popen(
            [binary, "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
            env=service_env,
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


class OllamaDiagnosticsWorker(QThread):
    """Fetches Ollama runtime/model diagnostics without blocking the Qt GUI."""

    finished = pyqtSignal(dict)

    def __init__(self, client, parent: QObject = None):
        super().__init__(parent)
        self.client = client

    def run(self):
        try:
            self.finished.emit(self.client.get_diagnostics())
        except Exception as exc:
            logging.exception("[Ollama] Diagnostics worker failed")
            self.finished.emit({"online": False, "error": str(exc), "version": "", "models": [], "running_models": []})


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
        self.keep_alive = cfg.get("keep_alive", "5m")
        self.session = get_ai_session(self.host, self.proxy_mode)

    def reconfigure(self, host: str, proxy_mode: str):
        """Re-initializes session when user changes host or proxy settings."""
        self.host = host.rstrip("/")
        if self.host.startswith("http://localhost:"):
            self.host = self.host.replace("http://localhost:", "http://127.0.0.1:")
        self.proxy_mode = proxy_mode
        cfg = load_ollama_config()
        self.timeout = cfg.get("timeout", 60)
        self.keep_alive = cfg.get("keep_alive", "5m")
        self.session = get_ai_session(self.host, self.proxy_mode)

    def chat(
        self,
        model: str,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        options: Optional[Dict[str, Any]] = None,
        keep_alive: Optional[str] = None,
        timeout: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Executes a complete non-streaming /api/chat request.

        Preserves the complete Ollama response, including native tool calls and
        timing/token metadata. Hardware selection belongs to the Ollama daemon,
        not to individual chat requests.
        """
        model = str(model or "").strip()
        if not model:
            raise ValueError("Ollama chat requires a model name.")
        if not isinstance(messages, list) or not messages:
            raise ValueError("Ollama chat requires at least one message.")

        payload: Dict[str, Any] = {
            "model": model,
            "messages": messages,
            "stream": False,
            "options": options or {"temperature": 0.2},
            "keep_alive": keep_alive if keep_alive is not None else self.keep_alive,
        }
        if tools:
            payload["tools"] = tools

        request_timeout = float(timeout if timeout is not None else self.timeout)
        endpoint = f"{self.host}/api/chat"
        with self.session.post(
            endpoint,
            json=payload,
            timeout=(5.0, request_timeout),
        ) as resp:
            resp.raise_for_status()
            data = resp.json()

        if not isinstance(data, dict):
            raise ValueError("Ollama /api/chat returned a non-object JSON response.")
        return data

    @staticmethod
    def _format_bytes(value: Any) -> str:
        """Formats Ollama byte counts for compact diagnostics UI display."""
        try:
            size = float(value or 0)
        except (TypeError, ValueError):
            return "-"
        if size <= 0:
            return "-"
        units = ("B", "KB", "MB", "GB", "TB")
        unit = 0
        while size >= 1024.0 and unit < len(units) - 1:
            size /= 1024.0
            unit += 1
        return f"{size:.1f} {units[unit]}" if unit >= 2 else f"{size:.0f} {units[unit]}"

    @staticmethod
    def _processor_summary(size: Any, size_vram: Any) -> str:
        """Derives Ollama's CPU/GPU split from /api/ps memory placement."""
        try:
            total = float(size or 0)
            vram = max(0.0, float(size_vram or 0))
        except (TypeError, ValueError):
            return "Unknown"
        if total <= 0:
            return "Unknown"
        gpu_pct = max(0, min(100, round((vram / total) * 100)))
        cpu_pct = 100 - gpu_pct
        if gpu_pct >= 99:
            return "100% GPU"
        if gpu_pct <= 1:
            return "100% CPU"
        return f"{gpu_pct}% GPU / {cpu_pct}% CPU"

    def get_diagnostics(self) -> Dict[str, Any]:
        """Returns version, installed-model metadata, and current runtime placement.

        /api/tags supplies installed model metadata. /api/ps supplies models currently
        loaded in memory, including size_vram and context_length. The GPU/CPU split is
        derived from size_vram / size and therefore describes observed runtime placement,
        not merely the configured Vulkan/iGPU preference.
        """
        result: Dict[str, Any] = {
            "online": False,
            "error": "",
            "version": "Unknown",
            "models": [],
            "running_models": [],
        }

        # Reuse the normal connectivity path first so diagnostics fail quickly offline.
        online, _, err = self.ping_and_get_models()
        if not online:
            result["error"] = err
            return result
        result["online"] = True

        # Version is informational: older Ollama builds may not expose this endpoint.
        try:
            resp = self.session.get(f"{self.host}/api/version", timeout=2.0)
            if resp.status_code == 200:
                result["version"] = str(resp.json().get("version") or "Unknown")
        except Exception as exc:
            logging.debug("[Ollama] Version diagnostics unavailable: %s", exc)

        try:
            resp = self.session.get(f"{self.host}/api/tags", timeout=3.0)
            resp.raise_for_status()
            data = resp.json()
            for item in data.get("models", []):
                details = item.get("details") or {}
                result["models"].append({
                    "name": str(item.get("name") or item.get("model") or ""),
                    "size": self._format_bytes(item.get("size")),
                    "parameter_size": str(details.get("parameter_size") or "-"),
                    "quantization": str(details.get("quantization_level") or "-"),
                    "family": str(details.get("family") or "-"),
                })
        except Exception as exc:
            result["error"] = f"Installed model query failed: {exc}"
            return result

        # Runtime information is optional so diagnostics still work with older servers.
        try:
            resp = self.session.get(f"{self.host}/api/ps", timeout=3.0)
            if resp.status_code == 200:
                data = resp.json()
                for item in data.get("models", []):
                    result["running_models"].append({
                        "name": str(item.get("name") or item.get("model") or ""),
                        "processor": self._processor_summary(item.get("size"), item.get("size_vram")),
                        "size": self._format_bytes(item.get("size")),
                        "vram": self._format_bytes(item.get("size_vram")),
                        "context_length": item.get("context_length") or "-",
                    })
            else:
                logging.debug("[Ollama] /api/ps unavailable: HTTP %s", resp.status_code)
        except Exception as exc:
            logging.debug("[Ollama] Runtime diagnostics unavailable: %s", exc)

        return result

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
            "options": options or {"temperature": 0.2},
            "keep_alive": self.keep_alive,
        }

        with self.session.post(
            endpoint,
            json=payload,
            stream=True,
            timeout=(5.0, float(self.timeout)),
        ) as resp:
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