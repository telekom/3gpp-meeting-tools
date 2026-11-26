"""Persistent Ollama configuration, host normalization, and TCP probe helpers."""

from __future__ import annotations

import json
import logging
import socket
import sys
import urllib.parse
from pathlib import Path
from typing import Any, Dict

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
    "enable_igpu": False,
}


def _resolve_config_path() -> Path:
    """Resolve CONFIG_PATH dynamically from the legacy shim if monkeypatched."""
    facade = sys.modules.get("core.ai.ollama_client")
    if facade is not None and hasattr(facade, "CONFIG_PATH"):
        return Path(facade.CONFIG_PATH)
    return CONFIG_PATH


def _normalize_ollama_host(host: str) -> str:
    """Normalizes localhost URLs (with or without port) to IPv4 loopback 127.0.0.1."""
    h = str(host or "").strip().rstrip("/")
    if not h:
        return "http://127.0.0.1:11434"
    for prefix in ("http://localhost", "https://localhost"):
        if h.lower().startswith(prefix):
            rest = h[len(prefix) :]
            if not rest or rest.startswith(":") or rest.startswith("/"):
                scheme = h.split("://", 1)[0]
                return f"{scheme}://127.0.0.1{rest}"
    return h


def load_ollama_config() -> dict:
    """Loads Ollama configuration from disk with safe defaults."""
    cfg = DEFAULT_CONFIG.copy()
    cfg_path = _resolve_config_path()
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            logging.warning(f"[Ollama] Error reading {cfg_path.name}: {e}")

    # Enforce IPv4 loopback to prevent Windows IPv6 or corporate DNS timeouts
    cfg["host"] = _normalize_ollama_host(cfg.get("host", ""))

    return cfg


def save_ollama_config(cfg_data: dict) -> None:
    """Persists Ollama configuration to disk."""
    cfg_path = _resolve_config_path()
    try:
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        with open(cfg_path, "w", encoding="utf-8") as f:
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
