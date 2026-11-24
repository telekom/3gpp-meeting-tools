"""Local Ollama daemon discovery and start/stop/restart lifecycle helpers."""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from .config import is_local_host, is_ollama_port_open, load_ollama_config


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
                        timeout=5.0,
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
