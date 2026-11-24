"""MCP telemetry paths, log retention constants, and Win32/POSIX PID checks."""

from __future__ import annotations

import ctypes
import logging
import os
import platform
import sys
from pathlib import Path

if platform.system() == "Windows":
    try:
        import ctypes.wintypes
    except Exception:
        pass

from core.utils.paths import get_app_data_root

logger = logging.getLogger(__name__)

# Max lines preserved in mcp_activity.jsonl before pruning older entries
MAX_ACTIVITY_LOG_LINES = 2500
TRIM_TARGET_LINES = 1500


def resolve_facade_attr(name: str, default):
    """Resolve hook/path helpers dynamically from core.ai.mcp_activity if patched."""
    facade = sys.modules.get("core.ai.mcp_activity")
    return getattr(facade, name, default) if facade is not None else default


def get_mcp_logs_dir() -> Path:
    """Return the application directory dedicated to MCP telemetry."""
    d = get_app_data_root() / "logs"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def get_mcp_activity_log_path() -> Path:
    """Return the path to the append-only MCP activity journal (`mcp_activity.jsonl`)."""
    return get_mcp_logs_dir() / "mcp_activity.jsonl"


def get_mcp_sessions_dir() -> Path:
    """Return the directory containing live MCP session lockfiles."""
    d = get_mcp_logs_dir() / "mcp_sessions"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def is_pid_alive(pid: int) -> bool:
    """Check if a process ID is currently running on the operating system."""
    if pid <= 0:
        return False

    if platform.system() == "Windows":
        try:
            kernel32 = ctypes.windll.kernel32
            SYNCHRONIZE = 0x00100000
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            handle = kernel32.OpenProcess(
                SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
            if not handle:
                return False
            exit_code = ctypes.wintypes.DWORD()
            success = kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code))
            kernel32.CloseHandle(handle)
            if not success:
                return False
            STILL_ACTIVE = 259
            return exit_code.value == STILL_ACTIVE
        except Exception:
            return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except (OSError, ProcessLookupError):
            return False
