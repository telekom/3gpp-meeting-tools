"""Cross-process MCP session lockfile registration, heartbeats, and cleanup."""

from __future__ import annotations

import datetime
import json
import logging
from pathlib import Path
from typing import Any, Dict, List

from .paths_and_pids import get_mcp_sessions_dir, is_pid_alive, resolve_facade_attr

logger = logging.getLogger(__name__)


def _sessions_dir() -> Path:
    return resolve_facade_attr("get_mcp_sessions_dir", get_mcp_sessions_dir)()


def register_mcp_session(
    pid: int,
    client: str = "unknown",
    version: str = "?",
    profile: str = "all",
    read_only: bool = False,
) -> Path:
    """
    Register an active MCP server process in the session directory.
    Cleans up stale session files for dead PIDs.
    """
    sessions_dir = _sessions_dir()
    try:
        sessions_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    clean_stale_mcp_sessions()

    session_file = sessions_dir / f"session_{pid}.json"
    now_iso = datetime.datetime.now().isoformat()

    data = {
        "pid": pid,
        "client": client or "unknown",
        "version": version or "?",
        "profile": profile or "all",
        "read_only": bool(read_only),
        "start_time": now_iso,
        "last_heartbeat": now_iso,
    }

    try:
        session_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.debug("Could not write MCP session file: %s", exc)

    return session_file


def heartbeat_mcp_session(pid: int) -> None:
    """Update heartbeat timestamp for an active MCP server session."""
    session_file = _sessions_dir() / f"session_{pid}.json"
    if not session_file.exists():
        return
    try:
        data = json.loads(session_file.read_text(encoding="utf-8"))
        data["last_heartbeat"] = datetime.datetime.now().isoformat()
        session_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


def unregister_mcp_session(pid: int) -> None:
    """Remove session file on process shutdown."""
    session_file = _sessions_dir() / f"session_{pid}.json"
    try:
        if session_file.exists():
            session_file.unlink()
    except Exception:
        pass


def clean_stale_mcp_sessions() -> None:
    """Inspect all session files and delete those belonging to terminated processes."""
    sessions_dir = _sessions_dir()
    if not sessions_dir.exists():
        return
    try:
        for f in sessions_dir.glob("session_*.json"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                pid = data.get("pid")
                if not pid or not is_pid_alive(int(pid)):
                    f.unlink(missing_ok=True)
            except Exception:
                f.unlink(missing_ok=True)
    except Exception:
        pass


def get_active_mcp_sessions() -> List[Dict[str, Any]]:
    """Return list of all currently active MCP server sessions."""
    clean_stale_mcp_sessions()
    sessions = []
    sessions_dir = _sessions_dir()
    if not sessions_dir.exists():
        return []
    try:
        for f in sorted(sessions_dir.glob("session_*.json")):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                pid = data.get("pid")
                if pid and is_pid_alive(int(pid)):
                    sessions.append(data)
            except Exception:
                pass
    except Exception:
        pass
    return sessions
