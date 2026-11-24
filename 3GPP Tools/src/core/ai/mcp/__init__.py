"""Modular MCP telemetry, session lockfiles, and client configuration package."""

from __future__ import annotations

from .client_configs import get_mcp_client_configurations
from .paths_and_pids import (
    MAX_ACTIVITY_LOG_LINES,
    TRIM_TARGET_LINES,
    get_mcp_activity_log_path,
    get_mcp_logs_dir,
    get_mcp_sessions_dir,
    is_pid_alive,
)
from .sessions import (
    clean_stale_mcp_sessions,
    get_active_mcp_sessions,
    heartbeat_mcp_session,
    register_mcp_session,
    unregister_mcp_session,
)
from .telemetry import (
    _maybe_prune_activity_log,
    _sanitize_for_json,
    clear_mcp_activity_log,
    generate_call_summary,
    get_mcp_activity_stats,
    get_recent_mcp_events,
    record_mcp_event,
)

__all__ = [
    "MAX_ACTIVITY_LOG_LINES",
    "TRIM_TARGET_LINES",
    "get_mcp_logs_dir",
    "get_mcp_activity_log_path",
    "get_mcp_sessions_dir",
    "is_pid_alive",
    "register_mcp_session",
    "heartbeat_mcp_session",
    "unregister_mcp_session",
    "clean_stale_mcp_sessions",
    "get_active_mcp_sessions",
    "generate_call_summary",
    "_sanitize_for_json",
    "record_mcp_event",
    "_maybe_prune_activity_log",
    "get_recent_mcp_events",
    "clear_mcp_activity_log",
    "get_mcp_activity_stats",
    "get_mcp_client_configurations",
]
