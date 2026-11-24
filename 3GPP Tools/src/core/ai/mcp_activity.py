# --- File: src/core/ai/mcp_activity.py ---
"""Backward-compatible re-export shim for ``core.ai.mcp``."""

from __future__ import annotations

import logging

from .mcp import (
    MAX_ACTIVITY_LOG_LINES,
    TRIM_TARGET_LINES,
    _maybe_prune_activity_log,
    _sanitize_for_json,
    clean_stale_mcp_sessions,
    clear_mcp_activity_log,
    generate_call_summary,
    get_active_mcp_sessions,
    get_mcp_activity_log_path,
    get_mcp_activity_stats,
    get_mcp_client_configurations,
    get_mcp_logs_dir,
    get_mcp_sessions_dir,
    get_recent_mcp_events,
    heartbeat_mcp_session,
    is_pid_alive,
    record_mcp_event,
    register_mcp_session,
    unregister_mcp_session,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
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
