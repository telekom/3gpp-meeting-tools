# `core.ai.mcp`

Acyclic package backing the legacy `core.ai.mcp_activity` shim.

## Architecture

- `paths_and_pids.py` — Telemetry directory/file paths, retention thresholds (`MAX_ACTIVITY_LOG_LINES`, `TRIM_TARGET_LINES`), Win32/POSIX PID liveness (`is_pid_alive`), and dynamic facade resolution.
- `sessions.py` — Live session lockfile lifecycle (`register_mcp_session`, `heartbeat_mcp_session`, `unregister_mcp_session`, `clean_stale_mcp_sessions`, `get_active_mcp_sessions`).
- `telemetry.py` — Event summary formatting, JSONL journaling (`record_mcp_event`, `get_recent_mcp_events`, `clear_mcp_activity_log`, `get_mcp_activity_stats`), and log pruning.
- `client_configs.py` — Auto-detected config snippets for Claude Desktop, Codex CLI, Claude CLI, Gemini CLI, Cursor, and Windsurf (`get_mcp_client_configurations`).
- `__init__.py` — Public package exports.

## Dependency Flow

`paths_and_pids` -> `sessions` -> `telemetry` (`client_configs` independent) -> `__init__`.
