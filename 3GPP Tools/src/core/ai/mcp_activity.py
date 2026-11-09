# --- File: src/core/ai/mcp_activity.py ---
"""
Model Context Protocol (MCP) Activity Logging, Session Tracking & Configuration.

Provides cross-process activity telemetry and session state tracking between
the standalone `mcp_server.py` / `cli_tools.py` processes and the 3GPP Tools
desktop interface, as well as configuration snippet generation for external
AI coding agents (Claude Desktop, Codex CLI, Claude CLI, Cursor, Windsurf).
"""

from __future__ import annotations

import atexit
import ctypes
import datetime
import json
import logging
import os
import platform
import shutil
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

if platform.system() == "Windows":
    try:
        import ctypes.wintypes
    except Exception:
        pass

from core.utils.paths import get_app_data_root, get_project_root

logger = logging.getLogger(__name__)

# Max lines preserved in mcp_activity.jsonl before pruning older entries
MAX_ACTIVITY_LOG_LINES = 2500
TRIM_TARGET_LINES = 1500


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


# =====================================================================
# Process Liveness Checks (Zero-overhead Win32 kernel API)
# =====================================================================

def is_pid_alive(pid: int) -> bool:
    """Check if a process ID is currently running on the operating system."""
    if pid <= 0:
        return False

    if platform.system() == "Windows":
        try:
            kernel32 = ctypes.windll.kernel32
            SYNCHRONIZE = 0x00100000
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            handle = kernel32.OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
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


# =====================================================================
# Session Registration & Heartbeats
# =====================================================================

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
    sessions_dir = get_mcp_sessions_dir()
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
    session_file = get_mcp_sessions_dir() / f"session_{pid}.json"
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
    session_file = get_mcp_sessions_dir() / f"session_{pid}.json"
    try:
        if session_file.exists():
            session_file.unlink()
    except Exception:
        pass


def clean_stale_mcp_sessions() -> None:
    """Inspect all session files and delete those belonging to terminated processes."""
    sessions_dir = get_mcp_sessions_dir()
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
    sessions_dir = get_mcp_sessions_dir()
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


# =====================================================================
# Activity Event Telemetry (Journaling)
# =====================================================================

def generate_call_summary(tool: str, action: str, args: Dict[str, Any], raw_data: Any) -> str:
    """Generate a crisp, human-readable one-sentence summary for the activity table."""
    try:
        if tool == "spec_text_and_clauses":
            spec = args.get("spec_number", "")
            if action == "search_text":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len((raw_data.get("items") or raw_data.get("matches") or raw_data.get("results") or []) if isinstance(raw_data, dict) else []))
                return f"TS {spec} search '{args.get('query', '')}' -> {count} clause hits"
            if action == "get_clause":
                cid = args.get("clause_id", "")
                length = len(str(raw_data.get("text", ""))) if isinstance(raw_data, dict) else 0
                return f"TS {spec} clause {cid} -> {length} characters"
            if action == "diff_versions":
                count = len(raw_data.get("clauses_changed", [])) if isinstance(raw_data, dict) else 0
                return f"TS {spec} diff {args.get('old_version', '')} vs {args.get('new_version', '')} -> {count} diffs"
            if action == "evolution_matrix":
                entries = len(raw_data.get("timeline", [])) if isinstance(raw_data, dict) else 0
                return f"TS {spec} clause {args.get('clause_id', '')} evolution across {entries} releases"

        elif tool == "protocol_inspector":
            proto = args.get("protocol", "")
            if action == "list_protocols":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len(raw_data.get("protocols", [])) if isinstance(raw_data, dict) else 0)
                return f"Listed {count} imported protocols (NAS, RRC, NGAP, PFCP...)"
            if action == "find_message":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len(raw_data.get("messages", [])) if isinstance(raw_data, dict) else 0)
                return f"{proto} find message '{args.get('query', '')}' -> {count} matches"
            if action == "find_ie":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len(raw_data.get("ies", [])) if isinstance(raw_data, dict) else 0)
                return f"{proto} find IE '{args.get('query', '')}' -> {count} matches"

        elif tool == "meetings_and_tdocs":
            wg = args.get("wg", "SA2")
            mtg = args.get("meeting_number", "")
            if action == "search_meetings":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len(raw_data.get("meetings", [])) if isinstance(raw_data, dict) else 0)
                return f"{wg} meeting search -> {count} meetings found"
            if action == "search_tdocs":
                count = (raw_data.get("count") if isinstance(raw_data, dict) and "count" in raw_data
                         else len(raw_data.get("tdocs", [])) if isinstance(raw_data, dict) else 0)
                co = args.get("source_company") or "all"
                return f"{wg}#{mtg} TDocs search (co: {co}) -> {count} TDocs matched"
            if action == "read_tdoc_markdown":
                length = len(str(raw_data.get("markdown", ""))) if isinstance(raw_data, dict) else 0
                return f"Read TDoc {args.get('tdoc_id', '')} -> {length} chars Markdown"
            if action == "meeting_stats":
                total = raw_data.get("total_tdocs", 0) if isinstance(raw_data, dict) else 0
                return f"{wg}#{mtg} statistics -> {total} total TDocs processed"
            if action == "get_or_set_stance":
                tdoc = args.get("tdoc_id", "")
                st = args.get("stance") or (raw_data.get("stance") if isinstance(raw_data, dict) else "queried")
                return f"TDoc {tdoc} stance: {st}"

        elif tool == "spec_history_and_attribution":
            spec = args.get("spec_number", "")
            if action == "query_history":
                count = len(raw_data.get("history", [])) if isinstance(raw_data, dict) else 0
                return f"TS {spec} change history query -> {count} revisions"
            if action == "attribute_clause":
                count = len(raw_data.get("candidates", [])) if isinstance(raw_data, dict) else 0
                return f"TS {spec} clause {args.get('clause_id', '')} attribution -> {count} candidate CRs"
            if action == "rank_origin_crs":
                count = len(raw_data.get("ranked_crs", [])) if isinstance(raw_data, dict) else 0
                return f"TS {spec} origin CR ranking for '{args.get('query', '')}' -> {count} CRs"

        elif tool == "onsite_meeting_cockpit":
            if action == "network_status":
                reach = raw_data.get("is_local_reachable") if isinstance(raw_data, dict) else False
                return f"Probed 10.10.10.10: local_reachable={reach}"
            if action == "browse_lan_server":
                count = len(raw_data.get("files", [])) if isinstance(raw_data, dict) else 0
                return f"Browse LAN {args.get('subfolder', '')} -> {count} files found"
            if action == "fetch_live_revisions":
                return f"Fetched live revisions for TDoc {args.get('tdoc_id', '')}"
            if action == "sync_chair_notes":
                return f"Synced Chairman Notes for {args.get('wg', 'SA2')}#{args.get('meeting_number', '')}"
            if action == "where_should_we_go_now":
                return "On-site agenda room & session routing recommendation"

        elif tool == "work_items_and_emails":
            if action == "search_work_items":
                count = len(raw_data.get("work_items", [])) if isinstance(raw_data, dict) else 0
                return f"Work Items search for '{args.get('query', '')}' -> {count} matches"
            if action == "summarize_email_thread":
                return f"Summarized reflector email thread '{args.get('subject', '')}'"

        elif tool == "cr_pcr_generator":
            if action == "generate_pcr":
                return f"Generated pCR contribution: {raw_data.get('docx_path', 'document.docx')}"
            if action == "generate_cr":
                return f"Generated formal CR contribution: {raw_data.get('docx_path', 'document.docx')}"
            if action == "preview_filename":
                return f"Filename preview: {raw_data.get('filename', '')}"

    except Exception:
        pass

    # Generic fallback summary
    summary_parts = [f"{tool}:{action}"]
    for k in ("spec_number", "clause_id", "tdoc_id", "query", "meeting_number"):
        if k in args and args[k]:
            summary_parts.append(f"{k}={args[k]}")
    return " | ".join(summary_parts)


def record_mcp_event(
    event_type: str,
    *,
    client: str = "unknown",
    client_ver: str = "?",
    pid: Optional[int] = None,
    tool: str = "",
    action: str = "",
    arguments: Optional[Dict[str, Any]] = None,
    elapsed_ms: float = 0.0,
    ok: bool = True,
    is_error: bool = False,
    error_msg: Optional[str] = None,
    raw_data: Any = None,
) -> None:
    """
    Append an activity record to `mcp_activity.jsonl` safely and non-blockingly.
    Never throws exceptions into callers.
    """
    try:
        log_path = get_mcp_activity_log_path()
        now_dt = datetime.datetime.now()
        cur_pid = pid if pid is not None else os.getpid()

        # Update session heartbeat
        heartbeat_mcp_session(cur_pid)

        args = arguments or {}
        summary = generate_call_summary(tool, action, args, raw_data)

        # Build payload preview string
        preview = ""
        chars_count = 0
        if raw_data is not None:
            try:
                preview = json.dumps(raw_data, ensure_ascii=False, indent=2, default=str)
                chars_count = len(preview)
                if len(preview) > 2000:
                    preview = preview[:2000] + "\n... [truncated preview]"
            except Exception:
                preview = str(raw_data)[:2000]
                chars_count = len(str(raw_data))
        elif error_msg:
            preview = f"Error: {error_msg}"
            chars_count = len(preview)

        entry = {
            "id": f"{int(now_dt.timestamp()*1000)}_{cur_pid}_{uuid.uuid4().hex[:6]}",
            "timestamp": now_dt.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
            "time_short": now_dt.strftime("%H:%M:%S"),
            "event_type": event_type,
            "client": client or "unknown",
            "client_ver": client_ver or "?",
            "pid": cur_pid,
            "tool": tool,
            "action": action,
            "args": args,
            "elapsed_ms": round(elapsed_ms, 1),
            "ok": bool(ok),
            "is_error": bool(is_error),
            "error": error_msg or "",
            "summary": summary,
            "response_chars": chars_count,
            "response_preview": preview,
        }

        line = json.dumps(entry, ensure_ascii=False) + "\n"

        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, mode="a", encoding="utf-8") as f:
            f.write(line)
            f.flush()

        # Self-trim if log gets excessively large
        _maybe_prune_activity_log(log_path)

    except Exception as exc:
        logger.debug("Failed to record MCP activity: %s", exc)


def _maybe_prune_activity_log(log_path: Path) -> None:
    """Trim older entries if the log file exceeds MAX_ACTIVITY_LOG_LINES."""
    try:
        if not log_path.exists():
            return
        if log_path.stat().st_size < 1_000_000:  # Skip line count if < 1MB
            return

        with open(log_path, mode="r", encoding="utf-8") as f:
            lines = f.readlines()

        if len(lines) > MAX_ACTIVITY_LOG_LINES:
            kept = lines[-TRIM_TARGET_LINES:]
            with open(log_path, mode="w", encoding="utf-8") as f:
                f.writelines(kept)
    except Exception:
        pass


def get_recent_mcp_events(limit: int = 250) -> List[Dict[str, Any]]:
    """
    Read the latest activity records from `mcp_activity.jsonl` in reverse
    chronological order (newest first).
    """
    log_path = get_mcp_activity_log_path()
    if not log_path.exists():
        return []

    try:
        with open(log_path, mode="r", encoding="utf-8") as f:
            lines = f.readlines()

        events = []
        for line in reversed(lines):
            line_str = line.strip()
            if not line_str:
                continue
            try:
                events.append(json.loads(line_str))
                if len(events) >= limit:
                    break
            except Exception:
                continue
        return events
    except Exception as exc:
        logger.debug("Error reading MCP events: %s", exc)
        return []


def clear_mcp_activity_log() -> None:
    """Clear the activity log file."""
    log_path = get_mcp_activity_log_path()
    try:
        if log_path.exists():
            log_path.write_text("", encoding="utf-8")
    except Exception:
        pass


def get_mcp_activity_stats() -> Dict[str, Any]:
    """Compute aggregate telemetry metrics over recent activity records."""
    events = get_recent_mcp_events(limit=1000)
    total_calls = len(events)
    if total_calls == 0:
        return {
            "total_calls": 0,
            "success_rate": 100.0,
            "avg_latency_ms": 0.0,
            "by_tool": {},
            "by_client": {},
            "last_active": None,
        }

    successes = sum(1 for e in events if e.get("ok") and not e.get("is_error"))
    latencies = [e.get("elapsed_ms", 0.0) for e in events if e.get("elapsed_ms") is not None]
    avg_lat = sum(latencies) / len(latencies) if latencies else 0.0

    by_tool: Dict[str, int] = {}
    by_client: Dict[str, int] = {}

    for e in events:
        t = e.get("tool") or "unknown"
        c = e.get("client") or "unknown"
        by_tool[t] = by_tool.get(t, 0) + 1
        by_client[c] = by_client.get(c, 0) + 1

    return {
        "total_calls": total_calls,
        "success_rate": round((successes / total_calls) * 100.0, 1),
        "avg_latency_ms": round(avg_lat, 1),
        "by_tool": by_tool,
        "by_client": by_client,
        "last_active": events[0].get("timestamp"),
    }


# =====================================================================
# AI Agent Client Configuration Snippet Generator
# =====================================================================

def get_mcp_client_configurations() -> Dict[str, Dict[str, Any]]:
    """
    Generate ready-to-use configuration blocks with auto-detected absolute
    paths for Claude Desktop, Codex CLI, Claude CLI, Cursor, and Windsurf.
    """
    root_dir = get_project_root()
    src_dir = root_dir / "src"
    mcp_server_script = src_dir / "mcp_server.py"
    cmd_launcher = src_dir / "Launch_MCP.cmd"
    python_exe = Path(sys.executable)

    # Validate paths
    has_python = python_exe.exists()
    has_server = mcp_server_script.exists()
    has_launcher = cmd_launcher.exists()

    server_script_str = str(mcp_server_script)
    python_exe_str = str(python_exe)
    cmd_launcher_str = str(cmd_launcher)

    # 1. Claude Desktop config (%APPDATA%\Claude\claude_desktop_config.json)
    claude_desktop_cfg = {
        "mcpServers": {
            "3gpp-tools": {
                "command": python_exe_str,
                "args": [
                    server_script_str,
                    "--profile",
                    "all",
                ],
            }
        }
    }

    # 2. Codex CLI configuration (~/.codex/config.json or flags)
    codex_cfg = {
        "mcp": {
            "servers": {
                "3gpp-tools": {
                    "command": python_exe_str,
                    "args": [server_script_str, "--profile", "all"],
                }
            }
        }
    }

    # 3. Claude CLI command snippet
    claude_cli_cmd = f'claude mcp add 3gpp-tools "{python_exe_str}" "{server_script_str}" -- --profile all'

    # 4. Cursor / Windsurf settings (Settings -> Features -> MCP -> Add Server)
    cursor_cfg = {
        "name": "3gpp-tools",
        "type": "command",
        "command": f'"{python_exe_str}" "{server_script_str}" --profile all',
    }

    # 5. Zero-config batch launcher alternative
    cmd_desktop_cfg = {
        "mcpServers": {
            "3gpp-tools": {
                "command": "cmd.exe",
                "args": [
                    "/c",
                    cmd_launcher_str,
                    "--profile",
                    "all",
                ],
            }
        }
    }

    # Detect Claude Desktop config location
    claude_config_file = Path.home() / "AppData" / "Roaming" / "Claude" / "claude_desktop_config.json"
    claude_config_dir = claude_config_file.parent

    return {
        "paths": {
            "python_exe": python_exe_str,
            "has_python": has_python,
            "server_script": server_script_str,
            "has_server": has_server,
            "cmd_launcher": cmd_launcher_str,
            "has_launcher": has_launcher,
            "claude_config_file": str(claude_config_file),
            "claude_config_exists": claude_config_file.exists(),
            "claude_config_dir": str(claude_config_dir),
        },
        "claude_desktop": {
            "title": "Claude Desktop",
            "target_file": str(claude_config_file),
            "config_json": json.dumps(claude_desktop_cfg, indent=2),
            "cmd_config_json": json.dumps(cmd_desktop_cfg, indent=2),
            "instructions": (
                "Add this entry to your claude_desktop_config.json under the 'mcpServers' key, "
                "then restart Claude Desktop. The 3GPP Tools hammer icon will appear in the input bar."
            ),
        },
        "codex_cli": {
            "title": "OpenAI Codex CLI",
            "config_json": json.dumps(codex_cfg, indent=2),
            "cli_command": f'codex --mcp "{python_exe_str} {server_script_str} --profile all"',
            "instructions": (
                "Pass the command above or add the JSON snippet into ~/.codex/config.json. "
                "Codex will automatically discover all 12 domain tools with zero token overhead."
            ),
        },
        "claude_cli": {
            "title": "Anthropic Claude CLI",
            "cli_command": claude_cli_cmd,
            "instructions": "Run this command in PowerShell or cmd to permanently register 3GPP Tools with Claude CLI.",
        },
        "cursor": {
            "title": "Cursor & Windsurf IDE",
            "config_json": json.dumps(cursor_cfg, indent=2),
            "command_line": f'"{python_exe_str}" "{server_script_str}" --profile all',
            "instructions": (
                "Go to Settings -> Features -> MCP -> 'Add New MCP Server'. "
                "Set Type to 'command' and paste the command line above."
            ),
        },
    }
