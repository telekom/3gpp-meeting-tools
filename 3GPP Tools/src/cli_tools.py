# --- File: src/cli_tools.py ---
"""
Headless Command-Line Interface (CLI) for 3GPP Tools.

Allows external CLI agents (Claude CLI, Codex CLI, shell scripts) or human
engineers to list, inspect, and invoke any of the 12 domain tools directly
from PowerShell/Bash with structured JSON output.

Usage Examples:
    python -m cli_tools list
    python -m cli_tools list --profile onsite
    python -m cli_tools schema onsite_meeting_cockpit
    python -m cli_tools call onsite_meeting_cockpit --action network_status
    python -m cli_tools call spec_text_and_clauses --action search_text --arg query="PEI" --arg spec_number=23501
    python -m cli_tools call protocol_inspector --json "{\"action\":\"list_protocols\"}"
    python -m cli_tools onsite-status --wg SA2
"""

from __future__ import annotations

import argparse
import faulthandler
import json
import logging
import math
import os
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

_SRC_DIR = Path(__file__).resolve().parent
_ROOT_DIR = _SRC_DIR.parent
for _p in (str(_ROOT_DIR), str(_SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.utils.paths import ensure_app_data_dirs, get_project_root
from core.ai.mcp_activity import record_mcp_event
from modules.assistant.core.tool_registry import ArgumentType
from modules.assistant.core.tools.external_agent_tools import (
    PROFILE_TOOLS,
    build_external_agent_tool_registry,
)

logger = logging.getLogger(__name__)


def _configure_cli_logging(log_level: str = "WARNING") -> None:
    try:
        faulthandler.disable()
    except Exception:
        pass

    log_file_path = get_project_root() / "3gpp_tools.log"
    level = getattr(logging, str(log_level).upper(), logging.WARNING)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] [CLI:%(name)s] %(message)s")
    try:
        fh = logging.FileHandler(log_file_path, mode="a", encoding="utf-8")
        fh.setFormatter(formatter)
        root_logger.addHandler(fh)
    except Exception:
        pass

    sh = logging.StreamHandler(sys.__stderr__)
    sh.setFormatter(formatter)
    root_logger.addHandler(sh)


def _sanitize_for_json(obj: Any) -> Any:
    """Recursively replaces non-finite floats (NaN, Infinity) with None for RFC 8259 compliance."""
    if isinstance(obj, float):
        if not math.isfinite(obj):
            return None
        return obj
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_for_json(v) for v in obj]
    return obj


def _safe_json_dumps(obj: Any, **kwargs: Any) -> str:
    """Strict JSON serializer that guarantees standard compliance without NaN or Infinity."""
    sanitized = _sanitize_for_json(obj)
    kwargs.setdefault("allow_nan", False)
    return json.dumps(sanitized, **kwargs)


def _coerce_arg_value(raw: str, expected_type: Optional[ArgumentType] = None) -> Any:
    """Coerce CLI `key=value` string into bool, int, float, JSON array/object, or string, guided by expected_type."""
    trimmed = raw.strip()
    if expected_type == ArgumentType.STRING:
        return trimmed
    if expected_type == ArgumentType.INTEGER:
        try:
            return int(trimmed)
        except Exception:
            return trimmed
    if expected_type == ArgumentType.BOOLEAN:
        if trimmed.lower() in ("true", "1", "yes"):
            return True
        if trimmed.lower() in ("false", "0", "no"):
            return False
        return trimmed

    if trimmed.lower() == "true":
        return True
    if trimmed.lower() == "false":
        return False
    if trimmed.lower() in ("null", "none"):
        return None
    if (trimmed.startswith("[") and trimmed.endswith("]")) or (
        trimmed.startswith("{") and trimmed.endswith("}")
    ):
        try:
            return json.loads(trimmed)
        except Exception:
            pass
    # Keep 5-digit spec numbers like "23501" or dotted clauses like "5.5.1" as strings if needed,
    # or parse ints for limit/max_results while handlers safely cast via str() or int().
    try:
        if trimmed.isdigit() and not (len(trimmed) in (4, 5) and trimmed.startswith(("2", "3"))):
            return int(trimmed)
    except Exception:
        pass
    return raw


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="cli_tools",
        description="Headless CLI for 3GPP Tools (12 domain tools for specs, protocols, on-site 10.10.10.10 LAN, TDocs, and Word/PlantUML authoring).",
    )
    parser.add_argument(
        "--profile",
        choices=sorted(PROFILE_TOOLS.keys()),
        default="all",
        help="Tool profile filter ('all', 'knowledge', 'onsite', 'authoring').",
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Disable document-writing and database-mutating actions.",
    )
    parser.add_argument(
        "--log-level",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Log level emitted to stderr and 3gpp_tools.log.",
    )

    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    # list
    sub_list = subparsers.add_parser("list", help="List available tools and their supported actions.")
    sub_list.add_argument(
        "--json",
        action="store_true",
        dest="as_json",
        help="Output full MCP tool schemas as JSON.",
    )

    # schema
    sub_schema = subparsers.add_parser("schema", help="Print the JSON inputSchema for a specific tool.")
    sub_schema.add_argument("tool_name", help="Name of the tool (e.g. onsite_meeting_cockpit).")

    # call
    sub_call = subparsers.add_parser("call", help="Invoke a tool action and print JSON result.")
    sub_call.add_argument("tool_name", help="Name of the tool to invoke.")
    sub_call.add_argument("--action", help="Action name (e.g. search_text, network_status).")
    sub_call.add_argument(
        "--json",
        "--args",
        dest="json_payload",
        help="JSON object string of arguments (e.g. '{\"action\":\"network_status\"}').",
    )
    sub_call.add_argument(
        "--arg",
        action="append",
        default=[],
        help="Key=Value argument pair (repeatable, e.g. --arg spec_number=23501 --arg query=PEI).",
    )

    # onsite-status shortcut
    sub_onsite = subparsers.add_parser(
        "onsite-status",
        help="Quick on-site meeting status (10.10.10.10 LAN reachability + active session streams).",
    )
    sub_onsite.add_argument("--wg", default="SA2", help="Working group (default: SA2).")
    sub_onsite.add_argument("--meeting-id", default=None, help="Optional meeting ID.")

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    _configure_cli_logging(args.log_level)

    # Ensure UTF-8 output on Windows consoles so emoji badges (🎯, ⚪, etc.) never raise UnicodeEncodeError
    real_stdout = sys.__stdout__
    for stream in (real_stdout, sys.__stderr__):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    sys.stdout = sys.stderr

    try:
        ensure_app_data_dirs()
    except Exception:
        pass

    registry = build_external_agent_tool_registry(
        profile=args.profile,
        read_only=args.read_only,
    )

    if args.subcommand == "list":
        mcp_tools = registry.mcp_tools()
        if args.as_json:
            real_stdout.write(json.dumps({"profile": args.profile, "tools": mcp_tools}, indent=2) + "\n")
        else:
            summary = []
            for t in mcp_tools:
                actions = (
                    t.get("inputSchema", {})
                    .get("properties", {})
                    .get("action", {})
                    .get("enum", [])
                )
                summary.append(
                    {
                        "name": t["name"],
                        "description": t["description"],
                        "actions": actions,
                    }
                )
            real_stdout.write(
                json.dumps(
                    {
                        "profile": args.profile,
                        "tool_count": len(summary),
                        "tools": summary,
                    },
                    indent=2,
                )
                + "\n"
            )
        real_stdout.flush()
        return 0

    if args.subcommand == "schema":
        defn = registry.get(args.tool_name)
        if defn is None:
            real_stdout.write(
                json.dumps(
                    {"ok": False, "error": f"Unknown tool: {args.tool_name}"},
                    indent=2,
                )
                + "\n"
            )
            real_stdout.flush()
            return 1
        real_stdout.write(json.dumps(defn.mcp_schema(), indent=2) + "\n")
        real_stdout.flush()
        return 0

    if args.subcommand == "call":
        call_args: Dict[str, Any] = {}
        if args.json_payload:
            try:
                parsed = json.loads(args.json_payload)
                if not isinstance(parsed, dict):
                    raise ValueError("--json payload must be a JSON object")
                call_args.update(parsed)
            except Exception as exc:
                real_stdout.write(
                    json.dumps({"ok": False, "error": f"Invalid --json payload: {exc}"}, indent=2) + "\n"
                )
                real_stdout.flush()
                return 1

        if args.action:
            call_args["action"] = args.action

        tool_defn = registry.get(args.tool_name)
        type_map = {a.name: a.type for a in tool_defn.arguments} if tool_defn else {}

        for kv in args.arg or []:
            if "=" not in kv:
                real_stdout.write(
                    json.dumps({"ok": False, "error": f"Invalid --arg '{kv}' (expected key=value)"}, indent=2)
                    + "\n"
                )
                real_stdout.flush()
                return 1
            k, v = kv.split("=", 1)
            clean_k = k.strip()
            call_args[clean_k] = _coerce_arg_value(v, type_map.get(clean_k))

        t0 = time.perf_counter()
        res = registry.execute(args.tool_name, call_args)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        if res.ok:
            is_err = isinstance(res.data, dict) and res.data.get("ok") is False
            try:
                record_mcp_event(
                    "cli/call",
                    client="cli",
                    client_ver="1.0.0",
                    pid=os.getpid(),
                    tool=args.tool_name,
                    action=call_args.get("action", ""),
                    arguments=call_args,
                    elapsed_ms=elapsed_ms,
                    ok=True,
                    is_error=bool(is_err),
                    raw_data=res.data,
                )
            except Exception:
                pass
            real_stdout.write(_safe_json_dumps(res.data, ensure_ascii=False, indent=2, default=str) + "\n")
            real_stdout.flush()
            return 1 if is_err else 0
        else:
            err_msg = getattr(res, "error", None) or getattr(res, "message", "Tool execution failed")
            try:
                record_mcp_event(
                    "cli/call",
                    client="cli",
                    client_ver="1.0.0",
                    pid=os.getpid(),
                    tool=args.tool_name,
                    action=call_args.get("action", ""),
                    arguments=call_args,
                    elapsed_ms=elapsed_ms,
                    ok=False,
                    is_error=True,
                    error_msg=err_msg,
                )
            except Exception:
                pass
            real_stdout.write(
                _safe_json_dumps({"ok": False, "tool": args.tool_name, "error": err_msg}, indent=2) + "\n"
            )
            real_stdout.flush()
            return 1

    if args.subcommand == "onsite-status":
        net_res = registry.execute(
            "onsite_meeting_cockpit",
            {"action": "network_status", "wg": args.wg, "force_probe": True},
        )
        now_res = registry.execute(
            "onsite_meeting_cockpit",
            {"action": "where_should_we_go_now", "wg": args.wg},
        )
        combined = {
            "ok": True,
            "network_status": net_res.data if net_res.ok else {"error": getattr(net_res, "message", "failed")},
            "where_should_we_go_now": now_res.data if now_res.ok else {"error": getattr(now_res, "message", "failed")},
        }
        real_stdout.write(_safe_json_dumps(combined, ensure_ascii=False, indent=2, default=str) + "\n")
        real_stdout.flush()
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
