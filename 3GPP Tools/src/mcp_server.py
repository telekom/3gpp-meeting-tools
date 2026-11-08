# --- File: src/mcp_server.py ---
"""
Model Context Protocol (MCP) Server for 3GPP Tools.

Exposes the 12 compact domain tools of 3GPP Tools to external coding/standards
agents such as Claude CLI (`claude`) and OpenAI Codex CLI (`codex`) over
standard input/output (`stdio`).

Features:
- Dual JSON-RPC 2.0 transport support:
  * Newline-delimited JSON-RPC (used by MCP 2024-11-05 / 2025-03-26, Claude CLI, Codex CLI)
  * `Content-Length:` HTTP-style header framing (LSP / legacy MCP clients)
- Token-conscious profile filtering (`--profile all|knowledge|onsite|authoring`)
  and `--read-only` safety mode.
- Strict stdout isolation: redirects Python `sys.stdout` to `sys.stderr` so any
  stray `print()` or third-party library output never corrupts the JSON-RPC
  stream, while logging structured execution traces to `3gpp_tools.log` and `stderr`.
"""

from __future__ import annotations

import argparse
import faulthandler
import json
import logging
import sys
import time
from pathlib import Path
from typing import Any, BinaryIO, Dict, Optional

# Ensure `src/` is on sys.path when invoked directly via `python src/mcp_server.py`
_SRC_DIR = Path(__file__).resolve().parent
_ROOT_DIR = _SRC_DIR.parent
for _p in (str(_ROOT_DIR), str(_SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.utils.paths import ensure_app_data_dirs, get_project_root
from modules.assistant.core.tool_registry import ToolRegistry
from modules.assistant.core.tools.external_agent_tools import (
    PROFILE_TOOLS,
    build_external_agent_tool_registry,
)

logger = logging.getLogger(__name__)

SERVER_NAME = "3gpp-tools"
SERVER_VERSION = "1.0.0"
DEFAULT_PROTOCOL_VERSION = "2024-11-05"
SUPPORTED_PROTOCOL_VERSIONS = {"2024-11-05", "2025-03-26", "2024-10-07"}


def configure_mcp_logging(log_level: str = "INFO") -> Path:
    """Configure file + stderr logging while keeping stdout strictly clean for MCP JSON-RPC."""
    try:
        faulthandler.disable()
    except Exception:
        pass

    log_file_path = get_project_root() / "3gpp_tools.log"
    level = getattr(logging, str(log_level).upper(), logging.INFO)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] [MCP:%(name)s] %(message)s")

    try:
        file_handler = logging.FileHandler(log_file_path, mode="a", encoding="utf-8")
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
    except Exception:
        pass

    stderr_handler = logging.StreamHandler(sys.__stderr__)
    stderr_handler.setFormatter(formatter)
    root_logger.addHandler(stderr_handler)

    return log_file_path


class MCPServer:
    """JSON-RPC 2.0 MCP Server backed by a `ToolRegistry`."""

    def __init__(
        self,
        registry: ToolRegistry,
        *,
        profile: str = "all",
        read_only: bool = False,
    ) -> None:
        self.registry = registry
        self.profile = profile
        self.read_only = read_only

    def handle_message(self, message: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Process a single JSON-RPC 2.0 request or notification."""
        if not isinstance(message, dict):
            return self._error_response(None, -32600, "Invalid Request: expected JSON object")

        method = message.get("method")
        msg_id = message.get("id")
        is_notification = "id" not in message

        if not method or not isinstance(method, str):
            if is_notification:
                return None
            return self._error_response(msg_id, -32600, "Invalid Request: missing 'method'")

        params = message.get("params") or {}
        if not isinstance(params, dict):
            params = {}

        logger.info("📥 Received MCP method='%s' id=%s", method, msg_id)

        # Notifications do not produce a response
        if is_notification or method.startswith("notifications/"):
            return None

        try:
            if method == "initialize":
                return self._handle_initialize(msg_id, params)
            if method == "ping":
                return self._result_response(msg_id, {})
            if method == "tools/list":
                return self._handle_tools_list(msg_id, params)
            if method == "tools/call":
                return self._handle_tools_call(msg_id, params)
            if method == "resources/list":
                return self._result_response(msg_id, {"resources": []})
            if method == "resources/templates/list":
                return self._result_response(msg_id, {"resourceTemplates": []})
            if method == "prompts/list":
                return self._result_response(msg_id, {"prompts": []})
            if method == "logging/setLevel":
                lvl = str(params.get("level") or "INFO").upper()
                logging.getLogger().setLevel(getattr(logging, lvl, logging.INFO))
                return self._result_response(msg_id, {})

            return self._error_response(msg_id, -32601, f"Method not found: {method}")
        except Exception as exc:
            logger.exception("❌ Unhandled exception in MCP method '%s': %s", method, exc)
            return self._error_response(msg_id, -32603, f"Internal error: {exc}")

    def _handle_initialize(self, msg_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        client_version = str(params.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION)
        negotiated_version = (
            client_version
            if client_version in SUPPORTED_PROTOCOL_VERSIONS
            else DEFAULT_PROTOCOL_VERSION
        )
        client_info = params.get("clientInfo") or {}
        logger.info(
            "🤝 MCP Initialize from client=%s version=%s (negotiated=%s, profile=%s, read_only=%s)",
            client_info.get("name", "unknown"),
            client_info.get("version", "?"),
            negotiated_version,
            self.profile,
            self.read_only,
        )
        return self._result_response(
            msg_id,
            {
                "protocolVersion": negotiated_version,
                "capabilities": {
                    "tools": {"listChanged": False},
                },
                "serverInfo": {
                    "name": SERVER_NAME,
                    "version": SERVER_VERSION,
                },
                "instructions": (
                    "3GPP Tools MCP Server: provides offline 3GPP specification search, clause evolution & "
                    "CR attribution, NAS/RRC/NGAP protocol inspection, on-site 10.10.10.10 meeting LAN cockpit, "
                    "TDoc & Session Plan tracking, Word/Markdown/CR authoring, and PlantUML-to-Visio/PPTX tools."
                ),
            },
        )

    def _handle_tools_list(self, msg_id: Any, _params: Dict[str, Any]) -> Dict[str, Any]:
        tools = self.registry.mcp_tools()
        logger.info("📋 Returning %d MCP tool schemas (profile='%s')", len(tools), self.profile)
        return self._result_response(msg_id, {"tools": tools})

    def _handle_tools_call(self, msg_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        tool_name = str(params.get("name") or "").strip()
        arguments = params.get("arguments") or {}
        if not isinstance(arguments, dict):
            return self._error_response(msg_id, -32602, "Invalid params: 'arguments' must be an object")

        if not tool_name:
            return self._error_response(msg_id, -32602, "Invalid params: missing tool 'name'")

        action = arguments.get("action", "")
        t0 = time.perf_counter()
        logger.info("🔧 Executing tool='%s' action='%s' args=%s", tool_name, action, list(arguments.keys()))

        exec_result = self.registry.execute(tool_name, arguments)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        if exec_result.ok:
            payload_text = json.dumps(exec_result.data, ensure_ascii=False, indent=2, default=str)
            is_logic_error = isinstance(exec_result.data, dict) and exec_result.data.get("ok") is False
            logger.info(
                "✅ Tool '%s' (action='%s') finished in %.1f ms (logic_error=%s)",
                tool_name,
                action,
                elapsed_ms,
                is_logic_error,
            )
            return self._result_response(
                msg_id,
                {
                    "content": [{"type": "text", "text": payload_text}],
                    "isError": bool(is_logic_error),
                },
            )
        else:
            err_msg = getattr(exec_result, "error", None) or getattr(exec_result, "message", "Tool execution failed")
            err_payload = json.dumps(
                {"ok": False, "tool": tool_name, "error": err_msg},
                ensure_ascii=False,
                indent=2,
            )
            logger.warning(
                "⚠️ Tool '%s' (action='%s') failed in %.1f ms: %s",
                tool_name,
                action,
                elapsed_ms,
                err_msg,
            )
            return self._result_response(
                msg_id,
                {
                    "content": [{"type": "text", "text": err_payload}],
                    "isError": True,
                },
            )

    @staticmethod
    def _result_response(msg_id: Any, result: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": result,
        }

    @staticmethod
    def _error_response(msg_id: Any, code: int, message: str, data: Any = None) -> Dict[str, Any]:
        err: Dict[str, Any] = {"code": code, "message": message}
        if data is not None:
            err["data"] = data
        return {
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": err,
        }


def _read_next_rpc_message(stream: BinaryIO) -> tuple[Optional[Dict[str, Any]], str]:
    """
    Read a single JSON-RPC message from a binary stream.

    Automatically detects:
    - `"jsonl"`: Newline-delimited JSON object (`{...}\\n`)
    - `"content-length"`: HTTP-header framed (`Content-Length: N\\r\\n\\r\\n{...}`)

    Returns `(parsed_dict_or_None_on_EOF, framing_mode)`.
    """
    while True:
        first_line = stream.readline()
        if not first_line:
            return None, "jsonl"

        stripped = first_line.strip()
        if not stripped:
            continue

        lower = stripped.lower()
        if lower.startswith(b"content-length:"):
            length_str = stripped.split(b":", 1)[1].strip()
            content_length = int(length_str)
            # Consume remaining headers up to blank line
            while True:
                header_line = stream.readline()
                if not header_line or header_line in (b"\r\n", b"\n"):
                    break
            body = b""
            while len(body) < content_length:
                chunk = stream.read(content_length - len(body))
                if not chunk:
                    break
                body += chunk
            if not body:
                return None, "content-length"
            return json.loads(body.decode("utf-8")), "content-length"

        # Otherwise treat as newline-delimited JSON
        return json.loads(stripped.decode("utf-8")), "jsonl"


def _write_rpc_message(stream: BinaryIO, response: Dict[str, Any], framing: str = "jsonl") -> None:
    payload = json.dumps(response, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if framing == "content-length":
        header = f"Content-Length: {len(payload)}\r\n\r\n".encode("ascii")
        stream.write(header)
        stream.write(payload)
    else:
        stream.write(payload + b"\n")
    stream.flush()


def serve_stdio(
    server: MCPServer,
    in_stream: Optional[BinaryIO] = None,
    out_stream: Optional[BinaryIO] = None,
) -> None:
    """Run the synchronous stdio loop for the MCP server."""
    reader = in_stream or sys.stdin.buffer
    writer = out_stream or sys.__stdout__.buffer

    logger.info(
        "🚀 3GPP Tools MCP Server started on stdio (profile='%s', read_only=%s, tools=%d)",
        server.profile,
        server.read_only,
        len(server.registry.definitions()),
    )

    while True:
        try:
            msg, framing = _read_next_rpc_message(reader)
            if msg is None:
                logger.info("👋 EOF on stdin; shutting down MCP Server cleanly.")
                break
        except json.JSONDecodeError as exc:
            logger.error("❌ JSON parse error on stdin: %s", exc)
            err_resp = MCPServer._error_response(None, -32700, f"Parse error: {exc}")
            _write_rpc_message(writer, err_resp, framing="jsonl")
            continue
        except Exception as exc:
            logger.exception("❌ Unexpected error reading from stdin: %s", exc)
            break

        response = server.handle_message(msg)
        if response is not None:
            _write_rpc_message(writer, response, framing=framing)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mcp_server",
        description="3GPP Tools Model Context Protocol (MCP) Server for Claude CLI and Codex CLI.",
    )
    parser.add_argument(
        "--profile",
        choices=sorted(PROFILE_TOOLS.keys()),
        default="all",
        help=(
            "Select tool subset to minimize system-prompt token overhead: "
            "'all' (12 tools, ~1.8k tokens), 'knowledge' (specs/clauses/protocols/history), "
            "'onsite' (10.10.10.10 LAN cockpit + TDocs + Session Plan + WIs), "
            "'authoring' (Word/Markdown/CR/pCR/Sanitizer/PlantUML/Visio)."
        ),
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Disable document-writing and state-mutating actions.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity written to 3gpp_tools.log and stderr.",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    configure_mcp_logging(args.log_level)

    # Redirect global `sys.stdout` to `sys.stderr` so any accidental `print()` inside
    # libraries or worker modules goes to stderr instead of corrupting MCP JSON-RPC.
    real_stdout_buffer = sys.__stdout__.buffer
    sys.stdout = sys.stderr

    try:
        ensure_app_data_dirs()
    except Exception as exc:
        logger.warning("⚠️ Could not initialize app data dirs: %s", exc)

    registry = build_external_agent_tool_registry(
        profile=args.profile,
        read_only=args.read_only,
    )
    server = MCPServer(
        registry,
        profile=args.profile,
        read_only=args.read_only,
    )
    serve_stdio(server, in_stream=sys.stdin.buffer, out_stream=real_stdout_buffer)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
