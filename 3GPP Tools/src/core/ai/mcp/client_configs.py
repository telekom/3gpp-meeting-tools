"""Client configuration snippet generators for external MCP-enabled AI tools."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict

from core.utils.paths import get_project_root


def get_mcp_client_configurations() -> Dict[str, Dict[str, Any]]:
    """
    Generate ready-to-use configuration blocks with auto-detected absolute
    paths for Claude Desktop, Codex CLI, Claude CLI, Gemini CLI, Cursor, and Windsurf.
    """
    root_dir = get_project_root()
    src_dir = root_dir if (root_dir / "mcp_server.py").exists() else (root_dir / "src")
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

    # 2. Codex CLI configuration (~/.codex/config.toml)
    codex_config_file = Path.home() / ".codex" / "config.toml"
    codex_toml = (
        "[mcp_servers.3gpp-tools]\n"
        f'command = {json.dumps(python_exe_str)}\n'
        f'args = [{json.dumps(server_script_str)}, "--profile", "all"]\n'
    )

    # 3. Claude CLI command snippet
    claude_cli_cmd = (
        f'claude mcp add 3gpp-tools -- "{python_exe_str}" "{server_script_str}" --profile all'
    )

    # 3b. Google Gemini CLI configuration (~/.gemini/settings.json)
    gemini_config_file = Path.home() / ".gemini" / "settings.json"
    gemini_cfg = {
        "mcpServers": {
            "3gpp-tools": {
                "command": python_exe_str,
                "args": [server_script_str, "--profile", "all"],
            }
        }
    }
    gemini_cli_cmd = (
        f'gemini mcp add 3gpp-tools "{python_exe_str}" "{server_script_str}" -- --profile all'
    )

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
    claude_config_file = (
        Path.home() / "AppData" / "Roaming" / "Claude" / "claude_desktop_config.json"
    )
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
            "target_file": str(codex_config_file),
            "config_toml": codex_toml,
            "config_json": codex_toml,
            "cli_command": f'codex mcp add 3gpp-tools -- "{python_exe_str}" "{server_script_str}" --profile all',
            "instructions": (
                "Run the CLI command above or add this [mcp_servers.3gpp-tools] block to ~/.codex/config.toml. "
                "Codex will automatically discover all 12 domain tools."
            ),
        },
        "claude_cli": {
            "title": "Anthropic Claude CLI",
            "cli_command": claude_cli_cmd,
            "instructions": (
                "Run this command in PowerShell or cmd to permanently register 3GPP Tools with Claude CLI."
            ),
        },
        "gemini_cli": {
            "title": "Google Gemini CLI",
            "target_file": str(gemini_config_file),
            "config_json": json.dumps(gemini_cfg, indent=2),
            "cli_command": gemini_cli_cmd,
            "instructions": (
                "Run the CLI command above or add this entry to ~/.gemini/settings.json under 'mcpServers'."
            ),
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
