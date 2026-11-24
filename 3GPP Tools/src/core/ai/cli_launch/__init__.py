"""Modular external LLM CLI context writer, terminal launcher, and Qt menu builder."""

from core.ai.cli_detector import detect_cli_agents
from core.ai.cli_launch.context_writer import (
    CLI_CONTEXT_FILENAME,
    DEFAULT_BOOTSTRAP_PROMPT,
    STANDING_CONTEXT_PATH,
    USER_NOTES_FILENAME,
    _BASE_TDOC_RE,
    _build_task_instruction,
    _resolve_cli_tools_script,
    collect_per_tdoc_user_notes,
    load_standing_cli_context,
    load_workspace_user_notes,
    logger,
    save_standing_cli_context,
    save_workspace_user_notes,
    write_cli_context_file,
)
from core.ai.cli_launch.menu_builder import (
    add_model_selector_submenu,
    open_mcp_agent_setup_dialog,
    populate_cli_agent_menu,
)
from core.ai.cli_launch.terminal_launch import (
    _ps_quote,
    _show_critical,
    _show_warning,
    build_powershell_launch_command,
    encode_powershell_command,
    launch_cli_in_terminal,
)

__all__ = [
    "CLI_CONTEXT_FILENAME",
    "DEFAULT_BOOTSTRAP_PROMPT",
    "STANDING_CONTEXT_PATH",
    "USER_NOTES_FILENAME",
    "_BASE_TDOC_RE",
    "_build_task_instruction",
    "_ps_quote",
    "_resolve_cli_tools_script",
    "_show_critical",
    "_show_warning",
    "add_model_selector_submenu",
    "build_powershell_launch_command",
    "collect_per_tdoc_user_notes",
    "detect_cli_agents",
    "encode_powershell_command",
    "launch_cli_in_terminal",
    "load_standing_cli_context",
    "load_workspace_user_notes",
    "logger",
    "open_mcp_agent_setup_dialog",
    "populate_cli_agent_menu",
    "save_standing_cli_context",
    "save_workspace_user_notes",
    "write_cli_context_file",
]
