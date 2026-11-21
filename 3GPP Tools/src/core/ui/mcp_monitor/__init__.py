"""MCP Server & External AI Agent Monitor Cockpit package."""

from core.ui.mcp_monitor.dialog import McpMonitorDialog, logger
from core.ui.mcp_monitor.presets import DEFAULT_PLAYGROUND_ARGUMENTS

__all__ = [
    "McpMonitorDialog",
    "DEFAULT_PLAYGROUND_ARGUMENTS",
    "logger",
]
