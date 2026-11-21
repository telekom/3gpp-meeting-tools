"""Backward-compatible re-export shim for McpMonitorDialog."""

from core.ui.mcp_monitor import (
    DEFAULT_PLAYGROUND_ARGUMENTS,
    McpMonitorDialog,
    logger as _dialog_logger,
)

logger = _dialog_logger

__all__ = [
    "McpMonitorDialog",
    "DEFAULT_PLAYGROUND_ARGUMENTS",
]
