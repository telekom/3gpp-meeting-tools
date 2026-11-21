"""Tab mixins composing McpMonitorDialog."""

from core.ui.mcp_monitor.mixins.activity_stream import McpActivityStreamMixin
from core.ui.mcp_monitor.mixins.client_setup import McpClientSetupMixin
from core.ui.mcp_monitor.mixins.tool_playground import McpToolPlaygroundMixin

__all__ = [
    "McpActivityStreamMixin",
    "McpClientSetupMixin",
    "McpToolPlaygroundMixin",
]
