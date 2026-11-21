# `mcp_monitor` Package

Modular implementation of `McpMonitorDialog`, decomposed cleanly from `src/core/ui/mcp_monitor_dialog.py`.

## Structure & Mixin Composition (MRO)

- `presets.py`: Default example tool arguments (`DEFAULT_PLAYGROUND_ARGUMENTS`).
- `dialog.py` (`McpMonitorDialog`): Non-modal top-level window shell (`super().__init__(None)`, `Qt.Window`), top system/metrics banner, 3-tab shell, and 2-second polling timer, inheriting from three focused mixins ahead of `QDialog`:
  1. `McpActivityStreamMixin` (`mixins/activity_stream.py`): Tab 1 activity table, search/tool/status filters, JSON args/response inspector, and live telemetry refresh.
  2. `McpClientSetupMixin` (`mixins/client_setup.py`): Tab 2 installed LLM CLI status strip, client snippet tabs, and clipboard feedback helpers.
  3. `McpToolPlaygroundMixin` (`mixins/tool_playground.py`): Tab 3 interactive tool/action selector and direct `ToolRegistry` call runner.

## Backward Compatibility

`src/core/ui/mcp_monitor_dialog.py` acts as a thin re-export shim exposing `McpMonitorDialog`, `DEFAULT_PLAYGROUND_ARGUMENTS`, and `logger`.
