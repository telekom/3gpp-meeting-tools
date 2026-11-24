# System Status Bar Package (`core/ui/status_bar`)

Modular status bar widget managing interactive pills (App Update, Ollama, MCP, Proxy, Wi-Fi/3GPP LAN, System menu, Right-panel dock toggles, Console & Queue drawer, World Clock) with content-driven Minimalist Mode hysteresis.

## Module Inventory

| File | Responsibility |
|---|---|
| `styles.py` | `_MINIMAL_HYSTERESIS_PX` and `_status_pill_qss()` |
| `mixins/responsive.py` | `StatusBarResponsiveMixin` — full-width estimation, `resizeEvent`, hysteresis minimalist-mode toggling |
| `mixins/ui_setup.py` | `StatusBarUiSetupMixin` — pill initialization, System menu, update badge, drawer & right-panel badges |
| `mixins/network_and_proxy.py` | `StatusBarNetworkProxyMixin` — `os.environ`/urllib proxy sync, `WifiMonitorThread`, network indicator |
| `mixins/ai_and_clock.py` | `StatusBarAiClockMixin` — `OllamaMonitorThread`, MCP polling, World Clock + background venue sync, `stop_monitors()` |
| `bar.py` | Composed `SystemStatusBar` `QStatusBar` class |

`core/ui/system_status_bar.py` remains as a thin backward-compatible re-export shim.
