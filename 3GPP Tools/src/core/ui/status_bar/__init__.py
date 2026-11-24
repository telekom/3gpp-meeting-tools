# --- File: src/core/ui/status_bar/__init__.py ---
"""Modular SystemStatusBar package."""

from .bar import SystemStatusBar
from .styles import _MINIMAL_HYSTERESIS_PX, _status_pill_qss

__all__ = [
    "_MINIMAL_HYSTERESIS_PX",
    "_status_pill_qss",
    "SystemStatusBar",
]
