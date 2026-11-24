# --- File: src/core/ui/system_status_bar.py ---
"""Backward-compatible re-export shim for the modular SystemStatusBar package."""

import logging

from core.ui.status_bar import _MINIMAL_HYSTERESIS_PX, SystemStatusBar, _status_pill_qss

logger = logging.getLogger(__name__)

__all__ = [
    "_MINIMAL_HYSTERESIS_PX",
    "_status_pill_qss",
    "SystemStatusBar",
]
