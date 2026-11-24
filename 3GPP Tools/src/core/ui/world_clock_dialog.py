"""Backward-compatible re-export shim for WorldClockDialog.

Canonical implementation lives in ``core.ui.world_clock``.
"""

from __future__ import annotations

import logging

from core.ui.world_clock import (
    WorldClockDialog,
    _VenueClockCard,
    _badge_style,
    _build_utc_offset_options,
    _format_delta_vs_laptop,
    _source_badge_spec,
)

logger = logging.getLogger(__name__)

__all__ = [
    "WorldClockDialog",
    "_VenueClockCard",
    "_badge_style",
    "_build_utc_offset_options",
    "_format_delta_vs_laptop",
    "_source_badge_spec",
    "logger",
]
