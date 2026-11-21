"""Backward-compatible re-export shim for SpecificationsTab and WG constants.

The implementation is organized under ``modules.specifications.ui.specs_tab``.
"""

from __future__ import annotations

from core.ui.ui_components import get_wg_badge_style_and_desc

from .specs_tab import (
    ACTIVE_WG_ORDER,
    FULL_WG_ORDER,
    LEGACY_CT,
    LEGACY_OTHER,
    LEGACY_RAN,
    LEGACY_WGS,
    PART_NUM_REGEX,
    USEFUL_CT,
    USEFUL_RAN,
    USEFUL_SA,
    USEFUL_WGS,
    SpecificationsTab,
)

__all__ = [
    "ACTIVE_WG_ORDER",
    "FULL_WG_ORDER",
    "LEGACY_CT",
    "LEGACY_OTHER",
    "LEGACY_RAN",
    "LEGACY_WGS",
    "PART_NUM_REGEX",
    "SpecificationsTab",
    "USEFUL_CT",
    "USEFUL_RAN",
    "USEFUL_SA",
    "USEFUL_WGS",
    "get_wg_badge_style_and_desc",
]
