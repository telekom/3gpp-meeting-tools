"""Modular Specifications Catalogue tab package."""

from __future__ import annotations

from .constants import (
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
)
from .tab import SpecificationsTab

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
]
