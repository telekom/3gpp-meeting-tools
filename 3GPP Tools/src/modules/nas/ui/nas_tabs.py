"""Backward-compatible re-export shim for NASTab and protocol UI constants.

The implementation is organized under ``modules.nas.ui.nas_tab``.
"""

from __future__ import annotations

from modules.nas.ui.nas_workers import NASMessageListWorker, ReverseLookupWorker

from .nas_tab import (
    RE_CLAUSE_FROM_REF,
    RE_CTX_LCH,
    RE_CTX_RLC,
    RE_CTX_SRB,
    RE_CTX_TRAILING_CAPTION,
    RE_SPEC_FROM_COL,
    RIBBON_BADGE_STYLE,
    SEGMENTED_MODE_STYLE,
    TINY_HEADER_BTN_STYLE,
    TOGGLE_BUTTON_STYLE,
    NASTab,
)

__all__ = [
    "NASMessageListWorker",
    "NASTab",
    "RE_CLAUSE_FROM_REF",
    "RE_CTX_LCH",
    "RE_CTX_RLC",
    "RE_CTX_SRB",
    "RE_CTX_TRAILING_CAPTION",
    "RE_SPEC_FROM_COL",
    "RIBBON_BADGE_STYLE",
    "ReverseLookupWorker",
    "SEGMENTED_MODE_STYLE",
    "TINY_HEADER_BTN_STYLE",
    "TOGGLE_BUTTON_STYLE",
]
