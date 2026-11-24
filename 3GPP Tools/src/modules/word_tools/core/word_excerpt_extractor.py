# --- File: src/modules/word_tools/core/word_excerpt_extractor.py ---
"""Backward-compatible re-export shim for ``modules.word_tools.core.word_excerpt``."""

from __future__ import annotations

import logging

from .word_excerpt import (
    NAMING_SCHEMES,
    SCHEME_SOURCE_CLAUSES_TIMESTAMP,
    HeadingInfo,
    WordExcerptEngine,
    WordExcerptThread,
    WordHeadingScanner,
    WordHeadingScanThread,
    ensure_3gpp_package_styles,
    generate_differentiated_export_name,
    natural_key,
    order_sources,
    set_sensitivity_label,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "HeadingInfo",
    "natural_key",
    "order_sources",
    "WordHeadingScanner",
    "WordExcerptEngine",
    "WordExcerptThread",
    "WordHeadingScanThread",
    "ensure_3gpp_package_styles",
    "set_sensitivity_label",
    "NAMING_SCHEMES",
    "SCHEME_SOURCE_CLAUSES_TIMESTAMP",
    "generate_differentiated_export_name",
]
