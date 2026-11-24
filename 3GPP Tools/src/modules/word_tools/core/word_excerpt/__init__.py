"""Modular Word heading scanner and clause excerpt extraction package."""

from __future__ import annotations

from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import ensure_3gpp_package_styles
from modules.word_tools.core.word_excerpt_naming import (
    NAMING_SCHEMES,
    SCHEME_SOURCE_CLAUSES_TIMESTAMP,
    generate_differentiated_export_name,
)
from .engine import WordExcerptEngine
from .scanner import HeadingInfo, WordHeadingScanner, natural_key, order_sources
from .threads import WordExcerptThread, WordHeadingScanThread

__all__ = [
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
