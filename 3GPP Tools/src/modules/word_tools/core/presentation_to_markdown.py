"""Backward-compatible re-export shim for PowerPoint/PDF-to-Markdown conversion.

Canonical implementation lives in ``modules.word_tools.core.presentation``.
"""

from __future__ import annotations

import logging

from modules.word_tools.core.presentation import (
    CONVERTER_VERSION,
    OPENXML_NS,
    TDOC_RE,
    ExtractedSlide,
    _build_presentation_header,
    _clean_tdoc_prefix,
    _extract_via_openxml,
    _extract_via_pdf,
    _extract_via_powerpoint_com,
    _format_markdown_table,
    _safe_log,
    convert_presentation_to_markdown,
)

logger = logging.getLogger(__name__)

__all__ = [
    "CONVERTER_VERSION",
    "ExtractedSlide",
    "OPENXML_NS",
    "TDOC_RE",
    "_build_presentation_header",
    "_clean_tdoc_prefix",
    "_extract_via_openxml",
    "_extract_via_pdf",
    "_extract_via_powerpoint_com",
    "_format_markdown_table",
    "_safe_log",
    "convert_presentation_to_markdown",
    "logger",
]
