"""Modular PowerPoint / PDF presentation-to-Markdown package."""

from modules.word_tools.core.presentation.com_extractor import _extract_via_powerpoint_com
from modules.word_tools.core.presentation.converter import (
    _build_presentation_header,
    convert_presentation_to_markdown,
)
from modules.word_tools.core.presentation.models import (
    CONVERTER_VERSION,
    OPENXML_NS,
    TDOC_RE,
    ExtractedSlide,
    _clean_tdoc_prefix,
    _format_markdown_table,
    _safe_log,
)
from modules.word_tools.core.presentation.openxml_and_pdf import (
    _extract_via_openxml,
    _extract_via_pdf,
)

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
]
