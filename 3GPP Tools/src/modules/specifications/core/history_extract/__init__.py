"""Headless OpenXML Change History extraction package."""

from .extractor import SpecHistoryExtractor
from .threads import SpecHistoryBatchScanThread, SpecHistoryExtractThread
from .xml_utils import (
    NS,
    RE_TDOC_TOKEN,
    RE_VERSION,
    detect_header_columns,
    expand_row_cells,
    extract_cr_cover_sheet_metadata,
    normalize_version_token,
    parse_cell_text,
)

__all__ = [
    "NS",
    "RE_TDOC_TOKEN",
    "RE_VERSION",
    "SpecHistoryBatchScanThread",
    "SpecHistoryExtractThread",
    "SpecHistoryExtractor",
    "detect_header_columns",
    "expand_row_cells",
    "extract_cr_cover_sheet_metadata",
    "normalize_version_token",
    "parse_cell_text",
]
