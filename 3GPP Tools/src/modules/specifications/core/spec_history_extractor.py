# --- File: src/modules/specifications/core/spec_history_extractor.py ---
"""Backward-compatible facade re-exporting `modules.specifications.core.history_extract`."""

import logging

from modules.specifications.core.history_extract import (
    NS,
    RE_TDOC_TOKEN,
    RE_VERSION,
    SpecHistoryBatchScanThread,
    SpecHistoryExtractor,
    SpecHistoryExtractThread,
    detect_header_columns,
    expand_row_cells,
    extract_cr_cover_sheet_metadata,
    normalize_version_token,
    parse_cell_text,
)

logger = logging.getLogger(__name__)

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
    "logger",
    "normalize_version_token",
    "parse_cell_text",
]
