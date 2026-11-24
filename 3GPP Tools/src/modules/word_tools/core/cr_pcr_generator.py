"""Backward-compatible re-export shim for 3GPP CR and pCR contribution generation.

Canonical implementation lives in ``modules.word_tools.core.cr_pcr``.
"""

from __future__ import annotations

import logging

from modules.word_tools.core.cr_pcr import (
    C_PROP_NS,
    ChangeBlock,
    ContributionMetadata,
    CrPcrGeneratorEngine,
    CrPcrGeneratorThread,
    R_NS,
    VT_NS,
    W_NS,
    _update_cr_custom_properties,
    _update_cr_document_xml,
    build_cr_base_document,
    build_pcr_base_document,
    extract_clause_summary,
    group_selected_headings_into_blocks,
)

logger = logging.getLogger(__name__)

__all__ = [
    "C_PROP_NS",
    "ChangeBlock",
    "ContributionMetadata",
    "CrPcrGeneratorEngine",
    "CrPcrGeneratorThread",
    "R_NS",
    "VT_NS",
    "W_NS",
    "_update_cr_custom_properties",
    "_update_cr_document_xml",
    "build_cr_base_document",
    "build_pcr_base_document",
    "extract_clause_summary",
    "group_selected_headings_into_blocks",
    "logger",
]
