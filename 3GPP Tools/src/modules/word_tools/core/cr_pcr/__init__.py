"""Modular 3GPP CR and pseudo-CR (pCR) contribution generator package."""

from modules.word_tools.core.cr_pcr.cr_builder import (
    _update_cr_custom_properties,
    _update_cr_document_xml,
    build_cr_base_document,
)
from modules.word_tools.core.cr_pcr.engine import (
    CrPcrGeneratorEngine,
    CrPcrGeneratorThread,
)
from modules.word_tools.core.cr_pcr.models import (
    C_PROP_NS,
    ChangeBlock,
    ContributionMetadata,
    R_NS,
    VT_NS,
    W_NS,
    extract_clause_summary,
    group_selected_headings_into_blocks,
)
from modules.word_tools.core.cr_pcr.pcr_builder import build_pcr_base_document

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
]
