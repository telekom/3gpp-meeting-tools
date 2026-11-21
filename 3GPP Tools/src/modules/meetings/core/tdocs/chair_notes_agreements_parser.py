"""Backward-compatible re-export shim for ChairNotesAgreementsParser and helpers.

The implementation is organized under ``modules.meetings.core.tdocs.chair_notes``.
"""

from __future__ import annotations

from .chair_notes import (
    _BLUE_HEX_PREFIXES,
    _CHAIR_RESULT_PREFIXES,
    _GREEN_HEX_PREFIXES,
    _PREPOSITION_CONTINUATION_RE,
    _PUA_RE,
    _RAN2_PAREN_WI_RE,
    _RAN3_WID_BRACKET_RE,
    _RELEASE_RE,
    _TDOC_ID_RE,
    _TRIVIAL_STATUS_SET,
    CHAIR_NOTES_CACHE_VERSION,
    ChairNotesAgreementsParser,
    _classify_agreement_kind,
    _get_paragraph_dominant_color,
    _resolve_inherited_field,
    clean_chair_notes_text,
    iter_block_items,
    parse_chair_result,
)

__all__ = [
    "CHAIR_NOTES_CACHE_VERSION",
    "ChairNotesAgreementsParser",
    "_BLUE_HEX_PREFIXES",
    "_CHAIR_RESULT_PREFIXES",
    "_GREEN_HEX_PREFIXES",
    "_PREPOSITION_CONTINUATION_RE",
    "_PUA_RE",
    "_RAN2_PAREN_WI_RE",
    "_RAN3_WID_BRACKET_RE",
    "_RELEASE_RE",
    "_TDOC_ID_RE",
    "_TRIVIAL_STATUS_SET",
    "_classify_agreement_kind",
    "_get_paragraph_dominant_color",
    "_resolve_inherited_field",
    "clean_chair_notes_text",
    "iter_block_items",
    "parse_chair_result",
]
