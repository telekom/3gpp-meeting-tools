# --- File: src/modules/word_tools/core/document_sanitizer.py ---
"""Backward-compatible re-export shim for ``modules.word_tools.core.sanitizer``."""

from __future__ import annotations

import logging

from .sanitizer import (
    A_NS,
    AuthorStats,
    CORE_PROPS_XML,
    CP_NS,
    DCMITYPE_NS,
    DCTERMS_NS,
    DC_NS,
    DocumentHygieneAudit,
    LANGUAGE_PRESETS,
    LanguageStats,
    STORY_XML_PATTERN,
    STYLES_XML_PATTERN,
    SanitizeConfig,
    W_AUTHOR,
    W_BIDI,
    W_EAST_ASIA,
    W_INITIALS,
    W_LANG,
    W_NOPROOF,
    W_NS,
    W_VAL,
    XSI_NS,
    _extract_rpr_signature,
    _get_tag_name,
    _is_pure_language_change,
    close_word_document_for_update,
    inspect_document,
    is_file_open_in_word,
    sanitize_document,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "AuthorStats",
    "W_AUTHOR",
    "W_INITIALS",
    "W_NS",
    "A_NS",
    "CP_NS",
    "DC_NS",
    "DCTERMS_NS",
    "DCMITYPE_NS",
    "XSI_NS",
    "CORE_PROPS_XML",
    "W_VAL",
    "W_EAST_ASIA",
    "W_BIDI",
    "W_LANG",
    "W_NOPROOF",
    "STORY_XML_PATTERN",
    "STYLES_XML_PATTERN",
    "LANGUAGE_PRESETS",
    "LanguageStats",
    "DocumentHygieneAudit",
    "SanitizeConfig",
    "close_word_document_for_update",
    "is_file_open_in_word",
    "_get_tag_name",
    "_extract_rpr_signature",
    "_is_pure_language_change",
    "inspect_document",
    "sanitize_document",
]
