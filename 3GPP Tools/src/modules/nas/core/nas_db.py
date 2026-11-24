# --- File: src/modules/nas/core/nas_db.py ---
"""Backward-compatible re-export shim for modular NASDatabase package."""

import logging

from modules.nas.core.db import (
    RE_NAS_IE_EXACT_LEN,
    RE_NAS_IE_MIN_MAX_LEN,
    RE_NAS_IE_TYPE,
    RE_TRAILING_ASN1_CAPTION,
    NASDatabase,
    parse_version_tuple,
)

logger = logging.getLogger(__name__)

__all__ = [
    "RE_NAS_IE_TYPE",
    "RE_NAS_IE_MIN_MAX_LEN",
    "RE_NAS_IE_EXACT_LEN",
    "RE_TRAILING_ASN1_CAPTION",
    "parse_version_tuple",
    "NASDatabase",
]
