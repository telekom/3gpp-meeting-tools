# --- File: src/modules/nas/core/db/__init__.py ---
"""Modular Protocol SQLite Database (`3gpp_protocol_data.db`) package."""

from .database import NASDatabase
from .schema import (
    RE_NAS_IE_EXACT_LEN,
    RE_NAS_IE_MIN_MAX_LEN,
    RE_NAS_IE_TYPE,
    RE_TRAILING_ASN1_CAPTION,
    parse_version_tuple,
)

__all__ = [
    "RE_NAS_IE_TYPE",
    "RE_NAS_IE_MIN_MAX_LEN",
    "RE_NAS_IE_EXACT_LEN",
    "RE_TRAILING_ASN1_CAPTION",
    "parse_version_tuple",
    "NASDatabase",
]
