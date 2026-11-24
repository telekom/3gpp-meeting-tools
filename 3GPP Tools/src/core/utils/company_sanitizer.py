# --- File: src/core/utils/company_sanitizer.py ---
"""Backward-compatible re-export shim for the modular CompanySanitizer package."""

import logging

from core.utils.company_sanitization import CompanySanitizer

logger = logging.getLogger(__name__)

__all__ = ["CompanySanitizer"]
