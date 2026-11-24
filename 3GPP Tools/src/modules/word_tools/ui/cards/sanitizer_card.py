"""Backward-compatible re-export shim for ``modules.word_tools.ui.cards.sanitizer``."""

from __future__ import annotations

import logging

from .sanitizer import WordSanitizerCard

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "WordSanitizerCard",
]
