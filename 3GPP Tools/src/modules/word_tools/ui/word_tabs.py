"""Backward-compatible re-export shim for ``modules.word_tools.ui.workbench``."""

from __future__ import annotations

import logging

from .workbench import WordExtractorTab, get_active_word_document

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "WordExtractorTab",
    "get_active_word_document",
]
