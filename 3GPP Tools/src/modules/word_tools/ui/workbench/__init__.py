"""Modular Word Tools workbench tab package."""

from __future__ import annotations

from modules.word_tools.core.document_opener import get_active_word_document
from .tab import WordExtractorTab

__all__ = [
    "WordExtractorTab",
    "get_active_word_document",
]
