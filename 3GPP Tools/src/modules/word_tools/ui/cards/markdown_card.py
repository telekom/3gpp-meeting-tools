"""Backward-compatible re-export shim for ``modules.word_tools.ui.cards.markdown``."""

from __future__ import annotations

import logging

from .markdown import WordMarkdownCard

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "WordMarkdownCard",
]
