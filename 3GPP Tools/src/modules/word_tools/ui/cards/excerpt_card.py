"""Backward-compatible facade re-exporting `modules.word_tools.ui.cards.excerpt`."""

import logging

from modules.word_tools.core.word_excerpt_extractor import (
    WordExcerptThread,
    WordHeadingScanThread,
)
from modules.word_tools.ui.cards.excerpt import WordExcerptCard

logger = logging.getLogger(__name__)

__all__ = [
    "WordExcerptCard",
    "WordExcerptThread",
    "WordHeadingScanThread",
    "logger",
]
