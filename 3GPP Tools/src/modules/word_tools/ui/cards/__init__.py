"""Operation cards and modular widgets for Word Tools."""

from modules.word_tools.ui.cards.compare_card import WordCompareCard
from modules.word_tools.ui.cards.converter_card import WordConverterCard, WordVisioCard
from modules.word_tools.ui.cards.cr_pcr_card import WordCrPcrCard
from modules.word_tools.ui.cards.excerpt_card import WordExcerptCard
from modules.word_tools.ui.cards.markdown_card import WordMarkdownCard
from modules.word_tools.ui.cards.sanitizer_card import WordSanitizerCard

__all__ = [
    "WordCompareCard",
    "WordConverterCard",
    "WordVisioCard",
    "WordCrPcrCard",
    "WordExcerptCard",
    "WordMarkdownCard",
    "WordSanitizerCard",
]
