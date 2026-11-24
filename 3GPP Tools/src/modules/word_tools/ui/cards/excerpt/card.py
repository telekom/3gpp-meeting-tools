"""Card: Interactive Heading Excerpt & Chapter Slicing."""

from typing import Optional

from PyQt5.QtWidgets import QWidget

from modules.word_tools.core.word_excerpt_extractor import (
    WordExcerptThread,
    WordHeadingScanThread,
)

from .mixins import (
    ExcerptCardExtractionMixin,
    ExcerptCardScanningTreeMixin,
    ExcerptCardUiSetupMixin,
)


class WordExcerptCard(
    ExcerptCardUiSetupMixin,
    ExcerptCardScanningTreeMixin,
    ExcerptCardExtractionMixin,
    QWidget,
):
    """Card: Interactive clause heading scanner, hierarchical tree viewer, and clause excerpt generator."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._excerpt_paths: list = []
        self._excerpt_input_paths: list = []
        self._excerpt_headings: list = []
        self._excerpt_scan_thread: Optional[WordHeadingScanThread] = None
        self._excerpt_extract_thread: Optional[WordExcerptThread] = None
        self._active_scan_threads: list = []
        self._excerpt_scan_gen: int = 0
        self._custom_filename_active: bool = False
        self._setup_ui()
