"""Asynchronous non-blocking launcher for Microsoft Word and local specification documents."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable, Optional, Set, Union

from PyQt5.QtCore import QObject, QThread, pyqtSignal

from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)

# Strong references to running threads so Qt never GCs an active QThread mid-execution.
_ACTIVE_THREADS: Set["DocumentOpenThread"] = set()


class DocumentOpenThread(QThread):
    """Background QThread calling `open_document()` off the Qt GUI event loop."""

    opened = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, file_path: Union[str, Path], parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self.file_path = Path(file_path)

    def run(self) -> None:
        try:
            open_document(self.file_path)
            self.opened.emit(str(self.file_path))
        except Exception as exc:
            logger.error("Async open_document failed for '%s': %s", self.file_path, exc)
            self.failed.emit(str(exc))


def open_document_async(
    file_path: Union[str, Path],
    on_opened: Optional[Callable[[str], None]] = None,
    on_error: Optional[Callable[[str], None]] = None,
    parent: Optional[QObject] = None,
) -> DocumentOpenThread:
    """Launch `open_document()` asynchronously in a worker thread without blocking Qt UI."""
    thread = DocumentOpenThread(file_path, parent=parent)
    _ACTIVE_THREADS.add(thread)

    if on_opened is not None:
        thread.opened.connect(on_opened)
    if on_error is not None:
        thread.failed.connect(on_error)

    def _cleanup() -> None:
        _ACTIVE_THREADS.discard(thread)
        thread.deleteLater()

    thread.finished.connect(_cleanup)
    thread.start()
    return thread
