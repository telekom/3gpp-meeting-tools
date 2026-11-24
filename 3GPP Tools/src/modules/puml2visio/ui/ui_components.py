"""Backward-compatible re-export shim for PlantUML Mini-IDE editor components.

Canonical implementation lives in ``modules.puml2visio.ui.editor``.
"""

from modules.puml2visio.ui.editor import (
    PUML_EXTENSIONS,
    SUPPORTED_EDITOR_EXTENSIONS,
    VISIO_EXTENSIONS,
    CodeDropTextEdit,
    FindReplaceBar,
    LineNumberArea,
)

__all__ = [
    "CodeDropTextEdit",
    "FindReplaceBar",
    "LineNumberArea",
    "PUML_EXTENSIONS",
    "SUPPORTED_EDITOR_EXTENSIONS",
    "VISIO_EXTENSIONS",
]
