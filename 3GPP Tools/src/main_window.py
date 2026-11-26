# --- File: src/main_window.py ---
"""
Backward-compatible import shim for the main application shell (`DragDropUI`).

The implementation was split into ``src/core/ui/main_window/`` (see
``core/ui/main_window/README.md`` for a 'where to edit what' map).

Note for tests: ``QTimer`` and ``QMessageBox`` are re-exported so patches such
as ``patch("main_window.QMessageBox.question")`` and
``patch("main_window.QTimer.singleShot")`` continue to work seamlessly.
"""
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMessageBox

from core.ui.main_window import DragDropUI

__all__ = ["DragDropUI", "QMessageBox", "QTimer"]
