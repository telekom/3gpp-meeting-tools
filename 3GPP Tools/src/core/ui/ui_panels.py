# --- File: src/core/ui/ui_panels.py ---
import html
import logging
import os
from pathlib import Path
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QListWidget, QDialog, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QTableWidget, QTableWidgetItem, QMessageBox, QApplication, QCheckBox, QMenu
)
from PyQt5.QtCore import pyqtSignal, Qt, QObject
from PyQt5.QtGui import QColor, QBrush, QFont

from core.process_manager import ProcessManager
from core.ui.database_inspector import DatabaseInspectorDialog
from core.ui.theme import C, R, T
from core.ui.ui_components import (
    TREE_STYLE_CLEAN,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_WARNING,
    BUTTON_STYLE_TOOLBAR_DANGER,
)
from core.utils.paths import get_project_root


# ==========================================
# --- HELPER: FILE SIZE FORMATTER ---
# ==========================================
def format_file_size(size_bytes: int) -> str:
    """Formats raw byte count into a human-readable string (KB, MB, GB)."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    for unit in ["KB", "MB", "GB"]:
        size_bytes /= 1024.0
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
    return f"{size_bytes:.2f} TB"


# ==========================================
# --- CUSTOM GUI LOG HANDLER ---
# ==========================================
class GuiLogHandler(logging.Handler, QObject):
    """
    Bridges standard Python logging into the Qt GUI console.

    The handler can be installed before ConsolePanel exists. Log records emitted
    during early application startup are buffered and replayed, in order, when
    attach() is called.

    Once attached, log records are forwarded through a Qt signal. This keeps
    logging from worker threads safe because the actual ConsolePanel update is
    performed on the Qt GUI thread.
    """

    log_emitted = pyqtSignal(str, int)

    def __init__(self):
        logging.Handler.__init__(self)
        QObject.__init__(self)

        self._attached = False
        self._buffer = []

    def emit(self, record):
        """
        Format and forward a LogRecord.

        Before the GUI console is attached, records are retained in memory.
        Afterwards they are emitted through the Qt signal immediately.
        """
        try:
            msg = self.format(record)
            level = record.levelno

            self.acquire()
            try:
                if not self._attached:
                    self._buffer.append((msg, level))
                    return
            finally:
                self.release()

            self.log_emitted.emit(msg, level)

        except Exception:
            self.handleError(record)

    def attach(self, slot):
        """
        Attach the GUI console slot and replay all buffered startup records.

        This should normally be called once, after ConsolePanel has been
        constructed. Repeated calls are ignored to prevent duplicate signal
        connections and duplicate log output.
        """
        self.acquire()
        try:
            if self._attached:
                return

            self.log_emitted.connect(slot)
            self._attached = True

            buffered_records = self._buffer
            self._buffer = []
        finally:
            self.release()

        # Replay outside the logging.Handler lock. The slot executes on the Qt
        # GUI thread and may indirectly cause further logging.
        for message, level in buffered_records:
            self.log_emitted.emit(message, level)


# Backward-compatibility re-exports for dialogs extracted to dedicated modules
from core.ui.database_maintenance_dialog import DatabaseMaintenanceDialog, format_file_size
from core.ui.process_manager_dialog import ProcessManagerDialog
from core.ui.mcp_monitor_dialog import McpMonitorDialog


def _panel_btn_qss() -> str:
    return (
        f"QPushButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_BODY}; "
        f"border: 1px solid {C.BORDER_STRONG}; border-radius: {R.SM}px; "
        f"padding: 3px 10px; font-family: {T.FONT_FAMILY_UI}; "
        f"font-size: {T.SIZE_CAPTION}px; font-weight: 500; }}"
        f"QPushButton:hover {{ background-color: {C.BG_CANVAS}; border-color: {C.BORDER_HOVER}; }}"
    )


# ==========================================
# --- CONSOLE PANEL ---
# ==========================================
class ConsolePanel(QWidget):
    proxy_requested = pyqtSignal()
    update_requested = pyqtSignal()
    java_info_requested = pyqtSignal()
    task_manager_requested = pyqtSignal()
    network_config_requested = pyqtSignal()
    db_maintenance_requested = pyqtSignal()
    mcp_monitor_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 4, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("Terminal Output")
        lbl.setStyleSheet(
            f"font-weight: 600; color: {C.TEXT_SECONDARY}; font-size: {T.SIZE_TABLE}px;"
        )

        btn_style = _panel_btn_qss()

        self.copy_log_btn = QPushButton("Copy Log")
        self.copy_log_btn.setCursor(Qt.PointingHandCursor)
        self.copy_log_btn.setStyleSheet(btn_style)
        self.copy_log_btn.setToolTip("Copy terminal output to clipboard")
        self.copy_log_btn.clicked.connect(self.copy_log)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet(btn_style)
        self.clear_btn.setToolTip("Clear terminal output")
        self.clear_btn.clicked.connect(self.clear_log)

        header.addWidget(lbl)
        header.addStretch()
        header.addWidget(self.copy_log_btn)
        header.addWidget(self.clear_btn)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        self.console.setMinimumHeight(50)

        layout.addLayout(header)
        layout.addWidget(self.console)
        self.setLayout(layout)

    def copy_log(self):
        text = self.console.toPlainText()
        if text:
            QApplication.clipboard().setText(text)

    def clear_log(self):
        self.console.clear()

    def set_proxy_status(self, enabled: bool, tooltip: str = ""):
        """Legacy compatibility hook (proxy state is displayed on SystemStatusBar)."""
        return None

    def log_message(self, message: str, level=logging.INFO):
        color = C.CONSOLE_TEXT
        if "❌" in message or "Error" in message:
            color = C.CONSOLE_LOG_ERR
        elif "⚠️" in message or "Warning" in message:
            color = C.CONSOLE_LOG_WARN
        elif "✅" in message or "Success" in message or "Ready" in message:
            color = C.CONSOLE_LOG_OK

        escaped_msg = html.escape(message).replace(chr(10), "<br>")
        html_msg = f'<span style="color: {color};">{escaped_msg}</span>'
        self.console.append(html_msg)

        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


# ==========================================
# --- QUEUE PANEL (SIDEBAR) ---
# ==========================================
class QueuePanel(QWidget):
    clear_requested = pyqtSignal()
    remove_requested = pyqtSignal(list)
    abort_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 4, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("⏳ Queue")
        lbl.setStyleSheet(
            f"font-weight: 600; color: {C.TEXT_SECONDARY}; font-size: {T.SIZE_TABLE}px;"
        )

        btn_style = _panel_btn_qss()
        abort_style = (
            f"QPushButton {{ background-color: {C.DANGER_SOFT}; color: {C.DANGER}; "
            f"border: 1px solid {C.DANGER_BORDER}; border-radius: {R.SM}px; "
            f"padding: 3px 10px; font-family: {T.FONT_FAMILY_UI}; "
            f"font-size: {T.SIZE_CAPTION}px; font-weight: 600; }}"
            f"QPushButton:hover {{ background-color: {C.DANGER_SOFT_HOVER}; border-color: {C.DANGER_BORDER_HOVER}; }}"
            f"QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}"
        )

        self.abort_btn = QPushButton("🛑 Abort")
        self.abort_btn.setCursor(Qt.PointingHandCursor)
        self.abort_btn.setStyleSheet(abort_style)
        self.abort_btn.setToolTip("Forcefully abort the currently running task.")
        self.abort_btn.setEnabled(False)
        self.abort_btn.clicked.connect(self.abort_requested.emit)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.setStyleSheet(btn_style)
        self.remove_btn.setToolTip("Remove selected item(s) from the waiting queue.")
        self.remove_btn.clicked.connect(self._on_remove_clicked)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet(btn_style)
        self.clear_btn.setToolTip("Remove all waiting items from the queue.")
        self.clear_btn.clicked.connect(self.clear_requested.emit)

        header.addWidget(lbl)
        header.addStretch()
        header.addWidget(self.abort_btn)
        header.addWidget(self.remove_btn)
        header.addWidget(self.clear_btn)

        self.queue_list = QListWidget()
        self.queue_list.setObjectName("queueList")
        self.queue_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.queue_list.setMinimumHeight(50)

        layout.addLayout(header)
        layout.addWidget(self.queue_list)
        self.setLayout(layout)

    def update_list(self, display_items: list):
        self.queue_list.clear()
        self.queue_list.addItems(display_items)

    def _on_remove_clicked(self):
        selected_items = self.queue_list.selectedItems()
        if not selected_items:
            return
        rows_to_remove = [self.queue_list.row(item) for item in selected_items]
        self.remove_requested.emit(rows_to_remove)
