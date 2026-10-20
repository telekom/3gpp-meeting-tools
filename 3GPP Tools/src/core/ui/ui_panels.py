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

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 5, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("Terminal Output")
        lbl.setStyleSheet("font-weight: bold; color: #555;")

        self.task_btn = QPushButton("🖥️ Task Manager")
        self.task_btn.setFixedSize(110, 24)
        self.task_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.task_btn.setToolTip("Manage hanging background COM processes.")
        self.task_btn.clicked.connect(self.task_manager_requested.emit)

        self.db_btn = QPushButton("🗄️ Database")
        self.db_btn.setFixedSize(85, 24)
        self.db_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.db_btn.setToolTip("Browse database contents and diagnostics, inspect sizes, and compact SQLite files.")
        self.db_btn.clicked.connect(self.db_maintenance_requested.emit)

        self.proxy_btn = QPushButton("📡 Proxy")
        self.proxy_btn.setFixedSize(95, 24)
        self.proxy_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.proxy_btn.setToolTip("Update network proxy settings and retry system initialization.")
        self.proxy_btn.clicked.connect(self.proxy_requested.emit)

        self.net_cfg_btn = QPushButton("⚙️ Network")
        self.net_cfg_btn.setFixedSize(70, 24)
        self.net_cfg_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.net_cfg_btn.setToolTip("Update network settings to make the scraper look more human in behavior.")
        self.net_cfg_btn.clicked.connect(self.network_config_requested.emit)

        self.java_btn = QPushButton("☕ Java")
        self.java_btn.setFixedSize(65, 24)
        self.java_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.java_btn.setToolTip("Inspect installed Java runtime version, vendor, architecture, and binary path.")
        self.java_btn.clicked.connect(self.java_info_requested.emit)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setFixedSize(60, 24)
        self.clear_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.clear_btn.clicked.connect(self.clear_log)

        header.addWidget(lbl)
        header.addStretch()
        header.addWidget(self.task_btn)
        header.addWidget(self.db_btn)
        header.addWidget(self.proxy_btn)
        header.addWidget(self.net_cfg_btn)
        header.addWidget(self.java_btn)
        header.addWidget(self.clear_btn)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        self.console.setMinimumHeight(50)

        layout.addLayout(header)
        layout.addWidget(self.console)
        self.setLayout(layout)

    def clear_log(self):
        self.console.clear()

    def set_proxy_status(self, enabled: bool, tooltip: str = ""):
        """Updates the Terminal toolbar Proxy button to reflect active proxy state."""
        if enabled:
            self.proxy_btn.setText("🛡️ Proxy: ON")
            self.proxy_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F0FDF4;
                    color: #15803D;
                    border: 1px solid #BBF7D0;
                    border-radius: 4px;
                    padding: 2px 6px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #DCFCE7;
                    border-color: #86EFAC;
                }
            """)
        else:
            self.proxy_btn.setText("📡 Proxy")
            self.proxy_btn.setStyleSheet("padding: 2px; font-size: 11px;")

        if tooltip:
            self.proxy_btn.setToolTip(tooltip)

    def log_message(self, message: str, level=logging.INFO):
        color = "#D4D4D4"
        if "❌" in message or "Error" in message:
            color = "#F44747"
        elif "⚠️" in message or "Warning" in message:
            color = "#D7BA7D"
        elif "✅" in message or "Success" in message or "Ready" in message:
            color = "#6A9955"

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
        layout.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("⏳ Queue")
        lbl.setStyleSheet("font-weight: bold; color: #555;")

        self.abort_btn = QPushButton("🛑 Abort")
        self.abort_btn.setFixedSize(65, 24)
        self.abort_btn.setStyleSheet("padding: 2px; font-size: 11px; color: #D32F2F; font-weight: bold;")
        self.abort_btn.setToolTip("Forcefully abort the currently running task.")
        self.abort_btn.setEnabled(False)
        self.abort_btn.clicked.connect(self.abort_requested.emit)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setFixedSize(60, 24)
        self.remove_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.remove_btn.setToolTip("Remove selected item(s) from the waiting queue.")
        self.remove_btn.clicked.connect(self._on_remove_clicked)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setFixedSize(60, 24)
        self.clear_btn.setStyleSheet("padding: 2px; font-size: 11px;")
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