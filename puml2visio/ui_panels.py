import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTextEdit, QListWidget, QApplication)
from PyQt5.QtCore import pyqtSignal


class ConsolePanel(QWidget):
    # --- Signals ---
    proxy_requested = pyqtSignal()
    update_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 5, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("Terminal Output")
        lbl.setStyleSheet("font-weight: bold; color: #555;")

        self.proxy_btn = QPushButton("📡 Proxy")
        self.proxy_btn.setFixedSize(70, 24)
        self.proxy_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.proxy_btn.setToolTip("Update network proxy settings and retry system initialization.")
        self.proxy_btn.clicked.connect(self.proxy_requested.emit)

        self.update_btn = QPushButton("🔄 Update JAR")
        self.update_btn.setFixedSize(85, 24)
        self.update_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.update_btn.setToolTip("Check online if a newer version of PlantUML is available.")
        self.update_btn.clicked.connect(self.update_requested.emit)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setFixedSize(60, 24)
        self.clear_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.clear_btn.clicked.connect(self.clear_log)

        header.addWidget(lbl)
        header.addStretch()
        header.addWidget(self.proxy_btn)
        header.addWidget(self.update_btn)
        header.addWidget(self.clear_btn)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")

        layout.addLayout(header)
        layout.addWidget(self.console)
        self.setLayout(layout)

    def clear_log(self):
        self.console.clear()

    def log_message(self, message: str, level=logging.INFO):
        """Handles HTML color coding and auto-scrolling."""
        color = "#D4D4D4"
        if "❌" in message or "Error" in message:
            color = "#F44747"
        elif "⚠️" in message or "Warning" in message:
            color = "#D7BA7D"
        elif "✅" in message or "Success" in message or "Ready" in message:
            color = "#6A9955"

        html_msg = f'<span style="color: {color};">{message.replace(chr(10), "<br>")}</span>'
        self.console.append(html_msg)

        QApplication.processEvents()
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


class QueuePanel(QWidget):
    # --- Signals ---
    remove_requested = pyqtSignal(list)  # Emits a list of row integers to delete
    clear_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 5, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("Queue")
        lbl.setStyleSheet("font-weight: bold; color: #555;")

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
        header.addWidget(self.remove_btn)
        header.addWidget(self.clear_btn)

        self.queue_list = QListWidget()
        self.queue_list.setObjectName("queueList")
        self.queue_list.setSelectionMode(QListWidget.ExtendedSelection)

        layout.addLayout(header)
        layout.addWidget(self.queue_list)
        self.setLayout(layout)

    def update_list(self, display_items: list):
        """Re-draws the list with the provided formatted strings."""
        self.queue_list.clear()
        self.queue_list.addItems(display_items)

    def _on_remove_clicked(self):
        """Finds which rows the user highlighted and passes them to the Traffic Cop."""
        selected_items = self.queue_list.selectedItems()
        if not selected_items: return

        rows = sorted([self.queue_list.row(item) for item in selected_items], reverse=True)
        self.remove_requested.emit(rows)