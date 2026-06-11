import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTextEdit, QListWidget, QApplication, QDialog, QFrame)
from PyQt5.QtCore import pyqtSignal, Qt, QTimer

from process_manager import ProcessManager


class ProcessManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setModal(True)
        self.resize(450, 300)
        self.setStyleSheet("background-color: #FAFAFA;")

        self.apps = {
            "Microsoft Visio": "visio",
            "Microsoft PowerPoint": "powerpnt",
            "Microsoft Word": "winword"
        }
        self.rows = {}

        self._setup_ui()
        self._refresh_stats()

        # Auto-refresh stats every 2 seconds while dialog is open
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._refresh_stats)
        self.timer.start(2000)

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("🖥️ Active Office Processes")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333; margin-bottom: 10px;")
        layout.addWidget(title)

        desc = QLabel(
            "Identify and kill background processes left hanging by crashes. 'Ghosts' are headless instances without a visible UI window.")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; margin-bottom: 15px;")
        layout.addWidget(desc)

        for display_name, exe_name in self.apps.items():
            row_widget = QFrame()
            row_widget.setStyleSheet(
                "background-color: white; border: 1px solid #DDD; border-radius: 6px; padding: 10px;")
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(10, 5, 10, 5)

            lbl_name = QLabel(display_name)
            lbl_name.setStyleSheet("font-weight: bold; font-size: 13px; border: none;")

            lbl_stats = QLabel("Total: 0 | Ghosts: 0")
            lbl_stats.setStyleSheet("color: #888; font-size: 12px; border: none;")

            btn_kill_ghosts = QPushButton("Kill Ghosts")
            btn_kill_ghosts.setStyleSheet(
                "background-color: #FDF4F0; color: #D83B01; border: 1px solid #F3C3B1; padding: 4px 10px;")
            btn_kill_ghosts.clicked.connect(lambda _, app=exe_name: self._kill(app, True))

            btn_kill_all = QPushButton("Kill All")
            btn_kill_all.setStyleSheet(
                "background-color: #FDEDED; color: #D32F2F; border: 1px solid #E5A4A4; padding: 4px 10px;")
            btn_kill_all.clicked.connect(lambda _, app=exe_name: self._kill(app, False))

            row_layout.addWidget(lbl_name)
            row_layout.addStretch()
            row_layout.addWidget(lbl_stats)
            row_layout.addWidget(btn_kill_ghosts)
            row_layout.addWidget(btn_kill_all)

            layout.addWidget(row_widget)
            self.rows[exe_name] = lbl_stats

        layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

        self.setLayout(layout)

    def _refresh_stats(self):
        data = ProcessManager.get_process_stats()

        # Reset counters
        stats = {exe: {"total": 0, "ghosts": 0} for exe in self.apps.values()}

        for p in data:
            exe = p["Name"].lower()
            if exe in stats:
                stats[exe]["total"] += 1
                if p["IsGhost"]:
                    stats[exe]["ghosts"] += 1

        # Update UI labels
        for exe, data in stats.items():
            color = "#D32F2F" if data["ghosts"] > 0 else "#6A9955"
            ghost_text = f"<span style='color: {color}; font-weight: bold;'>Ghosts: {data['ghosts']}</span>"
            self.rows[exe].setText(f"Total: {data['total']}  |  {ghost_text}")

    def _kill(self, app_name, ghosts_only):
        ProcessManager.kill_processes(app_name, ghosts_only)
        self._refresh_stats()


class ConsolePanel(QWidget):
    # --- Signals ---
    proxy_requested = pyqtSignal()
    update_requested = pyqtSignal()
    task_manager_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 5, 0, 0)

        header = QHBoxLayout()
        lbl = QLabel("Terminal Output")
        lbl.setStyleSheet("font-weight: bold; color: #555;")

        # --- NEW TASK MANAGER BUTTON ---
        self.task_btn = QPushButton("🖥️ Task Manager")
        self.task_btn.setFixedSize(110, 24)
        self.task_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.task_btn.setToolTip("Manage hanging background COM processes.")
        self.task_btn.clicked.connect(self.task_manager_requested.emit)

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
        header.addWidget(self.task_btn)
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
    remove_requested = pyqtSignal(list)
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
        self.queue_list.clear()
        self.queue_list.addItems(display_items)

    def _on_remove_clicked(self):
        selected_items = self.queue_list.selectedItems()
        if not selected_items: return

        rows = sorted([self.queue_list.row(item) for item in selected_items], reverse=True)
        self.remove_requested.emit(rows)