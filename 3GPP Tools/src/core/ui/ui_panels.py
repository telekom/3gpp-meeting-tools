import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTextEdit, QListWidget, QDialog, QTreeWidget, QTreeWidgetItem, QHeaderView)
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QColor, QBrush

from core.process_manager import ProcessManager


# ==========================================
# --- PROCESS MANAGER DIALOG (ACCORDION) ---
# ==========================================
class ProcessManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(650, 450)  # Made slightly wider to comfortably fit multiple buttons
        self.setStyleSheet("background-color: #FAFAFA;")

        self.apps = {
            "Microsoft Visio": "visio",
            "Microsoft PowerPoint": "powerpnt",
            "Microsoft Word": "winword"
        }

        self._setup_ui()
        self._refresh_stats()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("🖥️ Active Office Processes")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333; margin-bottom: 5px;")
        layout.addWidget(title)

        desc = QLabel(
            "Expand an application to view individual documents. You can kill specific frozen documents or safely purge all headless background ghosts.")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Application / Document", "Details", "Action"])

        # --- FIX: Dynamic Column Sizing instead of hardcoded pixels ---
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.tree.setStyleSheet("""
            QTreeWidget { background-color: white; border: 1px solid #DDD; border-radius: 6px; outline: none; }
            QTreeWidget::item { padding: 4px; border-bottom: 1px solid #F0F0F0; }
        """)
        layout.addWidget(self.tree)

        btn_layout = QHBoxLayout()

        # --- FIX: Removed fixed width cap so text can naturally expand ---
        refresh_btn = QPushButton("🔄 Refresh List")
        refresh_btn.setMinimumHeight(30)
        refresh_btn.clicked.connect(self._refresh_stats)

        close_btn = QPushButton("Close")
        close_btn.setMinimumHeight(30)
        close_btn.setMinimumWidth(80)
        close_btn.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(refresh_btn)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _refresh_stats(self):
        expanded_states = {}
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            expanded_states[item.text(0)] = item.isExpanded()

        self.tree.clear()
        data = ProcessManager.get_process_stats()

        for display_name, exe_name in self.apps.items():
            app_procs = [p for p in data if p["Name"].lower() == exe_name.lower()]

            if not app_procs:
                continue

            total = len(app_procs)
            ghosts = sum(1 for p in app_procs if p["IsGhost"])

            app_item = QTreeWidgetItem(self.tree)
            # --- FIX: Removed the manual unicode arrow to prevent collision ---
            app_item.setText(0, f"{display_name}")
            app_item.setText(1, f"Total: {total} | Ghosts: {ghosts}")
            app_item.setForeground(0, QBrush(QColor("#333333")))

            font = app_item.font(0)
            font.setBold(True)
            app_item.setFont(0, font)

            # --- FIX: Container to hold multiple buttons in the Action column ---
            action_widget = QWidget()
            act_layout = QHBoxLayout(action_widget)
            act_layout.setContentsMargins(0, 0, 0, 0)
            act_layout.setSpacing(5)

            # 1. Global "Kill All" Button (Always visible)
            btn_kill_all = QPushButton("Kill All")
            btn_kill_all.setStyleSheet(
                "background-color: #FDF4F0; color: #D83B01; border: 1px solid #F3C3B1; padding: 2px 8px; border-radius: 3px;")
            btn_kill_all.clicked.connect(lambda _, app=exe_name: self._kill_all(app))
            act_layout.addWidget(btn_kill_all)

            # 2. "Kill Ghosts" Button (Only if ghosts exist)
            if ghosts > 0:
                btn_kill_ghosts = QPushButton("Kill Ghosts")
                btn_kill_ghosts.setStyleSheet(
                    "background-color: #FAFAFA; color: #555; border: 1px solid #CCC; padding: 2px 8px; border-radius: 3px;")
                btn_kill_ghosts.clicked.connect(lambda _, app=exe_name: self._kill_ghosts(app))
                act_layout.addWidget(btn_kill_ghosts)

            self.tree.setItemWidget(app_item, 2, action_widget)

            if expanded_states.get(f"{display_name}", False):
                app_item.setExpanded(True)

            for p in app_procs:
                child = QTreeWidgetItem(app_item)

                if p["IsGhost"]:
                    child.setText(0, "👻 Headless Background Instance")
                    child.setForeground(0, QBrush(QColor("#D83B01")))
                else:
                    doc_title = p.get("Title", "").strip() or "Untitled Document"
                    child.setText(0, f"📄 {doc_title}")
                    child.setForeground(0, QBrush(QColor("#166534")))

                child.setText(1, f"PID: {p['Id']}")

                btn_kill_single = QPushButton("Kill")
                btn_kill_single.setStyleSheet(
                    "background-color: #FAFAFA; color: #555; border: 1px solid #CCC; padding: 2px 8px; border-radius: 3px;")
                btn_kill_single.clicked.connect(lambda _, pid=p['Id']: self._kill_single(pid))
                self.tree.setItemWidget(child, 2, btn_kill_single)

    def _kill_ghosts(self, app_name):
        ProcessManager.kill_app_ghosts(app_name)
        self._refresh_stats()

    def _kill_all(self, app_name):
        ProcessManager.kill_app_all(app_name)
        self._refresh_stats()

    def _kill_single(self, pid):
        ProcessManager.kill_process(pid)
        self._refresh_stats()


# ==========================================
# --- CONSOLE PANEL ---
# ==========================================
class ConsolePanel(QWidget):
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

        # Remove QApplication.processEvents()

        # Smoothly scroll to bottom without forcing a global UI lock
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
        # Pass the text (or indices) of selected items to the manager
        items_to_remove = [item.text() for item in selected_items]
        self.remove_requested.emit(items_to_remove)