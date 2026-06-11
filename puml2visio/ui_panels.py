from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTextEdit, QListWidget, QApplication, QDialog, QTreeWidget, QTreeWidgetItem)
from PyQt5.QtCore import pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QColor, QBrush

from process_manager import ProcessManager


class ProcessManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setModal(True)
        self.resize(550, 400)
        self.setStyleSheet("background-color: #FAFAFA;")

        self.apps = {
            "Microsoft Visio": "visio",
            "Microsoft PowerPoint": "powerpnt",
            "Microsoft Word": "winword"
        }

        self._setup_ui()
        self._refresh_stats()

        # Auto-refresh stats every 2 seconds
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._refresh_stats)
        self.timer.start(2000)

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

        # --- NEW: Expandable Tree Widget ---
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Application / Document", "Details", "Action"])
        self.tree.setColumnWidth(0, 300)
        self.tree.setColumnWidth(1, 100)
        self.tree.setStyleSheet("""
            QTreeWidget { background-color: white; border: 1px solid #DDD; border-radius: 6px; outline: none; }
            QTreeWidget::item { padding: 4px; border-bottom: 1px solid #F0F0F0; }
        """)
        layout.addWidget(self.tree)

        close_btn = QPushButton("Close")
        close_btn.setFixedSize(80, 30)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

        self.setLayout(layout)

    def _refresh_stats(self):
        # Save current expansion state so it doesn't snap closed during auto-refresh
        expanded_states = {}
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            expanded_states[item.text(0)] = item.isExpanded()

        self.tree.clear()
        data = ProcessManager.get_process_stats()

        for display_name, exe_name in self.apps.items():
            app_procs = [p for p in data if p["Name"].lower() == exe_name.lower()]

            if not app_procs:
                continue  # Skip showing apps that aren't running

            total = len(app_procs)
            ghosts = sum(1 for p in app_procs if p["IsGhost"])

            # --- Top Level Item (The Application) ---
            app_item = QTreeWidgetItem(self.tree)
            app_item.setText(0, f"▶ {display_name}")
            app_item.setText(1, f"Total: {total} | Ghosts: {ghosts}")
            app_item.setForeground(0, QBrush(QColor("#333333")))

            font = app_item.font(0)
            font.setBold(True)
            app_item.setFont(0, font)

            if ghosts > 0:
                btn_kill_ghosts = QPushButton("Kill Ghosts")
                btn_kill_ghosts.setStyleSheet(
                    "background-color: #FDF4F0; color: #D83B01; border: 1px solid #F3C3B1; padding: 2px 8px; border-radius: 3px;")
                btn_kill_ghosts.clicked.connect(lambda _, app=exe_name: self._kill_ghosts(app))
                self.tree.setItemWidget(app_item, 2, btn_kill_ghosts)

            # Restore expansion state
            if expanded_states.get(f"▶ {display_name}", False):
                app_item.setExpanded(True)

            # --- Child Items (The Individual Documents) ---
            for p in app_procs:
                child = QTreeWidgetItem(app_item)

                if p["IsGhost"]:
                    child.setText(0, "   👻 Headless Background Instance")
                    child.setForeground(0, QBrush(QColor("#D83B01")))
                else:
                    doc_title = p.get("Title", "").strip() or "Untitled Document"
                    child.setText(0, f"   📄 {doc_title}")
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

    def _kill_single(self, pid):
        ProcessManager.kill_process(pid)
        self._refresh_stats()