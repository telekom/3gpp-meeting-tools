import logging
import os
from pathlib import Path
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QListWidget, QDialog, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QTableWidget, QTableWidgetItem, QMessageBox, QApplication
)
from PyQt5.QtCore import pyqtSignal, Qt, QObject
from PyQt5.QtGui import QColor, QBrush, QFont

from core.process_manager import ProcessManager
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
    Intercepts standard Python logging calls globally and emits them as Qt signals.
    This safely bridges background thread logging to the main UI thread.
    """
    log_emitted = pyqtSignal(str, int)

    def __init__(self):
        logging.Handler.__init__(self)
        QObject.__init__(self)

    def emit(self, record):
        msg = self.format(record)
        self.log_emitted.emit(msg, record.levelno)


# ==========================================
# --- DATABASE MAINTENANCE DIALOG ---
# ==========================================
class DatabaseMaintenanceDialog(QDialog):
    """
    Provides inspection, space calculation, and manual VACUUM / WAL compaction
    for all SQLite databases managed by the application.
    """
    log_msg = pyqtSignal(str, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Database Maintenance & Compaction")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(680, 320)
        self.setStyleSheet("background-color: #FAFAFA;")

        self._setup_ui()
        self._refresh_databases()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        title = QLabel("🗄️ SQLite Database Maintenance")
        title.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B; margin-bottom: 2px;")
        layout.addWidget(title)

        desc = QLabel(
            "SQLite preserves deleted pages on an internal freelist, meaning database files do not shrink "
            "automatically. Compacting runs a full <b>VACUUM</b> and flushes Write-Ahead Logs (WAL) to reclaim "
            "disk space and defragment indices."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Database", "File Size", "WAL Log", "Status", "Action"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                gridline-color: #F1F5F9;
            }
            QTableWidget::item {
                padding: 1px 4px;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                padding: 3px 6px;
                font-weight: bold;
                color: #475569;
                border: 1px solid #E2E8F0;
                font-size: 11px;
            }
        """)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()

        self.vacuum_all_btn = QPushButton("🧹 Compact All Databases")
        self.vacuum_all_btn.setFixedHeight(26)
        self.vacuum_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #1E5C99;
                color: white;
                font-weight: bold;
                font-size: 11px;
                border-radius: 4px;
                padding: 3px 12px;
            }
            QPushButton:hover {
                background-color: #15426E;
            }
            QPushButton:disabled {
                background-color: #94A3B8;
            }
        """)
        self.vacuum_all_btn.clicked.connect(self._vacuum_all)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.setFixedHeight(26)
        refresh_btn.setStyleSheet("font-size: 11px; padding: 2px 10px;")
        refresh_btn.clicked.connect(self._refresh_databases)

        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(26)
        close_btn.setMinimumWidth(70)
        close_btn.setStyleSheet("font-size: 11px; padding: 2px 10px;")
        close_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.vacuum_all_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(refresh_btn)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _get_tracked_databases(self) -> List[Dict[str, Any]]:
        """Identifies all standard databases in the application root and subdirectories."""
        root = get_project_root()
        candidate_files = [
            {
                "name": "3GPP Core Data (Specs, Meetings, Work Items)",
                "path": root / "3gpp_data.db",
            },
            {
                "name": "3GPP Protocol Data (NAS, NGAP, RRC, etc.)",
                "path": root / "3gpp_protocol_data.db",
            },
        ]

        for extra_db in root.glob("*.db"):
            if extra_db not in [c["path"] for c in candidate_files] and not extra_db.name.endswith("-wal"):
                candidate_files.append({
                    "name": f"Auxiliary Database ({extra_db.name})",
                    "path": extra_db,
                })

        return candidate_files

    def _refresh_databases(self):
        """Populates the table with properly sized and themed entries."""
        db_entries = self._get_tracked_databases()
        self.table.setRowCount(0)

        # Inherit base application font
        base_font = self.table.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        for row_idx, entry in enumerate(db_entries):
            db_path: Path = entry["path"]
            self.table.insertRow(row_idx)

            # 1. Database Name & Tooltip Path
            name_item = QTableWidgetItem(f"📁 {entry['name']}")
            name_item.setToolTip(str(db_path))
            name_item.setFont(bold_font)
            name_item.setForeground(QBrush(QColor("#0F172A")))
            self.table.setItem(row_idx, 0, name_item)

            if db_path.exists():
                size = db_path.stat().st_size
                size_item = QTableWidgetItem(format_file_size(size))
                size_item.setFont(base_font)
                size_item.setForeground(QBrush(QColor("#334155")))
                size_item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 1, size_item)

                # Check for active WAL file
                wal_file = Path(str(db_path) + "-wal")
                if wal_file.exists() and wal_file.stat().st_size > 0:
                    wal_text = f"🟡 Active ({format_file_size(wal_file.stat().st_size)})"
                    wal_item = QTableWidgetItem(wal_text)
                    wal_item.setForeground(QBrush(QColor("#B45309")))
                else:
                    wal_item = QTableWidgetItem("⚪ Clean")
                    wal_item.setForeground(QBrush(QColor("#64748B")))
                wal_item.setFont(base_font)
                wal_item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 2, wal_item)

                status_item = QTableWidgetItem("🟢 Ready")
                status_item.setFont(base_font)
                status_item.setForeground(QBrush(QColor("#166534")))
                status_item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 3, status_item)

                # Action button styled to match standard tables
                btn_compact = QPushButton("Compact")
                btn_compact.setFixedHeight(22)
                btn_compact.setStyleSheet("""
                    QPushButton {
                        background-color: #F1F5F9;
                        border: 1px solid #CBD5E1;
                        border-radius: 3px;
                        padding: 1px 8px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #0369A1;
                    }
                    QPushButton:hover {
                        background-color: #E0F2FE;
                        border-color: #0284C7;
                    }
                """)
                btn_compact.clicked.connect(lambda _, p=db_path, r=row_idx: self._vacuum_single(p, r))
                self.table.setCellWidget(row_idx, 4, btn_compact)

            else:
                empty_size = QTableWidgetItem("Not Created")
                empty_size.setFont(base_font)
                empty_size.setForeground(QBrush(QColor("#94A3B8")))
                empty_size.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 1, empty_size)

                empty_wal = QTableWidgetItem("-")
                empty_wal.setFont(base_font)
                empty_wal.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 2, empty_wal)

                empty_status = QTableWidgetItem("⚪ Inactive")
                empty_status.setFont(base_font)
                empty_status.setForeground(QBrush(QColor("#94A3B8")))
                empty_status.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, 3, empty_status)

                lbl_none = QLabel("—")
                lbl_none.setFont(base_font)
                lbl_none.setAlignment(Qt.AlignCenter)
                self.table.setCellWidget(row_idx, 4, lbl_none)

    def _vacuum_database_file(self, db_path: Path) -> Tuple[bool, int, int, str]:
        """
        Executes WAL checkpointing and VACUUM on the specified SQLite database file.
        Returns: (success, bytes_before, bytes_after, message)
        """
        if not db_path.exists():
            return False, 0, 0, "File does not exist"

        wal_file = Path(str(db_path) + "-wal")
        bytes_before = db_path.stat().st_size + (wal_file.stat().st_size if wal_file.exists() else 0)

        try:
            conn = sqlite3.connect(str(db_path), check_same_thread=False)
            try:
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")
                conn.execute("PRAGMA optimize;")
            finally:
                conn.close()

            bytes_after = db_path.stat().st_size + (wal_file.stat().st_size if wal_file.exists() else 0)
            return True, bytes_before, bytes_after, "Success"
        except Exception as e:
            return False, bytes_before, bytes_before, str(e)

    def _vacuum_single(self, db_path: Path, row_idx: int):
        """Compacts an individual database and updates table status."""
        status_item = self.table.item(row_idx, 3)
        if status_item:
            status_item.setText("⏳ Compacting...")
            status_item.setForeground(QBrush(QColor("#0284C7")))
        QApplication.processEvents()

        success, before, after, msg = self._vacuum_database_file(db_path)

        if success:
            saved_bytes = max(0, before - after)
            saved_str = f" (reclaimed {format_file_size(saved_bytes)})" if saved_bytes > 0 else " (already compact)"
            logging.info(f"✅ Database compacted: {db_path.name}{saved_str}")
        else:
            logging.error(f"❌ Compaction error on {db_path.name}: {msg}")
            QMessageBox.critical(self, "Compaction Error", f"Failed to compact {db_path.name}:\n{msg}")

        self._refresh_databases()

    def _vacuum_all(self):
        """Compacts all existing tracked databases in sequence."""
        db_entries = self._get_tracked_databases()
        total_saved = 0
        compacted_count = 0

        self.vacuum_all_btn.setEnabled(False)
        self.vacuum_all_btn.setText("⏳ Compacting...")
        QApplication.processEvents()

        for entry in db_entries:
            db_path = entry["path"]
            if db_path.exists():
                success, before, after, msg = self._vacuum_database_file(db_path)
                if success:
                    compacted_count += 1
                    total_saved += max(0, before - after)
                else:
                    logging.error(f"❌ Failed to compact {db_path.name}: {msg}")

        self.vacuum_all_btn.setText("🧹 Compact All Databases")
        self.vacuum_all_btn.setEnabled(True)
        self._refresh_databases()

        summary_msg = f"Compacted {compacted_count} database(s).\nTotal disk space reclaimed: {format_file_size(total_saved)}"
        logging.info(f"✅ {summary_msg.replace(chr(10), ' ')}")
        QMessageBox.information(self, "Compaction Complete", summary_msg)


# ==========================================
# --- PROCESS MANAGER DIALOG (ACCORDION) ---
# ==========================================
class ProcessManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(650, 450)
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
            "Expand an application to view individual documents. You can kill specific frozen documents or safely purge all headless background ghosts."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Application / Document", "Details", "Action"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.tree.setStyleSheet("""
            QTreeWidget { background-color: white; border: 1px solid #DDD; border-radius: 6px; outline: none; }
            QTreeWidget::item { padding: 4px; border-bottom: 1px solid #F0F0F0; }
        """)
        layout.addWidget(self.tree)

        btn_layout = QHBoxLayout()

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
            app_item.setText(0, f"{display_name}")
            app_item.setText(1, f"Total: {total} | Ghosts: {ghosts}")
            app_item.setForeground(0, QBrush(QColor("#333333")))

            font = app_item.font(0)
            font.setBold(True)
            app_item.setFont(0, font)

            action_widget = QWidget()
            act_layout = QHBoxLayout(action_widget)
            act_layout.setContentsMargins(0, 0, 0, 0)
            act_layout.setSpacing(5)

            btn_kill_all = QPushButton("Kill All")
            btn_kill_all.setStyleSheet(
                "background-color: #FDF4F0; color: #D83B01; border: 1px solid #F3C3B1; padding: 2px 8px; border-radius: 3px;"
            )
            btn_kill_all.clicked.connect(lambda _, app=exe_name: self._kill_all(app))
            act_layout.addWidget(btn_kill_all)

            if ghosts > 0:
                btn_kill_ghosts = QPushButton("Kill Ghosts")
                btn_kill_ghosts.setStyleSheet(
                    "background-color: #FAFAFA; color: #555; border: 1px solid #CCC; padding: 2px 8px; border-radius: 3px;"
                )
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
                    "background-color: #FAFAFA; color: #555; border: 1px solid #CCC; padding: 2px 8px; border-radius: 3px;"
                )
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
        self.db_btn.setToolTip("Inspect database sizes, compact freelists (VACUUM), and flush WAL logs.")
        self.db_btn.clicked.connect(self.db_maintenance_requested.emit)

        self.proxy_btn = QPushButton("📡 Proxy")
        self.proxy_btn.setFixedSize(70, 24)
        self.proxy_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.proxy_btn.setToolTip("Update network proxy settings and retry system initialization.")
        self.proxy_btn.clicked.connect(self.proxy_requested.emit)

        self.net_cfg_btn = QPushButton("⚙️ Network")
        self.net_cfg_btn.setFixedSize(70, 24)
        self.net_cfg_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.net_cfg_btn.setToolTip("Update network settings to make the scraper look more human in behavior.")
        self.net_cfg_btn.clicked.connect(self.network_config_requested.emit)

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
        header.addWidget(self.db_btn)
        header.addWidget(self.proxy_btn)
        header.addWidget(self.net_cfg_btn)
        header.addWidget(self.update_btn)
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
        items_to_remove = [item.text() for item in selected_items]
        self.remove_requested.emit(items_to_remove)