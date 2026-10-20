# --- File: src/core/ui/database_maintenance_dialog.py ---
import logging
from pathlib import Path
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.database_inspector import DatabaseInspectorDialog
from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)

def format_file_size(size_bytes: int) -> str:
    """Formats raw byte count into a human-readable string (KB, MB, GB)."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    for unit in ["KB", "MB", "GB"]:
        size_bytes /= 1024.0
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
    return f"{size_bytes:.2f} TB"


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
        self.resize(780, 360)
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
            "Inspect database contents, schemas, and diagnostics in read-only mode. SQLite preserves deleted pages on an "
            "internal freelist, so compacting runs a full <b>VACUUM</b> and flushes Write-Ahead Logs (WAL) to "
            "reclaim disk space and defragment indices."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Database", "File Size", "WAL Log", "Status", "Actions"])
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
                "path": root / "db" / "3gpp_data.db",
            },
            {
                "name": "3GPP Protocol Data (NAS, NGAP, RRC, etc.)",
                "path": root / "db" / "3gpp_protocol_data.db",
            },
            {
                "name": "3GPP Specification Text",
                "path": root / "db" / "3gpp_spec_search.db",
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

                # Keep destructive maintenance separate from the read-only inspector.
                action_widget = QWidget()
                action_layout = QHBoxLayout(action_widget)
                action_layout.setContentsMargins(0, 0, 0, 0)
                action_layout.setSpacing(4)

                action_style = """
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
                """

                btn_inspect = QPushButton("Inspect")
                btn_inspect.setFixedHeight(22)
                btn_inspect.setStyleSheet(action_style)
                btn_inspect.setToolTip("Browse tables, schema, and diagnostics in read-only mode.")
                btn_inspect.clicked.connect(
                    lambda _, p=db_path, n=entry["name"]: self._inspect_database(p, n)
                )

                btn_compact = QPushButton("Compact")
                btn_compact.setFixedHeight(22)
                btn_compact.setStyleSheet(action_style)
                btn_compact.clicked.connect(lambda _, p=db_path, r=row_idx: self._vacuum_single(p, r))

                action_layout.addWidget(btn_inspect)
                action_layout.addWidget(btn_compact)
                self.table.setCellWidget(row_idx, 4, action_widget)

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

    def _inspect_database(self, db_path: Path, display_name: str):
        """Open the reusable read-only SQLite inspector for a tracked database."""
        if not db_path.exists():
            QMessageBox.warning(self, "Database Missing", f"Database does not exist:\n{db_path}")
            return

        dialog = DatabaseInspectorDialog(db_path, display_name, self)
        dialog.exec_()

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
