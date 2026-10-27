# --- File: src/core/ui/database_maintenance_dialog.py ---
import logging
import os
from pathlib import Path
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QDesktopServices, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QProgressBar,
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


class KpiCard(QFrame):
    """Executive KPI metric card with uppercase title, bold value, and subtle description."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setObjectName("kpiCard")
        self.setStyleSheet("""
            QFrame#kpiCard {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)

        self.title_label = QLabel(title.upper())
        self.title_label.setStyleSheet("font-size: 10px; font-weight: bold; color: #64748B; letter-spacing: 0.5px;")

        self.value_label = QLabel("—")
        self.value_label.setWordWrap(False)
        self.value_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #0F172A;")

        self.sub_label = QLabel("")
        self.sub_label.setStyleSheet("font-size: 10px; color: #94A3B8;")

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addWidget(self.sub_label)

    def set_data(self, value: str, sub: str, value_color: str = "#0F172A"):
        self.value_label.setText(value)
        self.value_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {value_color};")
        self.sub_label.setText(sub)


class DatabaseMaintenanceDialog(QDialog):
    """
    Executive database maintenance, storage KPI monitoring, and compaction hub
    for all SQLite databases managed by 3GPP Tools.
    """
    log_msg = pyqtSignal(str, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Database Maintenance & Compaction Hub")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(980, 540)
        self.setMinimumSize(900, 480)
        self.setStyleSheet("background-color: #FAFAFA;")

        self._setup_ui()
        self._refresh_databases()

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(18, 16, 18, 16)
        root_layout.setSpacing(12)

        # ----------------------------------------------------
        # 1. HEADER ROW: Title, Subtitle, & Quick Tool Actions
        # ----------------------------------------------------
        header_layout = QHBoxLayout()
        header_text_layout = QVBoxLayout()
        header_text_layout.setSpacing(2)

        title = QLabel("🗄️ SQLite Database Maintenance & Storage Hub")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #0F172A;")
        header_text_layout.addWidget(title)

        desc = QLabel(
            "Inspect table schemas, calculate reclaimable freelist pages, and run full VACUUM and WAL checkpoints "
            "to defragment indices and reclaim disk space."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px;")
        header_text_layout.addWidget(desc)

        header_layout.addLayout(header_text_layout, 1)

        # Quick tools: Open Folder and Refresh
        top_btn_style = """
            QPushButton {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 5px 12px;
                font-size: 11px;
                font-weight: bold;
                color: #334155;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
        """

        open_folder_btn = QPushButton("📂 Open DB Folder")
        open_folder_btn.setFixedHeight(28)
        open_folder_btn.setStyleSheet(top_btn_style)
        open_folder_btn.setToolTip("Open the SQLite databases directory in Windows File Explorer.")
        open_folder_btn.clicked.connect(self._open_db_folder)
        header_layout.addWidget(open_folder_btn)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.setFixedHeight(28)
        refresh_btn.setStyleSheet(top_btn_style)
        refresh_btn.setToolTip("Re-query database file sizes, freelist counts, and WAL statuses.")
        refresh_btn.clicked.connect(self._refresh_databases)
        header_layout.addWidget(refresh_btn)

        root_layout.addLayout(header_layout)

        # ----------------------------------------------------
        # 2. EXECUTIVE KPI CARDS ROW
        # ----------------------------------------------------
        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(10)

        self.kpi_storage = KpiCard("Total Storage")
        self.kpi_reclaimable = KpiCard("Reclaimable Space")
        self.kpi_wal = KpiCard("WAL Log Status")
        self.kpi_tables = KpiCard("Total User Tables")

        kpi_layout.addWidget(self.kpi_storage)
        kpi_layout.addWidget(self.kpi_reclaimable)
        kpi_layout.addWidget(self.kpi_wal)
        kpi_layout.addWidget(self.kpi_tables)

        root_layout.addLayout(kpi_layout)

        # ----------------------------------------------------
        # 3. STORAGE TABLE
        # ----------------------------------------------------
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Database & Purpose", "Tables", "File Size", "Reclaimable", "WAL Status", "Actions"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Fixed)

        self.table.setColumnWidth(1, 65)
        self.table.setColumnWidth(2, 85)
        self.table.setColumnWidth(3, 105)
        self.table.setColumnWidth(4, 90)
        self.table.setColumnWidth(5, 230)

        self.table.verticalHeader().setDefaultSectionSize(48)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setShowGrid(False)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                gridline-color: #F8FAFC;
                outline: none;
            }
            QTableWidget::item {
                border-bottom: 1px solid #F1F5F9;
                padding: 2px 4px;
            }
            QTableWidget::item:selected {
                background-color: #F8FAFC;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                padding: 8px 6px;
                font-weight: bold;
                color: #475569;
                border: none;
                border-bottom: 2px solid #E2E8F0;
                font-size: 11px;
            }
        """)
        root_layout.addWidget(self.table, 1)

        # ----------------------------------------------------
        # 4. FOOTER CONTROLS & COMPACTION PROGRESS
        # ----------------------------------------------------
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(10)

        self.vacuum_all_btn = QPushButton("🧹 Compact All Databases")
        self.vacuum_all_btn.setFixedHeight(30)
        self.vacuum_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #1E5C99;
                color: white;
                font-weight: bold;
                font-size: 12px;
                border-radius: 5px;
                padding: 4px 16px;
            }
            QPushButton:hover {
                background-color: #15426E;
            }
            QPushButton:disabled {
                background-color: #94A3B8;
            }
        """)
        self.vacuum_all_btn.clicked.connect(self._vacuum_all)
        footer_layout.addWidget(self.vacuum_all_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.setFixedWidth(180)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #E2E8F0;
                border: none;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
                border-radius: 4px;
            }
        """)
        self.progress_bar.setVisible(False)
        footer_layout.addWidget(self.progress_bar)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("font-size: 11px; color: #0284C7; font-weight: bold;")
        self.status_lbl.setVisible(False)
        footer_layout.addWidget(self.status_lbl)

        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(30)
        close_btn.setMinimumWidth(85)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 4px 16px;
                font-size: 11px;
                font-weight: bold;
                color: #475569;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
        """)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        root_layout.addLayout(footer_layout)

    def _open_db_folder(self):
        """Opens the directory containing SQLite databases in Windows File Explorer."""
        root = get_project_root()
        db_dir = root / "db"
        if not db_dir.exists():
            db_dir = root
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(db_dir)))

    def _get_tracked_databases(self) -> List[Dict[str, Any]]:
        """Identifies all standard databases in the application root and subdirectories."""
        root = get_project_root()
        candidate_files = [
            {
                "name": "3GPP Core Data (Specs, Meetings, Work Items)",
                "subtitle": "Specifications, meetings, agendas, and work items",
                "icon": "📋",
                "path": root / "db" / "3gpp_data.db",
            },
            {
                "name": "3GPP Protocol Data (NAS, NGAP, RRC, etc.)",
                "subtitle": "Protocol message definitions, IEs, and diff matrices",
                "icon": "📡",
                "path": root / "db" / "3gpp_protocol_data.db",
            },
            {
                "name": "3GPP Specification Text Search Index",
                "subtitle": "FTS5 trigram full-text search index and clause contents",
                "icon": "🔍",
                "path": root / "db" / "3gpp_spec_search.db",
            },
            {
                "name": "3GPP Specification Change History",
                "subtitle": "Historical CR/p-CR ledger and TDoc cross-reference matrix",
                "icon": "📜",
                "path": root / "db" / "3gpp_spec_history.db",
                "can_wipe": True,
            },
        ]

        seen_paths = {c["path"].resolve() for c in candidate_files}
        for search_dir in [root / "db", root]:
            if search_dir.exists():
                for extra_db in search_dir.glob("*.db"):
                    if extra_db.resolve() not in seen_paths and not extra_db.name.endswith("-wal"):
                        seen_paths.add(extra_db.resolve())
                        candidate_files.append({
                            "name": f"Auxiliary Database ({extra_db.name})",
                            "subtitle": "Custom or auxiliary storage file",
                            "icon": "🗄️",
                            "path": extra_db,
                        })

        return candidate_files

    def _query_db_metrics(self, db_path: Path) -> Dict[str, Any]:
        """Queries database size, freelist reclaimable space, table count, and WAL status."""
        if not db_path.exists():
            return {
                "exists": False,
                "size": 0,
                "wal_size": 0,
                "tables": 0,
                "reclaimable": 0,
                "freelist_pages": 0,
            }

        size = db_path.stat().st_size
        wal_path = Path(str(db_path) + "-wal")
        wal_size = wal_path.stat().st_size if wal_path.exists() else 0

        tables = 0
        reclaimable = 0
        freelist_pages = 0

        try:
            uri = db_path.resolve().as_uri() + "?mode=ro"
            conn = sqlite3.connect(uri, uri=True, timeout=1.0)
            try:
                fl_row = conn.execute("PRAGMA freelist_count;").fetchone()
                freelist_pages = fl_row[0] if fl_row else 0
                ps_row = conn.execute("PRAGMA page_size;").fetchone()
                page_size = ps_row[0] if ps_row else 4096
                reclaimable = freelist_pages * page_size

                tb_row = conn.execute(
                    "SELECT count(*) FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
                ).fetchone()
                tables = tb_row[0] if tb_row else 0
            finally:
                conn.close()
        except Exception as e:
            logger.debug("Failed to query SQLite metrics for %s: %s", db_path.name, e)

        return {
            "exists": True,
            "size": size,
            "wal_size": wal_size,
            "tables": tables,
            "reclaimable": reclaimable,
            "freelist_pages": freelist_pages,
        }

    def _refresh_databases(self):
        """Populates the table and updates executive KPI statistics."""
        db_entries = self._get_tracked_databases()
        self.table.setRowCount(0)

        base_font = self.table.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        total_storage = 0
        total_reclaimable = 0
        total_tables = 0
        active_wal_count = 0
        active_wal_size = 0
        active_db_count = 0

        for row_idx, entry in enumerate(db_entries):
            db_path: Path = entry["path"]
            self.table.insertRow(row_idx)

            metrics = self._query_db_metrics(db_path)

            if metrics["exists"]:
                active_db_count += 1
                total_storage += metrics["size"] + metrics["wal_size"]
                total_reclaimable += metrics["reclaimable"]
                total_tables += metrics["tables"]
                if metrics["wal_size"] > 0:
                    active_wal_count += 1
                    active_wal_size += metrics["wal_size"]

            # ----------------------------------------------------
            # Col 0: Database & Purpose (Two-tier custom widget)
            # ----------------------------------------------------
            cell_widget = QWidget()
            cell_layout = QVBoxLayout(cell_widget)
            cell_layout.setContentsMargins(10, 4, 10, 4)
            cell_layout.setSpacing(2)

            lbl_title = QLabel(f"{entry.get('icon', '🗄️')}  {entry['name']}")
            lbl_title.setStyleSheet("font-size: 12px; font-weight: bold; color: #0F172A;")

            sub_text = f"{db_path.name}   •   {entry.get('subtitle', '')}"
            lbl_sub = QLabel(sub_text)
            lbl_sub.setStyleSheet("font-size: 10px; color: #64748B;")

            cell_layout.addWidget(lbl_title)
            cell_layout.addWidget(lbl_sub)
            cell_widget.setToolTip(str(db_path))
            self.table.setCellWidget(row_idx, 0, cell_widget)

            # Background item for selection support
            bg_item = QTableWidgetItem()
            self.table.setItem(row_idx, 0, bg_item)

            # ----------------------------------------------------
            # Col 1: Tables Count
            # ----------------------------------------------------
            tbl_text = f"{metrics['tables']} tbl" if metrics["exists"] else "—"
            tbl_item = QTableWidgetItem(tbl_text)
            tbl_item.setTextAlignment(Qt.AlignCenter)
            tbl_item.setFont(bold_font)
            tbl_item.setForeground(QBrush(QColor("#334155") if metrics["exists"] else QColor("#94A3B8")))
            tbl_item.setToolTip(f"User tables in database: {metrics['tables']}" if metrics["exists"] else "Not created")
            self.table.setItem(row_idx, 1, tbl_item)

            # ----------------------------------------------------
            # Col 2: File Size
            # ----------------------------------------------------
            size_text = format_file_size(metrics["size"]) if metrics["exists"] else "Not Created"
            size_item = QTableWidgetItem(size_text)
            size_item.setTextAlignment(Qt.AlignCenter)
            size_item.setFont(bold_font)
            size_item.setForeground(QBrush(QColor("#0F172A") if metrics["exists"] else QColor("#94A3B8")))
            self.table.setItem(row_idx, 2, size_item)

            # ----------------------------------------------------
            # Col 3: Reclaimable Space (Freelist calculation)
            # ----------------------------------------------------
            if not metrics["exists"]:
                rec_item = QTableWidgetItem("—")
                rec_item.setForeground(QBrush(QColor("#94A3B8")))
            elif metrics["reclaimable"] > 0:
                rec_item = QTableWidgetItem(f"⚡ {format_file_size(metrics['reclaimable'])}")
                rec_item.setFont(bold_font)
                rec_item.setForeground(QBrush(QColor("#0284C7")))
                rec_item.setToolTip(
                    f"Freelist pages can be reclaimed via Compact:\n"
                    f"{format_file_size(metrics['reclaimable'])} ({metrics['freelist_pages']} unused pages)"
                )
            else:
                rec_item = QTableWidgetItem("✓ Optimal")
                rec_item.setFont(bold_font)
                rec_item.setForeground(QBrush(QColor("#166534")))
                rec_item.setToolTip("Database is defragmented with 0 unused freelist pages.")
            rec_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 3, rec_item)

            # ----------------------------------------------------
            # Col 4: WAL Status
            # ----------------------------------------------------
            if not metrics["exists"]:
                wal_item = QTableWidgetItem("⚪ Inactive")
                wal_item.setForeground(QBrush(QColor("#94A3B8")))
            elif metrics["wal_size"] > 0:
                wal_item = QTableWidgetItem(f"🟡 {format_file_size(metrics['wal_size'])}")
                wal_item.setFont(bold_font)
                wal_item.setForeground(QBrush(QColor("#B45309")))
                wal_item.setToolTip(
                    f"Active Write-Ahead Log: {format_file_size(metrics['wal_size'])}\n"
                    f"Will be checkpointed and truncated during compaction."
                )
            else:
                wal_item = QTableWidgetItem("🟢 Clean")
                wal_item.setFont(bold_font)
                wal_item.setForeground(QBrush(QColor("#166534")))
                wal_item.setToolTip("Write-Ahead Log is fully flushed and clean.")
            wal_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 4, wal_item)

            # ----------------------------------------------------
            # Col 5: Actions (Inspect, Compact, Wipe)
            # ----------------------------------------------------
            if metrics["exists"]:
                action_widget = QWidget()
                action_layout = QHBoxLayout(action_widget)
                action_layout.setContentsMargins(4, 2, 4, 2)
                action_layout.setSpacing(6)

                btn_inspect = QPushButton("🔍 Inspect")
                btn_inspect.setFixedHeight(26)
                btn_inspect.setFixedWidth(72)
                btn_inspect.setStyleSheet("""
                    QPushButton {
                        background-color: #F8FAFC;
                        border: 1px solid #CBD5E1;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #0369A1;
                    }
                    QPushButton:hover {
                        background-color: #E0F2FE;
                        border-color: #0284C7;
                    }
                """)
                btn_inspect.setToolTip("Browse tables, schemas, and diagnostics in read-only mode.")
                btn_inspect.clicked.connect(
                    lambda _, p=db_path, n=entry["name"]: self._inspect_database(p, n)
                )
                action_layout.addWidget(btn_inspect)

                btn_compact = QPushButton("⚡ Compact")
                btn_compact.setFixedHeight(26)
                btn_compact.setFixedWidth(76)
                btn_compact.setStyleSheet("""
                    QPushButton {
                        background-color: #F8FAFC;
                        border: 1px solid #CBD5E1;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #0F172A;
                    }
                    QPushButton:hover {
                        background-color: #E2E8F0;
                        border-color: #94A3B8;
                    }
                """)
                btn_compact.setToolTip("Run PRAGMA wal_checkpoint(TRUNCATE) and VACUUM on this database.")
                btn_compact.clicked.connect(lambda _, p=db_path, r=row_idx: self._vacuum_single(p, r))
                action_layout.addWidget(btn_compact)

                if entry.get("can_wipe"):
                    btn_wipe = QPushButton("🗑️ Wipe")
                    btn_wipe.setFixedHeight(26)
                    btn_wipe.setFixedWidth(64)
                    btn_wipe.setStyleSheet("""
                        QPushButton {
                            background-color: #FEF2F2;
                            border: 1px solid #FECACA;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                            color: #DC2626;
                        }
                        QPushButton:hover {
                            background-color: #FEE2E2;
                            border-color: #EF4444;
                        }
                    """)
                    btn_wipe.setToolTip(
                        "Completely wipe and clear all indexed change history and TDoc cross-references."
                    )
                    btn_wipe.clicked.connect(lambda _, p=db_path, r=row_idx: self._wipe_database(p, r))
                    action_layout.addWidget(btn_wipe)
                else:
                    action_layout.addStretch()

                self.table.setCellWidget(row_idx, 5, action_widget)
            else:
                lbl_none = QLabel("—")
                lbl_none.setFont(base_font)
                lbl_none.setAlignment(Qt.AlignCenter)
                self.table.setCellWidget(row_idx, 5, lbl_none)

        # ----------------------------------------------------
        # 5. UPDATE EXECUTIVE KPI CARDS
        # ----------------------------------------------------
        self.kpi_storage.set_data(
            format_file_size(total_storage),
            f"across {active_db_count} active database(s)",
            value_color="#0F172A"
        )

        if total_reclaimable > 0:
            self.kpi_reclaimable.set_data(
                f"~{format_file_size(total_reclaimable)}",
                "freelist pages ready to reclaim",
                value_color="#0284C7"
            )
        else:
            self.kpi_reclaimable.set_data(
                "✓ Optimal",
                "0 B unused freelist pages",
                value_color="#166534"
            )

        if active_wal_count > 0:
            self.kpi_wal.set_data(
                f"{active_wal_count} Active",
                f"{format_file_size(active_wal_size)} pending checkpoint",
                value_color="#B45309"
            )
        else:
            self.kpi_wal.set_data(
                "All Clean",
                "All WAL logs fully checkpointed",
                value_color="#166534"
            )

        self.kpi_tables.set_data(
            f"{total_tables} Tables",
            "across all active schemas",
            value_color="#0F172A"
        )

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

    def _wipe_database(self, db_path: Path, row_idx: int):
        """Completely wipes and reinitializes a wipeable auxiliary database (e.g. Change History)."""
        reply = QMessageBox.warning(
            self,
            "Confirm Database Wipe",
            f"Are you sure you want to completely wipe:\n{db_path.name}?\n\n"
            "All indexed specification change history and TDoc cross-references will be cleared.\n"
            "This will NOT affect your specifications catalogue or meeting data.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        from modules.specifications.core.spec_history_db import SpecHistoryDatabase
        success = SpecHistoryDatabase(db_path).wipe_database()
        if success:
            QMessageBox.information(self, "Wipe Complete", f"Successfully cleared and reinitialized {db_path.name}.")
        else:
            QMessageBox.critical(self, "Wipe Error", f"Failed to wipe {db_path.name}.")

        self._refresh_databases()

    def _vacuum_single(self, db_path: Path, row_idx: int):
        """Compacts an individual database and updates table status with real-time feedback."""
        self.status_lbl.setText(f"Compacting {db_path.name}...")
        self.status_lbl.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(50)
        QApplication.processEvents()

        success, before, after, msg = self._vacuum_database_file(db_path)

        self.progress_bar.setValue(100)
        if success:
            saved_bytes = max(0, before - after)
            saved_str = f" (reclaimed {format_file_size(saved_bytes)})" if saved_bytes > 0 else " (already compact)"
            logger.info(f"✅ Database compacted: {db_path.name}{saved_str}")
            self.status_lbl.setText(f"Compacted {db_path.name}{saved_str}")
        else:
            logger.error(f"❌ Compaction error on {db_path.name}: {msg}")
            QMessageBox.critical(self, "Compaction Error", f"Failed to compact {db_path.name}:\n{msg}")

        self._refresh_databases()
        self.progress_bar.setVisible(False)
        self.status_lbl.setVisible(False)

    def _vacuum_all(self):
        """Compacts all existing tracked databases in sequence with real-time progress."""
        db_entries = self._get_tracked_databases()
        existing_entries = [e for e in db_entries if e["path"].exists()]
        if not existing_entries:
            return

        total_count = len(existing_entries)
        total_saved = 0
        compacted_count = 0

        self.vacuum_all_btn.setEnabled(False)
        self.vacuum_all_btn.setText("⏳ Compacting...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_lbl.setVisible(True)
        QApplication.processEvents()

        for idx, entry in enumerate(existing_entries):
            db_path = entry["path"]
            self.status_lbl.setText(f"Compacting {db_path.name} ({idx + 1}/{total_count})...")
            self.progress_bar.setValue(int((idx / total_count) * 100))
            QApplication.processEvents()

            success, before, after, msg = self._vacuum_database_file(db_path)
            if success:
                compacted_count += 1
                total_saved += max(0, before - after)
            else:
                logger.error(f"❌ Failed to compact {db_path.name}: {msg}")

        self.progress_bar.setValue(100)
        self.status_lbl.setText(f"Compaction complete. Reclaimed {format_file_size(total_saved)}.")
        self.vacuum_all_btn.setText("🧹 Compact All Databases")
        self.vacuum_all_btn.setEnabled(True)
        self._refresh_databases()

        summary_msg = f"Compacted {compacted_count} database(s).\nTotal disk space reclaimed: {format_file_size(total_saved)}"
        logger.info(f"✅ {summary_msg.replace(chr(10), ' ')}")
        QMessageBox.information(self, "Compaction Complete", summary_msg)

        self.progress_bar.setVisible(False)
        self.status_lbl.setVisible(False)
