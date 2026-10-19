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


# ==========================================
# --- PROCESS MANAGER DIALOG (ACCORDION) ---
# ==========================================
class ProcessManagerDialog(QDialog):
    """Compact Office document/process manager with safe document-level closing."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(980, 600)

        self.apps = {
            "Microsoft Word": "winword",
            "Microsoft PowerPoint": "powerpnt",
            "Microsoft Visio": "visio",
        }
        # Session-only set of documents that the user has unchecked / chosen to keep open.
        self._kept_open = {"winword": set(), "powerpnt": set()}
        self._document_items = []
        self._app_items = {}
        self._updating_checks = False

        self._setup_ui()
        self._refresh_stats()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(8)

        title = QLabel("Active Office Documents")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        desc = QLabel(
            "Checked documents will be closed. Uncheck anything you want to keep open. "
            "Process termination is available under Emergency."
        )
        desc.setStyleSheet("color: #64748B; font-size: 11px;")
        layout.addWidget(desc)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Close / Application / Document", "Location / Status"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.setAlternatingRowColors(False)
        self.tree.setRootIsDecorated(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setStyleSheet(TREE_STYLE_CLEAN)
        self.tree.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.tree, 1)

        # Selection helpers are intentionally lightweight: they affect documents only.
        selection_bar = QHBoxLayout()
        selection_bar.setSpacing(6)
        selection_label = QLabel("Documents:")
        selection_label.setStyleSheet("color: #64748B; font-size: 11px;")
        selection_bar.addWidget(selection_label)

        select_all_btn = QPushButton("Select all")
        select_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        select_all_btn.clicked.connect(lambda: self._set_all_documents_checked(True))
        selection_bar.addWidget(select_all_btn)

        select_none_btn = QPushButton("Select none")
        select_none_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        select_none_btn.clicked.connect(lambda: self._set_all_documents_checked(False))
        selection_bar.addWidget(select_none_btn)
        selection_bar.addStretch()
        layout.addLayout(selection_bar)

        footer = QHBoxLayout()
        footer.setSpacing(8)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #64748B; font-size: 11px;")
        footer.addWidget(self.status_label)
        footer.addStretch()

        self.close_docs_btn = QPushButton("Close Selected")
        self.close_docs_btn.setToolTip(
            "Close checked Word and PowerPoint documents normally. Office may prompt to save modified documents."
        )
        self.close_docs_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_docs_btn.clicked.connect(lambda: self._close_selected_documents(False))
        footer.addWidget(self.close_docs_btn)

        self.discard_btn = QPushButton("Close && Discard Changes")
        self.discard_btn.setToolTip(
            "Close checked Word and PowerPoint documents and deliberately discard unsaved edits."
        )
        self.discard_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.discard_btn.clicked.connect(lambda: self._close_selected_documents(True))
        footer.addWidget(self.discard_btn)

        self.emergency_btn = QPushButton("Emergency ▾")
        self.emergency_btn.setToolTip("Force-terminate Office processes. Use only when graceful closing is not possible.")
        self.emergency_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.emergency_menu = QMenu(self.emergency_btn)
        self.emergency_btn.setMenu(self.emergency_menu)
        footer.addWidget(self.emergency_btn)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        refresh_btn.clicked.connect(self._refresh_stats)
        footer.addWidget(refresh_btn)

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.clicked.connect(self.accept)
        footer.addWidget(close_btn)
        layout.addLayout(footer)

    @staticmethod
    def _doc_key(doc):
        return (doc.get("full_name") or doc.get("name") or "").casefold()

    def _refresh_stats(self):
        expanded_states = {
            self.tree.topLevelItem(i).data(0, Qt.UserRole): self.tree.topLevelItem(i).isExpanded()
            for i in range(self.tree.topLevelItemCount())
        }

        self._updating_checks = True
        self.tree.clear()
        self._document_items = []
        self._app_items = {}
        self.emergency_menu.clear()
        data = ProcessManager.get_process_stats()

        total_documents = 0
        total_checked = 0

        for display_name, exe_name in self.apps.items():
            app_procs = [p for p in data if p["Name"].lower() == exe_name.lower()]
            documents = ProcessManager.get_open_office_documents(exe_name) if exe_name in self._kept_open else []
            if not app_procs and not documents:
                continue

            total = len(app_procs)
            ghosts = sum(1 for p in app_procs if p["IsGhost"])
            app_item = QTreeWidgetItem(self.tree)
            app_item.setData(0, Qt.UserRole, exe_name)
            app_item.setText(0, display_name)
            app_item.setForeground(0, QBrush(QColor("#1E293B")))
            font = app_item.font(0)
            font.setBold(True)
            app_item.setFont(0, font)
            self._app_items[exe_name] = app_item

            if documents:
                status = f"{len(documents)} document{'s' if len(documents) != 1 else ''} · {total} process{'es' if total != 1 else ''}"
                if ghosts:
                    status += f" · {ghosts} background"
                app_item.setText(1, status)

                for doc in documents:
                    child = QTreeWidgetItem(app_item)
                    key = self._doc_key(doc)
                    child.setData(0, Qt.UserRole, (exe_name, key))
                    child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                    checked = key not in self._kept_open[exe_name]
                    child.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)
                    name = doc.get("name", "Untitled Document")
                    if doc.get("modified") is True:
                        name += "  • modified"
                        child.setForeground(0, QBrush(QColor("#B45309")))
                    child.setText(0, name)
                    child.setText(1, doc.get("full_name") or "Unsaved document")
                    child.setToolTip(0, "Checked = close this document; unchecked = keep it open")
                    child.setToolTip(1, doc.get("full_name") or "Unsaved document")
                    self._document_items.append(child)
                    total_documents += 1
                    total_checked += int(checked)
            else:
                # Background-only applications stay compact. Process details remain available
                # by expanding the application row, but no action buttons clutter the tree.
                app_item.setText(1, f"{total} background process{'es' if total != 1 else ''}")
                app_item.setForeground(1, QBrush(QColor("#64748B")))
                for proc in app_procs:
                    child = QTreeWidgetItem(app_item)
                    title = proc.get("Title", "").strip()
                    child.setText(0, title or "Background instance")
                    child.setText(1, f"PID {proc['Id']}" + (" · headless" if proc["IsGhost"] else ""))
                    child.setForeground(0, QBrush(QColor("#64748B")))

            # Document-bearing apps are useful expanded; background-only apps default collapsed.
            default_expanded = bool(documents)
            app_item.setExpanded(expanded_states.get(exe_name, default_expanded))

            kill_all = self.emergency_menu.addAction(f"Kill all {display_name} processes")
            kill_all.triggered.connect(lambda _, app=exe_name: self._kill_all(app))
            if ghosts:
                kill_ghosts = self.emergency_menu.addAction(f"Kill {display_name} background processes ({ghosts})")
                kill_ghosts.triggered.connect(lambda _, app=exe_name: self._kill_ghosts(app))

        if self.emergency_menu.actions():
            self.emergency_menu.addSeparator()
            kill_all_office = self.emergency_menu.addAction("Kill all Office processes")
            kill_all_office.triggered.connect(self._kill_all_office)

        self._updating_checks = False
        self._update_selection_status(total_documents, total_checked)

    def _on_item_changed(self, item, column):
        if self._updating_checks or column != 0:
            return
        payload = item.data(0, Qt.UserRole)
        if not isinstance(payload, tuple) or len(payload) != 2:
            return
        app_name, key = payload
        if item.checkState(0) == Qt.Checked:
            self._kept_open[app_name].discard(key)
        else:
            self._kept_open[app_name].add(key)
        self._update_selection_status()

    def _set_all_documents_checked(self, checked):
        self._updating_checks = True
        state = Qt.Checked if checked else Qt.Unchecked
        for item in self._document_items:
            item.setCheckState(0, state)
            app_name, key = item.data(0, Qt.UserRole)
            if checked:
                self._kept_open[app_name].discard(key)
            else:
                self._kept_open[app_name].add(key)
        self._updating_checks = False
        self._update_selection_status()

    def _update_selection_status(self, total=None, checked=None):
        if total is None:
            total = len(self._document_items)
        if checked is None:
            checked = sum(item.checkState(0) == Qt.Checked for item in self._document_items)
        kept = total - checked
        if total:
            text = f"{checked} selected to close"
            if kept:
                text += f" · {kept} kept open"
        else:
            text = "No closable Word or PowerPoint documents"
        self.status_label.setText(text)
        self.close_docs_btn.setEnabled(checked > 0)
        self.discard_btn.setEnabled(checked > 0)

    def _close_selected_documents(self, discard_changes):
        selected_by_app = {"winword": 0, "powerpnt": 0}
        for item in self._document_items:
            if item.checkState(0) == Qt.Checked:
                app_name, _ = item.data(0, Qt.UserRole)
                selected_by_app[app_name] += 1
        selected_count = sum(selected_by_app.values())
        if not selected_count:
            return

        if discard_changes:
            answer = QMessageBox.warning(
                self,
                "Discard Unsaved Changes?",
                f"Close {selected_count} selected document(s) and permanently discard any unsaved changes?\n\n"
                "Unchecked documents will remain open.",
                QMessageBox.Yes | QMessageBox.Cancel,
                QMessageBox.Cancel,
            )
            if answer != QMessageBox.Yes:
                return

        failures = []
        closed = 0
        excluded = 0
        try:
            for app_name, count in selected_by_app.items():
                if count <= 0:
                    continue
                result = ProcessManager.close_office_documents(
                    app_name,
                    excluded_paths=list(self._kept_open[app_name]),
                    discard_changes=discard_changes,
                )
                closed += result.get("closed", 0)
                excluded += result.get("excluded", 0)
                failures.extend(result.get("failed", []))
        except Exception as exc:
            logging.exception("Graceful Office close failed")
            QMessageBox.critical(self, "Close Failed", str(exc))
        else:
            if failures:
                QMessageBox.warning(
                    self,
                    "Some Documents Could Not Be Closed",
                    f"Closed: {closed}\nKept open: {excluded}\nFailed: {len(failures)}",
                )
        self._refresh_stats()

    def _kill_ghosts(self, app_name):
        ProcessManager.kill_app_ghosts(app_name)
        self._refresh_stats()

    def _kill_all(self, app_name):
        ProcessManager.kill_app_all(app_name)
        self._refresh_stats()

    def _kill_all_office(self):
        answer = QMessageBox.warning(
            self,
            "Kill All Office Processes?",
            "This force-terminates Word, PowerPoint, and Visio processes. Unsaved work may be lost.",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if answer != QMessageBox.Yes:
            return
        for exe_name in self.apps.values():
            ProcessManager.kill_app_all(exe_name)
        self._refresh_stats()


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