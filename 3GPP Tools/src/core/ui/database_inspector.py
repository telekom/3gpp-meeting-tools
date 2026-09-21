import sqlite3
from pathlib import Path
from typing import Any, List, Optional, Sequence, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class DatabaseInspectorDialog(QDialog):
    """Read-only SQLite browser and diagnostics dialog."""

    PAGE_SIZE = 200

    def __init__(self, db_path: Path, display_name: str = "", parent=None):
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.display_name = display_name or self.db_path.name
        self._table_names: List[str] = []
        self._current_page = 0
        self._current_total = 0

        self.setWindowTitle(f"Database Inspector — {self.display_name}")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(1050, 700)
        self.setStyleSheet("background-color: #FAFAFA;")

        self._setup_ui()
        self._reload_metadata()

    @staticmethod
    def _quote_identifier(identifier: str) -> str:
        """Safely quote an SQLite identifier after it came from SQLite metadata."""
        return '"' + identifier.replace('"', '""') + '"'

    def _connect_read_only(self) -> sqlite3.Connection:
        """
        Open the database through SQLite's read-only URI mode.

        mode=ro prevents the inspector from creating or modifying database files.
        query_only provides a second guard against accidental writes.
        """
        uri = self.db_path.resolve().as_uri() + "?mode=ro"
        conn = sqlite3.connect(uri, uri=True, timeout=5.0)
        conn.execute("PRAGMA query_only = ON;")
        return conn

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        title = QLabel(f"🗄️ {self.display_name}")
        title.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        path_label = QLabel(str(self.db_path))
        path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        path_label.setStyleSheet("color: #64748B; font-size: 10px; margin-bottom: 4px;")
        layout.addWidget(path_label)

        notice = QLabel(
            "Read-only inspector. Browsing and diagnostics never modify database contents."
        )
        notice.setStyleSheet("color: #166534; font-size: 10px; margin-bottom: 4px;")
        layout.addWidget(notice)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_data_tab(), "Data")
        self.tabs.addTab(self._build_schema_tab(), "Schema")
        self.tabs.addTab(self._build_diagnostics_tab(), "Diagnostics")
        self.tabs.currentChanged.connect(self._tab_changed)
        layout.addWidget(self.tabs)

        buttons = QHBoxLayout()
        refresh = QPushButton("🔄 Refresh")
        refresh.clicked.connect(self._reload_metadata)
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        buttons.addWidget(refresh)
        buttons.addStretch()
        buttons.addWidget(close)
        layout.addLayout(buttons)

    def _build_data_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(6, 8, 6, 6)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Table:"))
        self.table_combo = QComboBox()
        self.table_combo.setMinimumWidth(220)
        self.table_combo.currentIndexChanged.connect(self._table_changed)
        controls.addWidget(self.table_combo)

        controls.addWidget(QLabel("Filter:"))
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Search all displayed columns…")
        self.filter_edit.returnPressed.connect(self._filter_changed)
        controls.addWidget(self.filter_edit, 1)

        apply_filter = QPushButton("Apply")
        apply_filter.clicked.connect(self._filter_changed)
        controls.addWidget(apply_filter)

        clear_filter = QPushButton("Clear")
        clear_filter.clicked.connect(self._clear_filter)
        controls.addWidget(clear_filter)
        layout.addLayout(controls)

        self.data_table = QTableWidget()
        self.data_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.data_table.setSelectionBehavior(QTableWidget.SelectItems)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSortingEnabled(False)
        self.data_table.verticalHeader().setDefaultSectionSize(23)
        self.data_table.horizontalHeader().setStretchLastSection(False)
        layout.addWidget(self.data_table, 1)

        pager = QHBoxLayout()
        self.prev_btn = QPushButton("◀ Previous")
        self.prev_btn.clicked.connect(self._previous_page)
        self.next_btn = QPushButton("Next ▶")
        self.next_btn.clicked.connect(self._next_page)
        self.page_label = QLabel("")
        self.page_label.setAlignment(Qt.AlignCenter)
        pager.addWidget(self.prev_btn)
        pager.addStretch()
        pager.addWidget(self.page_label)
        pager.addStretch()
        pager.addWidget(self.next_btn)
        layout.addLayout(pager)
        return tab

    def _build_schema_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(6, 8, 6, 6)

        top = QHBoxLayout()
        top.addWidget(QLabel("Table:"))
        self.schema_combo = QComboBox()
        self.schema_combo.setMinimumWidth(220)
        self.schema_combo.currentIndexChanged.connect(self._load_schema)
        top.addWidget(self.schema_combo)
        top.addStretch()
        layout.addLayout(top)

        self.schema_table = QTableWidget()
        self.schema_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.schema_table.setAlternatingRowColors(True)
        layout.addWidget(self.schema_table, 1)

        self.index_table = QTableWidget()
        self.index_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.index_table.setAlternatingRowColors(True)
        layout.addWidget(QLabel("Indexes"))
        layout.addWidget(self.index_table, 1)
        return tab

    def _build_diagnostics_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(6, 8, 6, 6)

        run = QPushButton("▶ Run Diagnostics")
        run.setFixedWidth(150)
        run.clicked.connect(self._run_diagnostics)
        layout.addWidget(run)

        self.diagnostics_text = QTextEdit()
        self.diagnostics_text.setReadOnly(True)
        self.diagnostics_text.setFont(QFont("Consolas"))
        self.diagnostics_text.setPlaceholderText(
            "Run diagnostics to check SQLite integrity and application data consistency."
        )
        layout.addWidget(self.diagnostics_text, 1)
        return tab

    def _reload_metadata(self):
        if not self.db_path.exists():
            QMessageBox.warning(self, "Database Missing", f"Database does not exist:\n{self.db_path}")
            return

        try:
            with self._connect_read_only() as conn:
                rows = conn.execute(
                    """
                    SELECT name
                    FROM sqlite_master
                    WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
                    ORDER BY name COLLATE NOCASE
                    """
                ).fetchall()
            self._table_names = [row[0] for row in rows]
        except sqlite3.Error as exc:
            QMessageBox.critical(self, "Database Error", f"Could not inspect {self.db_path.name}:\n{exc}")
            return

        current_data = self.table_combo.currentText()
        current_schema = self.schema_combo.currentText()

        self.table_combo.blockSignals(True)
        self.schema_combo.blockSignals(True)
        self.table_combo.clear()
        self.schema_combo.clear()
        self.table_combo.addItems(self._table_names)
        self.schema_combo.addItems(self._table_names)

        if current_data in self._table_names:
            self.table_combo.setCurrentText(current_data)
        if current_schema in self._table_names:
            self.schema_combo.setCurrentText(current_schema)

        self.table_combo.blockSignals(False)
        self.schema_combo.blockSignals(False)

        self._current_page = 0
        self._load_data()
        self._load_schema()

    def _table_changed(self):
        self._current_page = 0
        self._load_data()

    def _filter_changed(self):
        self._current_page = 0
        self._load_data()

    def _clear_filter(self):
        self.filter_edit.clear()
        self._filter_changed()

    def _column_names(self, conn: sqlite3.Connection, table_name: str) -> List[str]:
        quoted = self._quote_identifier(table_name)
        return [row[1] for row in conn.execute(f"PRAGMA table_info({quoted})").fetchall()]

    def _build_filter(
        self, columns: Sequence[str], filter_text: str
    ) -> Tuple[str, List[Any]]:
        text = filter_text.strip()
        if not text or not columns:
            return "", []

        predicates = [
            f"CAST({self._quote_identifier(column)} AS TEXT) LIKE ?"
            for column in columns
        ]
        value = f"%{text}%"
        return " WHERE " + " OR ".join(predicates), [value] * len(columns)

    def _load_data(self):
        table_name = self.table_combo.currentText()
        if not table_name:
            self.data_table.clear()
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            self.page_label.setText("No tables")
            self.prev_btn.setEnabled(False)
            self.next_btn.setEnabled(False)
            return

        try:
            with self._connect_read_only() as conn:
                columns = self._column_names(conn, table_name)
                where_sql, params = self._build_filter(
                    columns, self.filter_edit.text()
                )
                quoted_table = self._quote_identifier(table_name)
                self._current_total = conn.execute(
                    f"SELECT COUNT(*) FROM {quoted_table}{where_sql}", params
                ).fetchone()[0]

                max_page = max(0, (self._current_total - 1) // self.PAGE_SIZE)
                self._current_page = min(self._current_page, max_page)
                offset = self._current_page * self.PAGE_SIZE
                rows = conn.execute(
                    f"SELECT * FROM {quoted_table}{where_sql} "
                    f"LIMIT ? OFFSET ?",
                    [*params, self.PAGE_SIZE, offset],
                ).fetchall()
        except sqlite3.Error as exc:
            QMessageBox.warning(self, "Query Error", f"Could not read table {table_name}:\n{exc}")
            return

        self.data_table.setUpdatesEnabled(False)
        try:
            self.data_table.clear()
            self.data_table.setColumnCount(len(columns))
            self.data_table.setHorizontalHeaderLabels(columns)
            self.data_table.setRowCount(len(rows))

            null_brush = QBrush(QColor("#94A3B8"))
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    if value is None:
                        item = QTableWidgetItem("NULL")
                        item.setForeground(null_brush)
                        font = item.font()
                        font.setItalic(True)
                        item.setFont(font)
                    elif isinstance(value, bytes):
                        item = QTableWidgetItem(f"<BLOB: {len(value)} bytes>")
                        item.setToolTip("Binary value")
                    elif value == "":
                        item = QTableWidgetItem('""')
                        item.setForeground(QBrush(QColor("#B45309")))
                        item.setToolTip("Empty string")
                    else:
                        item = QTableWidgetItem(str(value))
                    self.data_table.setItem(row_index, column_index, item)

            self.data_table.resizeColumnsToContents()
            for column in range(self.data_table.columnCount()):
                if self.data_table.columnWidth(column) > 350:
                    self.data_table.setColumnWidth(column, 350)
        finally:
            self.data_table.setUpdatesEnabled(True)

        start = self._current_page * self.PAGE_SIZE + 1 if self._current_total else 0
        end = min((self._current_page + 1) * self.PAGE_SIZE, self._current_total)
        self.page_label.setText(
            f"Rows {start:,}–{end:,} of {self._current_total:,} "
            f"(page {self._current_page + 1:,})"
        )
        self.prev_btn.setEnabled(self._current_page > 0)
        self.next_btn.setEnabled(end < self._current_total)

    def _previous_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._load_data()

    def _next_page(self):
        if (self._current_page + 1) * self.PAGE_SIZE < self._current_total:
            self._current_page += 1
            self._load_data()

    def _load_schema(self):
        table_name = self.schema_combo.currentText()
        if not table_name:
            self.schema_table.clear()
            self.index_table.clear()
            return

        quoted = self._quote_identifier(table_name)
        try:
            with self._connect_read_only() as conn:
                columns = conn.execute(f"PRAGMA table_info({quoted})").fetchall()
                indexes = conn.execute(f"PRAGMA index_list({quoted})").fetchall()
        except sqlite3.Error as exc:
            QMessageBox.warning(self, "Schema Error", f"Could not inspect schema:\n{exc}")
            return

        self.schema_table.clear()
        self.schema_table.setColumnCount(6)
        self.schema_table.setHorizontalHeaderLabels(
            ["CID", "Column", "Type", "NOT NULL", "Default", "Primary Key"]
        )
        self.schema_table.setRowCount(len(columns))
        for r, row in enumerate(columns):
            values = [row[0], row[1], row[2], bool(row[3]), row[4], bool(row[5])]
            for c, value in enumerate(values):
                self.schema_table.setItem(
                    r, c, QTableWidgetItem("NULL" if value is None else str(value))
                )
        self.schema_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.schema_table.resizeColumnsToContents()

        self.index_table.clear()
        self.index_table.setColumnCount(5)
        self.index_table.setHorizontalHeaderLabels(
            ["Sequence", "Name", "Unique", "Origin", "Partial"]
        )
        self.index_table.setRowCount(len(indexes))
        for r, row in enumerate(indexes):
            values = row[:5]
            for c, value in enumerate(values):
                self.index_table.setItem(r, c, QTableWidgetItem(str(value)))
        self.index_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.index_table.resizeColumnsToContents()

    def _tab_changed(self, index: int):
        if self.tabs.tabText(index) == "Schema":
            self._load_schema()

    def _run_diagnostics(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            lines = [
                f"Database: {self.db_path.name}",
                f"Path:     {self.db_path}",
                "",
                "SQLite diagnostics",
                "────────────────────────────────────────────────────────",
            ]

            with self._connect_read_only() as conn:
                integrity_rows = conn.execute("PRAGMA integrity_check;").fetchall()
                integrity = ", ".join(str(row[0]) for row in integrity_rows)
                lines.append(
                    f"Integrity check:              {'✓ OK' if integrity.lower() == 'ok' else '⚠ ' + integrity}"
                )

                fk_rows = conn.execute("PRAGMA foreign_key_check;").fetchall()
                lines.append(
                    f"Foreign-key violations:      {len(fk_rows):,}"
                    + (" ✓" if not fk_rows else " ⚠")
                )

                tables = {
                    row[0]
                    for row in conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
                lines.append(f"Application tables:          {len(self._table_names):,}")

                lines.extend(["", "Table row counts", "────────────────────────────────────────────────────────"])
                for table_name in self._table_names:
                    quoted = self._quote_identifier(table_name)
                    count = conn.execute(f"SELECT COUNT(*) FROM {quoted}").fetchone()[0]
                    lines.append(f"{table_name:<30} {count:>12,}")

                # Application-aware checks are additive and only run when the
                # expected table/columns actually exist.
                if "files" in tables:
                    file_columns = set(self._column_names(conn, "files"))
                    checks = []
                    if "version" in file_columns:
                        empty_versions = conn.execute(
                            """
                            SELECT COUNT(*) FROM files
                            WHERE version IS NULL OR TRIM(version) = ''
                            """
                        ).fetchone()[0]
                        checks.append(
                            ("Files with empty version", empty_versions)
                        )
                    if "url" in file_columns:
                        empty_urls = conn.execute(
                            """
                            SELECT COUNT(*) FROM files
                            WHERE url IS NULL OR TRIM(url) = ''
                            """
                        ).fetchone()[0]
                        checks.append(("Files with empty URL", empty_urls))

                    if checks:
                        lines.extend([
                            "",
                            "Specification data checks",
                            "────────────────────────────────────────────────────────",
                        ])
                        for label, count in checks:
                            marker = "✓" if count == 0 else "⚠"
                            lines.append(f"{label + ':':<30} {count:>12,} {marker}")

            self.diagnostics_text.setPlainText("\n".join(lines))
        except sqlite3.Error as exc:
            self.diagnostics_text.setPlainText(f"Diagnostics failed:\n{exc}")
        finally:
            QApplication.restoreOverrideCursor()
