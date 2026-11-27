"""Modeless SQLite Database Maintenance & Compaction Hub dialog."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QVBoxLayout,
)

from core.ui.database_inspector import DatabaseInspectorDialog

from .actions_mixin import _MaintenanceActionsMixin
from .storage_ops import (
    get_tracked_databases,
    query_db_metrics,
    vacuum_database_file,
)
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

from .styles import (
    DIALOG_BG_STYLE,
    HEADER_DESC_STYLE,
    HEADER_TITLE_STYLE,
    PROGRESS_BAR_STYLE,
    STATUS_LABEL_STYLE,
    STORAGE_TABLE_STYLE,
    TOP_BTN_STYLE,
    VACUUM_ALL_BTN_STYLE,
)
from .table_mixin import _MaintenanceTableMixin
from .widgets import KpiCard

logger = logging.getLogger(__name__)


class DatabaseMaintenanceDialog(
    _MaintenanceTableMixin,
    _MaintenanceActionsMixin,
    QDialog,
):
    """
    Executive database maintenance, storage KPI monitoring, and compaction hub
    for all SQLite databases managed by 3GPP Tools.

    Configured as a non-blocking modeless floating window (Qt.Window, Qt.NonModal)
    without Qt.WindowStaysOnTopHint so delegates can freely navigate the rest
    of the application and switch between external desktop tools.
    """

    log_msg = pyqtSignal(str, int)

    def __init__(self, parent=None):
        # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top)
        super().__init__(None)
        self._parent_ref = parent
        self.setWindowTitle("Database Maintenance & Compaction Hub")
        # Standard window without WindowStaysOnTopHint so it doesn't force itself over external apps
        self.setWindowFlags(Qt.Window)
        # Non-modal so it never blocks the main window or other dialogs
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.resize(1020, 560)
        self.setMinimumSize(940, 480)
        self.setStyleSheet(DIALOG_BG_STYLE)

        self._inspector_dialogs: Dict[str, DatabaseInspectorDialog] = {}
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
        title.setStyleSheet(HEADER_TITLE_STYLE)
        header_text_layout.addWidget(title)

        desc = QLabel(
            "Inspect table schemas, calculate reclaimable freelist pages, and run full VACUUM and WAL checkpoints "
            "to defragment indices and reclaim disk space."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(HEADER_DESC_STYLE)
        header_text_layout.addWidget(desc)

        header_layout.addLayout(header_text_layout, 1)

        open_folder_btn = QPushButton("📂 Open DB Folder")
        open_folder_btn.setFixedHeight(28)
        open_folder_btn.setStyleSheet(TOP_BTN_STYLE)
        open_folder_btn.setToolTip("Open the SQLite databases directory in Windows File Explorer.")
        open_folder_btn.clicked.connect(self._open_db_folder)
        header_layout.addWidget(open_folder_btn)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.setFixedHeight(28)
        refresh_btn.setStyleSheet(TOP_BTN_STYLE)
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
        # Generous actions column width so Inspect, Compact, Wipe never truncate on high-DPI
        self.table.setColumnWidth(5, 280)

        self.table.verticalHeader().setDefaultSectionSize(48)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setShowGrid(False)
        self.table.setStyleSheet(STORAGE_TABLE_STYLE)
        root_layout.addWidget(self.table, 1)

        # ----------------------------------------------------
        # 4. FOOTER CONTROLS & COMPACTION PROGRESS
        # ----------------------------------------------------
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(10)

        self.vacuum_all_btn = QPushButton("🧹 Compact All Databases")
        self.vacuum_all_btn.setFixedHeight(30)
        self.vacuum_all_btn.setStyleSheet(VACUUM_ALL_BTN_STYLE)
        self.vacuum_all_btn.clicked.connect(self._vacuum_all)
        footer_layout.addWidget(self.vacuum_all_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.setFixedWidth(180)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)
        self.progress_bar.setVisible(False)
        footer_layout.addWidget(self.progress_bar)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet(STATUS_LABEL_STYLE)
        self.status_lbl.setVisible(False)
        footer_layout.addWidget(self.status_lbl)

        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        root_layout.addLayout(footer_layout)

    def _get_tracked_databases(self) -> List[Dict[str, Any]]:
        """Delegate database discovery while preserving subclass/monkeypatch hooks."""
        return get_tracked_databases()

    def _query_db_metrics(self, db_path: Path) -> Dict[str, Any]:
        """Delegate read-only SQLite metrics inspection."""
        return query_db_metrics(db_path)

    def _vacuum_database_file(self, db_path: Path) -> Tuple[bool, int, int, str]:
        """Delegate WAL checkpointing and VACUUM on a single database file."""
        return vacuum_database_file(db_path)
