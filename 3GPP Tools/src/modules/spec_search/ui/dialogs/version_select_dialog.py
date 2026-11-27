"""Universal master-detail specification and release selection dialog."""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QThreadPool, Qt, QTimer
from PyQt5.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton,
    QSplitter, QTableView, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from core.ui.theme import C
from modules.spec_search.core.spec_search_db import (
    SpecSearchDatabase,
    extract_major_release,
)
from modules.specifications.core.database import SpecsDatabase


from .specification_list import CatalogueQuery, ReleaseQuery, SpecificationListModel


class SpecSearchVersionSelectDialog(QDialog):
    """Universal Dialog to browse, filter, inspect dates, and batch-select releases with checkboxes."""

    def __init__(
        self,
        specs_db: SpecsDatabase,
        search_db: SpecSearchDatabase,
        cache_dir: Path,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.search_db = search_db
        self._release_request = 0
        self._catalogue_request = 0
        self._imported_entries = None
        self.cache_dir = cache_dir
        self.logger = logging.getLogger(__name__)
        self.selected_files_info: List[Dict[str, Any]] = []

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._load_specifications_list)

        self.setWindowTitle("📥 Ingest Specification Releases into Text Search DB")
        self.resize(1120, 650)
        self._setup_ui()
        self._populate_dynamic_filters()
        self._load_specifications_list()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        # ---------------------------------------------------------------------
        # Top Filter Bar: Search, Series, Working Group
        # ---------------------------------------------------------------------
        filter_bar = QHBoxLayout()

        filter_bar.addWidget(QLabel("🔍 Spec / Title:"))
        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("Search spec #, title, or topic (e.g. 38.211, 23.501, emergency, slicing)...")
        self.spec_search_input.setClearButtonEnabled(True)
        self.spec_search_input.textChanged.connect(lambda: self._search_timer.start())
        filter_bar.addWidget(self.spec_search_input, stretch=3)

        filter_bar.addWidget(QLabel("Series:"))
        self.series_combo = QComboBox()
        self.series_combo.currentIndexChanged.connect(self._load_specifications_list)
        filter_bar.addWidget(self.series_combo, stretch=1)

        filter_bar.addWidget(QLabel("WG:"))
        self.wg_combo = QComboBox()
        self.wg_combo.currentIndexChanged.connect(self._load_specifications_list)
        filter_bar.addWidget(self.wg_combo, stretch=1)

        layout.addLayout(filter_bar)

        # ---------------------------------------------------------------------
        # Main Splitter (Left: Specs List, Right: Versions List)
        # ---------------------------------------------------------------------
        splitter = QSplitter(Qt.Horizontal)

        # --- Left Panel: Specifications Browser ---
        left_group = QGroupBox("1. Select Specification(s)")
        left_layout = QVBoxLayout(left_group)
        left_layout.setContentsMargins(6, 8, 6, 6)

        self.lbl_specs_count = QLabel("Loading specifications...")
        self.lbl_specs_count.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        left_layout.addWidget(self.lbl_specs_count)

        self.specs_table = QTableView()
        self.spec_model = SpecificationListModel(self.specs_table)
        self.specs_table.setModel(self.spec_model)
        s_header = self.specs_table.horizontalHeader()
        s_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.specs_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.specs_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.specs_table.selectionModel().selectionChanged.connect(self._on_spec_selection_changed)
        left_layout.addWidget(self.specs_table)
        splitter.addWidget(left_group)

        # --- Right Panel: Versions & Ingestion Control ---
        right_group = QGroupBox("2. Select Specification Releases to Index")
        right_layout = QVBoxLayout(right_group)
        right_layout.setContentsMargins(6, 8, 6, 6)

        batch_bar = QHBoxLayout()
        self.btn_sel_unindexed = QPushButton("⚡ Select All Unindexed")
        self.btn_sel_unindexed.clicked.connect(self._select_unindexed)
        batch_bar.addWidget(self.btn_sel_unindexed)

        self.btn_sel_latest = QPushButton("⭐ Select Latest per Release")
        self.btn_sel_latest.clicked.connect(self._select_latest_per_release)
        batch_bar.addWidget(self.btn_sel_latest)

        self.btn_sel_all = QPushButton("☑️ Select All")
        self.btn_sel_all.clicked.connect(self._select_all_visible)
        batch_bar.addWidget(self.btn_sel_all)

        self.btn_desel = QPushButton("◻️ Clear")
        self.btn_desel.clicked.connect(self._deselect_all)
        batch_bar.addWidget(self.btn_desel)

        batch_bar.addStretch()
        self.lbl_selected_count = QLabel("Selected: 0 version(s)")
        self.lbl_selected_count.setStyleSheet(f"font-weight: bold; color: {C.ACCENT_SKY};")
        batch_bar.addWidget(self.lbl_selected_count)
        right_layout.addLayout(batch_bar)

        self.versions_table = QTableWidget()
        self.versions_table.setColumnCount(7)
        self.versions_table.setHorizontalHeaderLabels([
            "Select", "Spec", "Version", "Release Date", "Filename", "Local Cache", "Search DB Status"
        ])
        v_header = self.versions_table.horizontalHeader()
        v_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        v_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        v_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        v_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        v_header.setSectionResizeMode(4, QHeaderView.Stretch)
        v_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        v_header.setSectionResizeMode(6, QHeaderView.ResizeToContents)

        self.versions_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.versions_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.versions_table.itemChanged.connect(self._on_version_item_changed)
        self.versions_table.itemDoubleClicked.connect(self._on_version_double_clicked)
        right_layout.addWidget(self.versions_table)
        splitter.addWidget(right_group)

        splitter.setSizes([400, 680])
        layout.addWidget(splitter)

        # ---------------------------------------------------------------------
        # Dialog Action Buttons
        # ---------------------------------------------------------------------
        self.chk_auto_supersede = QCheckBox("Auto-supersede older minor versions within the same release (recommended)")
        self.chk_auto_supersede.setChecked(True)
        self.chk_auto_supersede.setToolTip("When checked, ingesting a newer release (e.g. 15.17.0) will automatically remove any older indexed version of that release (e.g. 15.16.0).")
        layout.addWidget(self.chk_auto_supersede)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.ok_btn = btn_box.button(QDialogButtonBox.Ok)
        self.ok_btn.setText("📥 Ingest Selected Releases")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    # -------------------------------------------------------------------------
    # Data Population & Filtering
    # -------------------------------------------------------------------------

    def _populate_dynamic_filters(self):
        """Loads all actual Series and Working Groups dynamically from the database."""
        options = self.specs_db.get_filter_options()

        self.series_combo.blockSignals(True)
        self.series_combo.clear()
        self.series_combo.addItem("All Series", "")
        for s in options.get("series", []):
            self.series_combo.addItem(f"Series {s}", s)
        self.series_combo.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.clear()
        self.wg_combo.addItem("All Working Groups", "")
        for g in options.get("groups", []):
            self.wg_combo.addItem(g, g)
        self.wg_combo.blockSignals(False)

    def _load_specifications_list(self):
        """Queries specifications matching active filters and populates the left table."""
        search_kw = self.spec_search_input.text().strip()
        series_filter = self.series_combo.currentData()
        wg_filter = self.wg_combo.currentData()

        self._catalogue_request += 1
        worker = CatalogueQuery(self.specs_db, self._catalogue_request, (search_kw, series_filter, wg_filter))
        worker.signals.loaded.connect(self._show_catalogue, Qt.QueuedConnection)
        QThreadPool.globalInstance().start(worker)

    def _show_catalogue(self, request_id, spec_rows):
        if request_id != self._catalogue_request:
            return
        self.spec_model.replace(spec_rows)
        self.lbl_specs_count.setText(f"Found {len(spec_rows)} specification(s)")

        if len(spec_rows) > 0:
            self.specs_table.selectRow(0)
        else:
            self._load_versions_for_specs([])

    def _on_spec_selection_changed(self):
        """Fires when user selects one or more specifications in the left table."""
        rows = self.specs_table.selectionModel().selectedRows()
        if not rows:
            self._load_versions_for_specs([])
            return
        selected_spec_numbers = [str(index.data(Qt.UserRole)) for index in sorted(rows, key=lambda i: i.row())]
        self._load_versions_for_specs(selected_spec_numbers)

    def _load_versions_for_specs(self, spec_numbers: List[str]):
        """Populates the right table with all releases for the selected specification(s)."""
        self._release_request += 1
        self.versions_table.setRowCount(0)
        self._update_selected_count()
        if not spec_numbers:
            return
        worker = ReleaseQuery(self.specs_db, self.search_db, self._release_request, spec_numbers, self._imported_entries)
        worker.signals.loaded.connect(self._show_releases, Qt.QueuedConnection)
        QThreadPool.globalInstance().start(worker)

    def _show_releases(self, request_id, result):
        if request_id != self._release_request:
            return
        self._imported_entries, sorted_files = result
        self.versions_table.setUpdatesEnabled(False)
        self.versions_table.blockSignals(True)
        self.versions_table.clearContents()
        self.versions_table.setRowCount(len(sorted_files))

        for row_idx, (row_data, cached_file) in enumerate(sorted_files):
            s_num, title, sp_type, filename, version, url, upload_date = row_data
            display_type = sp_type if sp_type else "TS"

            task_data = {
                "spec_number": s_num,
                "version": version,
                "filename": filename,
                "file_url": url,
                "release_date": upload_date or "",
            }

            # Column 0: Checkbox Selector
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            chk_item.setCheckState(Qt.Unchecked)
            chk_item.setData(Qt.UserRole, task_data)
            self.versions_table.setItem(row_idx, 0, chk_item)

            # Column 1: Spec Number
            s_item = QTableWidgetItem(f"{display_type} {s_num}")
            self.versions_table.setItem(row_idx, 1, s_item)

            # Column 2: Version
            v_item = QTableWidgetItem(f"v{version}")
            self.versions_table.setItem(row_idx, 2, v_item)

            # Column 3: Release Date
            date_str = str(upload_date) if upload_date else "-"
            d_item = QTableWidgetItem(date_str)
            if not upload_date:
                d_item.setForeground(Qt.gray)
            self.versions_table.setItem(row_idx, 3, d_item)

            # Column 4: Filename
            f_item = QTableWidgetItem(filename)
            f_item.setToolTip(f"{title}\n{filename}")
            self.versions_table.setItem(row_idx, 4, f_item)

            # Column 5: Cache status
            if cached_file:
                c_item = QTableWidgetItem(f"🟢 Cached ({cached_file.suffix[1:].upper()})")
                c_item.setForeground(Qt.darkGreen)
            else:
                c_item = QTableWidgetItem("🌐 Remote (FTP)")
                c_item.setForeground(Qt.darkGray)
            self.versions_table.setItem(row_idx, 5, c_item)

            # Column 6: Search DB Status
            in_db = (s_num, version) in self._imported_entries
            db_item = QTableWidgetItem("✅ Indexed" if in_db else "⚪ Ready")
            if in_db:
                db_item.setForeground(Qt.blue)
            self.versions_table.setItem(row_idx, 6, db_item)

        self.versions_table.setUpdatesEnabled(True)
        self.versions_table.blockSignals(False)
        self._update_selected_count()

    # -------------------------------------------------------------------------
    # Batch Selection Actions (Operates on Checkboxes)
    # -------------------------------------------------------------------------

    def _on_version_item_changed(self, item: QTableWidgetItem):
        if item.column() == 0:
            self._update_selected_count()

    def _update_selected_count(self):
        checked_count = sum(
            1
            for r in range(self.versions_table.rowCount())
            if self.versions_table.item(r, 0) and self.versions_table.item(r, 0).checkState() == Qt.Checked
        )
        self.lbl_selected_count.setText(f"Selected: {checked_count} version(s)")

    def _select_unindexed(self):
        """Checks all versions that are not yet indexed into the Search DB."""
        self.versions_table.blockSignals(True)
        for r in range(self.versions_table.rowCount()):
            status_item = self.versions_table.item(r, 6)
            chk_item = self.versions_table.item(r, 0)
            if chk_item and status_item:
                if "Ready" in status_item.text():
                    chk_item.setCheckState(Qt.Checked)
                else:
                    chk_item.setCheckState(Qt.Unchecked)
        self.versions_table.blockSignals(False)
        self._update_selected_count()

    def _select_latest_per_release(self):
        """Checks the newest version for each major Release (Rel-15, Rel-16, Rel-17, Rel-18, etc.)."""
        self.versions_table.blockSignals(True)
        seen_major_releases = set()

        for r in range(self.versions_table.rowCount()):
            chk_item = self.versions_table.item(r, 0)
            s_item = self.versions_table.item(r, 1)
            v_item = self.versions_table.item(r, 2)
            if not chk_item or not s_item or not v_item:
                continue

            spec_num = s_item.text().replace("TS", "").replace("TR", "").strip()
            ver_str = v_item.text().lstrip("v").strip()
            rel_key = extract_major_release(ver_str)
            group_key = (spec_num, rel_key)

            if group_key not in seen_major_releases:
                seen_major_releases.add(group_key)
                chk_item.setCheckState(Qt.Checked)
            else:
                chk_item.setCheckState(Qt.Unchecked)

        self.versions_table.blockSignals(False)
        self._update_selected_count()

    def _select_all_visible(self):
        """Checks all currently visible rows."""
        self.versions_table.blockSignals(True)
        for r in range(self.versions_table.rowCount()):
            chk_item = self.versions_table.item(r, 0)
            if chk_item:
                chk_item.setCheckState(Qt.Checked)
        self.versions_table.blockSignals(False)
        self._update_selected_count()

    def _deselect_all(self):
        """Unchecks all rows."""
        self.versions_table.blockSignals(True)
        for r in range(self.versions_table.rowCount()):
            chk_item = self.versions_table.item(r, 0)
            if chk_item:
                chk_item.setCheckState(Qt.Unchecked)
        self.versions_table.blockSignals(False)
        self._update_selected_count()

    def _on_version_double_clicked(self, item: QTableWidgetItem):
        """Double clicking a row immediately imports that single version."""
        row = item.row()
        chk_item = self.versions_table.item(row, 0)
        if chk_item:
            self.selected_files_info = [chk_item.data(Qt.UserRole)]
            self.accept()

    def _on_accept(self):
        """Collects all checked rows and submits them for ingestion."""
        selected_tasks = []
        for r in range(self.versions_table.rowCount()):
            chk_item = self.versions_table.item(r, 0)
            if chk_item and chk_item.checkState() == Qt.Checked:
                data = chk_item.data(Qt.UserRole)
                if data:
                    selected_tasks.append(data)

        if not selected_tasks:
            QMessageBox.warning(self, "Selection Required", "Please check at least one specification version to index.")
            return

        self.auto_replace_older = self.chk_auto_supersede.isChecked()
        self.selected_files_info = selected_tasks
        self.accept()
