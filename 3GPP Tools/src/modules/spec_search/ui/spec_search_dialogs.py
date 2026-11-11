"""
Specification and Release Selection Dialog.
Provides universal master-detail browsing, dynamic series & WG filtering,
spec title search, explicit row checkboxes, and reliable batch selection.
"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QFrame,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_threads import find_cached_spec_file
from modules.spec_search.core.spec_clause_diff import build_llm_clause_prompt, generate_unified_diff
from modules.spec_search.core.spec_search_db import (
    SpecSearchDatabase,
    extract_major_release,
    parse_version_tuple,
)
from modules.specifications.core.database import SpecsDatabase


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
        self.lbl_specs_count.setStyleSheet("color: #475569; font-size: 11px;")
        left_layout.addWidget(self.lbl_specs_count)

        self.specs_table = QTableWidget()
        self.specs_table.setColumnCount(4)
        self.specs_table.setHorizontalHeaderLabels(["Spec", "WG", "Releases", "Title"])
        s_header = self.specs_table.horizontalHeader()
        s_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.specs_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.specs_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.specs_table.itemSelectionChanged.connect(self._on_spec_selection_changed)
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
        self.lbl_selected_count.setStyleSheet("font-weight: bold; color: #0284C7;")
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

        query = """
            SELECT sp.number, sp.title, sp.type, s.name AS series_name, p_grp.name AS wg_name, COUNT(f.id) AS file_count
            FROM specifications sp
            JOIN series s ON sp.series_id = s.id
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN files f ON f.spec_id = sp.id
            WHERE 1=1
        """
        params: List[Any] = []

        if series_filter:
            query += " AND s.name = ?"
            params.append(str(series_filter))

        if wg_filter:
            query += " AND p_grp.name = ?"
            params.append(str(wg_filter))

        if search_kw:
            search_pat = f"%{search_kw}%"
            query += " AND (sp.number LIKE ? OR sp.title LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR ua.tags LIKE ? OR ua.comments LIKE ?)"
            params.extend([search_pat, search_pat, search_pat, search_pat, search_pat])

        query += " GROUP BY sp.id HAVING file_count > 0 ORDER BY CAST(s.name AS INTEGER) ASC, sp.number ASC"

        try:
            with self.specs_db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                spec_rows = cursor.fetchall()
        except Exception as e:
            self.logger.error(f"Error fetching specs summary: {e}")
            spec_rows = []

        self.specs_table.setUpdatesEnabled(False)
        self.specs_table.blockSignals(True)
        self.specs_table.clearContents()
        self.specs_table.setRowCount(len(spec_rows))

        for r_idx, row in enumerate(spec_rows):
            spec_num, title, sp_type, series_name, wg_name, file_count = row
            display_type = sp_type if sp_type else "TS"

            # Spec Number
            s_item = QTableWidgetItem(f"{display_type} {spec_num}")
            s_item.setData(Qt.UserRole, spec_num)
            self.specs_table.setItem(r_idx, 0, s_item)

            # Working Group
            wg_item = QTableWidgetItem(wg_name or "-")
            self.specs_table.setItem(r_idx, 1, wg_item)

            # Releases count
            rel_item = QTableWidgetItem(str(file_count))
            rel_item.setTextAlignment(Qt.AlignCenter)
            self.specs_table.setItem(r_idx, 2, rel_item)

            # Title
            t_item = QTableWidgetItem(title or "")
            t_item.setToolTip(f"{display_type} {spec_num}: {title}")
            self.specs_table.setItem(r_idx, 3, t_item)

        self.specs_table.setUpdatesEnabled(True)
        self.specs_table.blockSignals(False)
        self.lbl_specs_count.setText(f"Found {len(spec_rows)} specification(s)")

        if len(spec_rows) > 0:
            self.specs_table.selectRow(0)
        else:
            self.versions_table.setRowCount(0)
            self._update_selected_count()

    def _on_spec_selection_changed(self):
        """Fires when user selects one or more specifications in the left table."""
        selected_spec_rows = sorted(list(set(item.row() for item in self.specs_table.selectedItems())))
        if not selected_spec_rows:
            self.versions_table.setRowCount(0)
            self._update_selected_count()
            return

        selected_spec_numbers = [
            str(self.specs_table.item(r, 0).data(Qt.UserRole))
            for r in selected_spec_rows
            if self.specs_table.item(r, 0)
        ]

        self._load_versions_for_specs(selected_spec_numbers)

    def _load_versions_for_specs(self, spec_numbers: List[str]):
        """Populates the right table with all releases for the selected specification(s)."""
        if not spec_numbers:
            self.versions_table.setRowCount(0)
            self._update_selected_count()
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            placeholders = ",".join("?" for _ in spec_numbers)
            query = f"""
                SELECT sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date
                FROM files f
                JOIN specifications sp ON f.spec_id = sp.id
                WHERE sp.number IN ({placeholders})
            """
            with self.specs_db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, spec_numbers)
                file_rows = cursor.fetchall()

            imported_entries = {
                (v["spec_number"], v["version"]) for v in self.search_db.get_imported_versions()
            }

            # Stable two-pass sort: version descending first, then spec number ascending
            sorted_files = sorted(file_rows, key=lambda r: parse_version_tuple(r[4]), reverse=True)
            sorted_files.sort(key=lambda r: r[0])

            self.versions_table.setUpdatesEnabled(False)
            self.versions_table.blockSignals(True)
            self.versions_table.clearContents()
            self.versions_table.setRowCount(len(sorted_files))

            for row_idx, row_data in enumerate(sorted_files):
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
                cached_file = find_cached_spec_file(filename, s_num)
                if cached_file:
                    c_item = QTableWidgetItem(f"🟢 Cached ({cached_file.suffix[1:].upper()})")
                    c_item.setForeground(Qt.darkGreen)
                else:
                    c_item = QTableWidgetItem("🌐 Remote (FTP)")
                    c_item.setForeground(Qt.darkGray)
                self.versions_table.setItem(row_idx, 5, c_item)

                # Column 6: Search DB Status
                in_db = (s_num, version) in imported_entries
                db_item = QTableWidgetItem("✅ Indexed" if in_db else "⚪ Ready")
                if in_db:
                    db_item.setForeground(Qt.blue)
                self.versions_table.setItem(row_idx, 6, db_item)

            self.versions_table.setUpdatesEnabled(True)
            self.versions_table.blockSignals(False)
            self._update_selected_count()

        finally:
            QApplication.restoreOverrideCursor()

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





class SpecSearchPruneDialog(QDialog):
    """
    Dialog to review and prune older specification versions superseded by newer releases
    within the same 3GPP major release.
    """

    def __init__(
        self,
        db: SpecSearchDatabase,
        superseded_versions: List[Dict[str, Any]],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.superseded_versions = superseded_versions
        self.pruned_count = 0
        self.setWindowTitle("🧹 Prune Superseded Specification Versions")
        self.resize(780, 500)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Header Info Banner
        info_banner = QFrame()
        info_banner.setStyleSheet("""
            QFrame {
                background-color: #EFF6FF;
                border: 1px solid #BFDBFE;
                border-radius: 6px;
                padding: 8px;
            }
        """)
        ib_layout = QVBoxLayout(info_banner)
        ib_layout.setContentsMargins(6, 6, 6, 6)
        ib_layout.setSpacing(3)

        title_lbl = QLabel(f"<b>Found {len(self.superseded_versions)} Superseded Specification Version(s)</b>")
        title_lbl.setStyleSheet("color: #1E3A8A; font-size: 12px;")
        desc_lbl = QLabel(
            "These older minor releases belong to the same 3GPP Release as a newer indexed version.\n"
            "Pruning them will remove their clause records and FTS indices, freeing database space and cleaning the evolution matrix."
        )
        desc_lbl.setStyleSheet("color: #1E40AF; font-size: 11px;")
        desc_lbl.setWordWrap(True)
        ib_layout.addWidget(title_lbl)
        ib_layout.addWidget(desc_lbl)
        layout.addWidget(info_banner)

        # Quick selection bar
        sel_bar = QHBoxLayout()
        btn_sel_all = QPushButton("Select All")
        btn_sel_all.clicked.connect(self._select_all)
        sel_bar.addWidget(btn_sel_all)

        btn_desel_all = QPushButton("Deselect All")
        btn_desel_all.clicked.connect(self._deselect_all)
        sel_bar.addWidget(btn_desel_all)

        sel_bar.addStretch()
        self.lbl_selected = QLabel(f"Selected: {len(self.superseded_versions)} version(s)")
        self.lbl_selected.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 11px;")
        sel_bar.addWidget(self.lbl_selected)
        layout.addLayout(sel_bar)

        # Table of superseded versions
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Prune", "Specification", "Superseded Version", "Release", "Active Replacement", "Clauses"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)

        self._populate_table()
        layout.addWidget(self.table)

        # Action Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_cancel)

        self.btn_prune = QPushButton(f"🧹 Prune Selected ({len(self.superseded_versions)})")
        self.btn_prune.setStyleSheet("""
            QPushButton {
                background-color: #DC2626;
                color: #FFFFFF;
                border: none;
                border-radius: 4px;
                padding: 6px 14px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #B91C1C;
            }
            QPushButton:disabled {
                background-color: #CBD5E1;
                color: #94A3B8;
            }
        """)
        self.btn_prune.clicked.connect(self._on_prune_confirmed)
        btn_box.addWidget(self.btn_prune)
        layout.addLayout(btn_box)

    def _populate_table(self):
        self.table.setRowCount(len(self.superseded_versions))
        for r, item in enumerate(self.superseded_versions):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            chk.setCheckState(Qt.Checked)
            self.table.setItem(r, 0, chk)

            s_item = QTableWidgetItem(f"TS {item['spec_number']}")
            s_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 1, s_item)

            v_item = QTableWidgetItem(f"v{item['version']}")
            v_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 2, v_item)

            rel_item = QTableWidgetItem(str(item.get("release", "")))
            rel_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 3, rel_item)

            sup_item = QTableWidgetItem(f"v{item.get('superseded_by', '')} (Newer)")
            sup_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            sup_item.setForeground(QColor("#15803D"))
            self.table.setItem(r, 4, sup_item)

            c_item = QTableWidgetItem(f"{item.get('total_clauses', 0):,}")
            c_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            c_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(r, 5, c_item)

        self.table.itemChanged.connect(self._update_selected_count)

    def _update_selected_count(self):
        selected = 0
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item and item.checkState() == Qt.Checked:
                selected += 1
        self.lbl_selected.setText(f"Selected: {selected} version(s)")
        self.btn_prune.setText(f"🧹 Prune Selected ({selected})")
        self.btn_prune.setEnabled(selected > 0)

    def _select_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(Qt.Checked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _deselect_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _on_prune_confirmed(self):
        target_ids = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item and item.checkState() == Qt.Checked:
                target_ids.append(self.superseded_versions[r]["id"])

        if not target_ids:
            return

        confirm = QMessageBox.question(
            self,
            "Confirm Pruning",
            f"Are you sure you want to permanently delete {len(target_ids)} superseded version(s) from the search index?\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Yes:
            self.pruned_count = self.db.prune_superseded_versions(target_ids)
            self.accept()


# Backward compatibility alias for extracted diff dialog
from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog
