"""
Specification and Release Selection Dialog.
Provides universal master-detail browsing, dynamic series & WG filtering,
spec title search, explicit row checkboxes, and reliable batch selection.
"""

import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QFont
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




class SpecSmartImportDialog(QDialog):
    """
    Targeted Intelligent Import Dialog for a single 3GPP Specification.
    Groups all historical versions by 3GPP Release (Rel-15, Rel-16, Rel-17, Rel-18, etc.),
    identifies the newest version per release, computes delta status against the local Spec Search DB
    (Up to Date, Update Available, New Release), and allows 1-click batch ingestion.
    """

    def __init__(
        self,
        specs_db: SpecsDatabase,
        search_db: SpecSearchDatabase,
        cache_dir: Path,
        initial_spec: Optional[str] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.search_db = search_db
        self.cache_dir = cache_dir
        self.initial_spec = (initial_spec or "").strip()
        self.logger = logging.getLogger(__name__)

        self.current_spec = ""
        self.current_spec_title = ""
        self.current_spec_wg = ""
        self.releases_data: List[Dict[str, Any]] = []
        self.selected_tasks: List[Dict[str, Any]] = []
        self.auto_replace_older: bool = True

        self.setWindowTitle("⭐ Smart Import Specification (Latest per Release)")
        self.resize(880, 560)
        self._setup_ui()

        if self.initial_spec:
            self.spec_input.setText(self.initial_spec)
            self._on_find_releases()
        else:
            self.spec_input.setFocus()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Header Frame: Search input & spec info
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
            }
        """)
        h_layout = QVBoxLayout(header_frame)
        h_layout.setContentsMargins(10, 10, 10, 10)
        h_layout.setSpacing(8)

        # Search Bar Row
        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        lbl_prompt = QLabel("🔍 Specification:")
        lbl_prompt.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B;")
        search_row.addWidget(lbl_prompt)

        self.spec_input = QLineEdit()
        self.spec_input.setPlaceholderText("Enter spec number (e.g. 23.501, 38.300, 24.501, 38.211)...")
        self.spec_input.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 5px 8px;
                font-size: 12px;
                font-weight: 500;
                color: #0F172A;
            }
            QLineEdit:focus {
                border-color: #0284C7;
            }
        """)
        self.spec_input.setClearButtonEnabled(True)
        self.spec_input.returnPressed.connect(self._on_find_releases)
        search_row.addWidget(self.spec_input, stretch=1)

        self.btn_find = QPushButton("🔍 Find Releases")
        self.btn_find.setCursor(Qt.PointingHandCursor)
        self.btn_find.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: #FFFFFF;
                border: 1px solid #0284C7;
                border-radius: 4px;
                padding: 6px 14px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #0369A1;
            }
        """)
        self.btn_find.clicked.connect(self._on_find_releases)
        search_row.addWidget(self.btn_find)
        h_layout.addLayout(search_row)

        # Spec Metadata Strip
        self.info_strip = QWidget()
        info_layout = QHBoxLayout(self.info_strip)
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(8)

        self.lbl_spec_title = QLabel("Enter a 3GPP specification number above to load available releases.")
        self.lbl_spec_title.setStyleSheet("color: #475569; font-size: 11px; font-weight: 500;")
        self.lbl_spec_title.setWordWrap(True)
        info_layout.addWidget(self.lbl_spec_title, stretch=1)

        self.lbl_status_badge = QLabel("")
        self.lbl_status_badge.setStyleSheet("""
            background-color: #F1F5F9;
            color: #334155;
            border: 1px solid #CBD5E1;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 10px;
            font-weight: bold;
        """)
        self.lbl_status_badge.setVisible(False)
        info_layout.addWidget(self.lbl_status_badge)

        h_layout.addWidget(self.info_strip)
        layout.addWidget(header_frame)

        # Quick Filter Chips Bar
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(6)

        lbl_select = QLabel("Select:")
        lbl_select.setStyleSheet("font-size: 11px; font-weight: bold; color: #475569;")
        chips_layout.addWidget(lbl_select)

        self.btn_chip_updates = QPushButton("⚡ Updates && New Only")
        self.btn_chip_updates.setToolTip("Select only releases that have updates or are not yet indexed")
        self.btn_chip_updates.setCursor(Qt.PointingHandCursor)
        self.btn_chip_updates.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 600;
                color: #C2410C;
                background-color: #FFF7ED;
                border: 1px solid #FDBA74;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #FFEDD5;
                border-color: #FB923C;
            }
        """)
        self.btn_chip_updates.clicked.connect(self._select_updates_and_new)
        chips_layout.addWidget(self.btn_chip_updates)

        self.btn_chip_all = QPushButton("⭐ All Releases")
        self.btn_chip_all.setToolTip("Select the latest version of all releases")
        self.btn_chip_all.setCursor(Qt.PointingHandCursor)
        self.btn_chip_all.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 600;
                color: #0369A1;
                background-color: #F0F9FF;
                border: 1px solid #BAE6FD;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #E0F2FE;
            }
        """)
        self.btn_chip_all.clicked.connect(self._select_all)
        chips_layout.addWidget(self.btn_chip_all)

        self.btn_chip_r18 = QPushButton("Rel-18+")
        self.btn_chip_r18.setToolTip("Select Release 18 and newer (5G-Advanced & 6G)")
        self.btn_chip_r18.setCursor(Qt.PointingHandCursor)
        self.btn_chip_r18.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 600;
                color: #4338CA;
                background-color: #EEF2FF;
                border: 1px solid #C7D2FE;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #E0E7FF;
            }
        """)
        self.btn_chip_r18.clicked.connect(lambda: self._select_min_release(18))
        chips_layout.addWidget(self.btn_chip_r18)

        self.btn_chip_r17 = QPushButton("Rel-17+")
        self.btn_chip_r17.setToolTip("Select Release 17 and newer")
        self.btn_chip_r17.setCursor(Qt.PointingHandCursor)
        self.btn_chip_r17.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 600;
                color: #475569;
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
            }
        """)
        self.btn_chip_r17.clicked.connect(lambda: self._select_min_release(17))
        chips_layout.addWidget(self.btn_chip_r17)

        self.btn_chip_clear = QPushButton("◻️ Clear")
        self.btn_chip_clear.setToolTip("Deselect all releases")
        self.btn_chip_clear.setCursor(Qt.PointingHandCursor)
        self.btn_chip_clear.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 500;
                color: #64748B;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #F8FAFC;
                color: #0F172A;
            }
        """)
        self.btn_chip_clear.clicked.connect(self._clear_selection)
        chips_layout.addWidget(self.btn_chip_clear)

        chips_layout.addStretch()

        self.lbl_selection_summary = QLabel("Selected: 0 release(s)")
        self.lbl_selection_summary.setStyleSheet("font-weight: bold; font-size: 11px; color: #0284C7;")
        chips_layout.addWidget(self.lbl_selection_summary)

        layout.addLayout(chips_layout)

        # Table of Available Releases
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Select", "Release", "Latest Version", "Status", "Release Date", "Filename"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                gridline-color: #F1F5F9;
                font-size: 11px;
            }
            QTableWidget::item:selected {
                background-color: #EBF3FC;
                color: #0F172A;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                border: none;
                border-bottom: 1px solid #CBD5E1;
                padding: 4px 6px;
                font-weight: bold;
                font-size: 11px;
                color: #334155;
            }
        """)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        layout.addWidget(self.table)

        # Options & Ingestion Bar
        opt_layout = QHBoxLayout()
        self.chk_auto_supersede = QCheckBox("Auto-supersede older minor versions within the same release (recommended)")
        self.chk_auto_supersede.setChecked(True)
        self.chk_auto_supersede.setToolTip(
            "When checked, ingesting newer versions (e.g. v20.3.0) will automatically remove older\n"
            "superseded versions (e.g. v20.2.0) of that release from the search index."
        )
        self.chk_auto_supersede.setStyleSheet("font-size: 11px; color: #334155; font-weight: 500;")
        opt_layout.addWidget(self.chk_auto_supersede)
        opt_layout.addStretch()
        layout.addLayout(opt_layout)

        # Action Buttons
        btn_box = QHBoxLayout()
        btn_box.setSpacing(8)
        btn_box.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 6px 14px;
                font-size: 11px;
                font-weight: 500;
                color: #334155;
            }
            QPushButton:hover {
                background-color: #F8FAFC;
            }
        """)
        self.btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cancel)

        self.btn_ingest = QPushButton("📥 Ingest Selected Releases (0)")
        self.btn_ingest.setEnabled(False)
        self.btn_ingest.setCursor(Qt.PointingHandCursor)
        self.btn_ingest.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: #FFFFFF;
                border: 1px solid #0284C7;
                border-radius: 4px;
                padding: 6px 16px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #0369A1;
            }
            QPushButton:disabled {
                background-color: #CBD5E1;
                color: #94A3B8;
                border-color: #CBD5E1;
            }
        """)
        self.btn_ingest.clicked.connect(self._on_ingest_confirmed)
        btn_box.addWidget(self.btn_ingest)

        layout.addLayout(btn_box)

    def _on_find_releases(self):
        raw = self.spec_input.text().strip()
        if not raw:
            return

        clean_spec = re.sub(r"^(?:TS|TR)\s*", "", raw, flags=re.IGNORECASE).strip()
        if re.match(r"^\d{5}$", clean_spec):
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:]}"
        elif re.match(r"^\d{6}$", clean_spec):
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:5]}-{clean_spec[5:]}"

        self.current_spec = clean_spec

        # Query metadata
        spec_details = self.specs_db.get_spec_details(clean_spec) if hasattr(self.specs_db, "get_spec_details") else {}
        self.current_spec_title = spec_details.get("title", "")
        self.current_spec_wg = spec_details.get("primary_group", "")

        files = self.specs_db.get_spec_files(clean_spec)
        if not files:
            self.table.setRowCount(0)
            self.lbl_spec_title.setText(f"❌ Specification <b>TS {clean_spec}</b> was not found in 3GPP specifications database.")
            self.lbl_spec_title.setStyleSheet("color: #DC2626; font-size: 11px;")
            self.lbl_status_badge.setVisible(False)
            self._update_selection_count()
            return

        # Display header
        title_str = f"<b>TS {clean_spec}</b>"
        if self.current_spec_title:
            title_str += f": {self.current_spec_title}"
        if self.current_spec_wg:
            title_str += f" — <i>{self.current_spec_wg}</i>"
        self.lbl_spec_title.setText(title_str)
        self.lbl_spec_title.setStyleSheet("color: #1E293B; font-size: 11px;")

        # Group files by release
        rel_files: Dict[str, Dict[str, Any]] = {}
        for f in files:
            v_str = str(f.get("version", "")).lstrip("vV").strip()
            if not v_str:
                continue
            rel = extract_major_release(v_str)
            if rel not in rel_files:
                rel_files[rel] = f

        # Query currently indexed versions in search_db
        indexed_rows = self.search_db.get_versions_for_spec(clean_spec)
        indexed_by_rel: Dict[str, List[str]] = {}
        for row in indexed_rows:
            v_str = str(row.get("version", "")).lstrip("vV").strip()
            rel = extract_major_release(v_str)
            indexed_by_rel.setdefault(rel, []).append(v_str)

        def _rel_sort_key(rel_str: str) -> int:
            num = rel_str.replace("Rel-", "").strip()
            return int(num) if num.isdigit() else -1

        sorted_releases = sorted(rel_files.keys(), key=_rel_sort_key, reverse=True)

        self.releases_data = []
        updates_count = 0
        new_count = 0
        current_count = 0

        for rel in sorted_releases:
            f = rel_files[rel]
            avail_ver = str(f.get("version", "")).lstrip("vV").strip()
            curr_indexed = indexed_by_rel.get(rel, [])

            if not curr_indexed:
                status = "new"
                status_text = "🔵 New Release"
                status_desc = "Not indexed in Spec Search DB yet."
                new_count += 1
            elif parse_version_tuple(curr_indexed[0]) < parse_version_tuple(avail_ver):
                status = "update"
                status_text = f"🟠 Update (v{curr_indexed[0]} → v{avail_ver})"
                status_desc = f"Currently indexed v{curr_indexed[0]} is older than available v{avail_ver}."
                updates_count += 1
            else:
                status = "current"
                status_text = f"🟢 Up to date (v{avail_ver})"
                status_desc = "Indexed version matches the latest available release."
                current_count += 1

            self.releases_data.append({
                "release": rel,
                "release_num": _rel_sort_key(rel),
                "version": avail_ver,
                "status": status,
                "status_text": status_text,
                "status_desc": status_desc,
                "upload_date": f.get("upload_date", ""),
                "filename": f.get("filename", ""),
                "url": f.get("url", ""),
            })

        self.lbl_status_badge.setText(f"{len(self.releases_data)} releases: {updates_count} update(s), {new_count} new, {current_count} current")
        self.lbl_status_badge.setStyleSheet("""
            background-color: #F1F5F9;
            color: #334155;
            border: 1px solid #CBD5E1;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 10px;
            font-weight: bold;
        """)
        self.lbl_status_badge.setVisible(True)

        self._populate_table()
        # Default: Select Updates & New
        self._select_updates_and_new()

    def _populate_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.releases_data))

        for r, item in enumerate(self.releases_data):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(r, 0, chk)

            rel_item = QTableWidgetItem(item["release"])
            rel_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            rel_item.setTextAlignment(Qt.AlignCenter)
            font_b = QFont()
            font_b.setBold(True)
            rel_item.setFont(font_b)
            self.table.setItem(r, 1, rel_item)

            ver_item = QTableWidgetItem(f"v{item['version']}")
            ver_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 2, ver_item)

            status_item = QTableWidgetItem(item["status_text"])
            status_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            status_item.setToolTip(item["status_desc"])
            if item["status"] == "update":
                status_item.setForeground(QColor("#C2410C"))
            elif item["status"] == "new":
                status_item.setForeground(QColor("#0369A1"))
            else:
                status_item.setForeground(QColor("#15803D"))
            self.table.setItem(r, 3, status_item)

            d_item = QTableWidgetItem(item.get("upload_date", ""))
            d_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 4, d_item)

            fn_item = QTableWidgetItem(item.get("filename", ""))
            fn_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 5, fn_item)

        self.table.blockSignals(False)

    def _select_updates_and_new(self):
        self.table.blockSignals(True)
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked if item["status"] in ("update", "new") else Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _select_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _select_min_release(self, min_rel: int):
        self.table.blockSignals(True)
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked if item["release_num"] >= min_rel else Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _clear_selection(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _on_table_item_changed(self, item: QTableWidgetItem):
        if item.column() == 0:
            self._update_selection_count()

    def _on_cell_double_clicked(self, row: int, col: int):
        chk = self.table.item(row, 0)
        if chk:
            new_st = Qt.Unchecked if chk.checkState() == Qt.Checked else Qt.Checked
            chk.setCheckState(new_st)
            self._update_selection_count()

    def _update_selection_count(self):
        selected = 0
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk and chk.checkState() == Qt.Checked:
                selected += 1

        self.lbl_selection_summary.setText(f"Selected: {selected} release(s)")
        self.btn_ingest.setText(f"📥 Ingest Selected Releases ({selected})")
        self.btn_ingest.setEnabled(selected > 0)

    def _on_ingest_confirmed(self):
        self.selected_tasks = []
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk and chk.checkState() == Qt.Checked:
                self.selected_tasks.append({
                    "spec_number": self.current_spec,
                    "version": item["version"],
                    "filename": item["filename"],
                    "file_url": item["url"],
                    "release_date": item["upload_date"],
                    "spec_title": self.current_spec_title or "",
                })

        if not self.selected_tasks:
            return

        self.auto_replace_older = self.chk_auto_supersede.isChecked()
        self.accept()


# Backward compatibility alias for extracted diff dialog
from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog
