"""
Specification and Release Selection Dialog.
Provides universal master-detail browsing, dynamic series & WG filtering,
spec title search, explicit row checkboxes, and reliable batch selection.
"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_threads import find_cached_spec_file
from modules.spec_search.core.spec_clause_diff import build_llm_clause_prompt, generate_unified_diff
from modules.spec_search.core.spec_search_db import SpecSearchDatabase, parse_version_tuple
from modules.specifications.core.database import SpecsDatabase


def extract_major_release(ver_str: str) -> str:
    """
    Extracts the major 3GPP release identifier from a version string.
    Handles dot notation ('18.4.0' -> 'Rel-18') and lettered codes ('i40' -> 'Rel-18', 'g30' -> 'Rel-16').
    """
    clean = str(ver_str).lstrip("vV").strip()
    if not clean:
        return "Unknown"

    # Standard dot notation (e.g., '18.4.0', '15.1.0')
    if "." in clean:
        first_part = clean.split(".")[0]
        if first_part.isdigit():
            return f"Rel-{first_part}"
        return first_part

    # 3GPP lettered 3-digit versions (e.g., 'g40' where 'g' is the 7th letter after 9 -> Rel-16)
    if len(clean) == 3:
        c0 = clean[0].lower()
        if c0.isdigit():
            return f"Rel-{c0}"
        if "a" <= c0 <= "z":
            rel_num = ord(c0) - ord("a") + 10
            return f"Rel-{rel_num}"

    return clean


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
            query += " AND (sp.number LIKE ? OR sp.title LIKE ? OR (sp.type || ' ' || sp.number) LIKE ?)"
            params.extend([search_pat, search_pat, search_pat])

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

        self.selected_files_info = selected_tasks
        self.accept()

from PyQt5.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QGroupBox,
    QRadioButton,
    QTextEdit,
)

class SpecClauseDiffDialog(QDialog):
    """
    Interactive Dialog to configure clause comparison parameters, select context depth tiers,
    preview the structured Markdown diff, and export or copy prompts for LLM analysis.
    """

    def __init__(
        self,
        db: SpecSearchDatabase,
        spec_number: str,
        current_version: str,
        clause_number: str,
        clause_title: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.spec_number = spec_number
        self.current_version = current_version
        self.clause_number = clause_number
        self.clause_title = clause_title
        self._generated_prompt = ""

        self.setWindowTitle(f"🤖 Compare Clause {self.clause_number} for LLM Analysis (TS {self.spec_number})")
        self.resize(960, 680)
        self._setup_ui()
        self._load_available_versions()
        self._generate_diff_preview()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Top Controls: Version Selectors & Analysis Focus
        # ---------------------------------------------------------------------
        config_group = QGroupBox("Comparison Configuration")
        config_layout = QVBoxLayout(config_group)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("<b>Base Release (Old):</b>"))
        self.base_ver_combo = QComboBox()
        self.base_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.base_ver_combo, stretch=1)

        row1.addWidget(QLabel("<b>Target Release (New):</b>"))
        self.target_ver_combo = QComboBox()
        self.target_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.target_ver_combo, stretch=1)

        row1.addWidget(QLabel("<b>Analysis Focus:</b>"))
        self.focus_combo = QComboBox()
        self.focus_combo.addItem("📋 General Standards & Functional Impact", "standards")
        self.focus_combo.addItem("⚖️ Patent & Prior Art Evaluation", "patent")
        self.focus_combo.addItem("📡 Signalling & Protocol Encoding", "signalling")
        self.focus_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.focus_combo, stretch=2)
        config_layout.addLayout(row1)

        # Context Depth Tier Radios
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("<b>Context Scope:</b>"))
        self.tier_group = QButtonGroup(self)

        self.radio_tier1 = QRadioButton("Tier 1: Exact Clause Only")
        self.radio_tier1.setToolTip("Includes only the delta of the target clause.")
        self.tier_group.addButton(self.radio_tier1, 1)
        row2.addWidget(self.radio_tier1)

        self.radio_tier2 = QRadioButton("Tier 2: + Parent Procedure Scope (Recommended)")
        self.radio_tier2.setToolTip("Includes hierarchical breadcrumbs, parent intro/preconditions, and the clause delta.")
        self.radio_tier2.setChecked(True)
        self.tier_group.addButton(self.radio_tier2, 2)
        row2.addWidget(self.radio_tier2)

        self.radio_tier3 = QRadioButton("Tier 3: + Full Procedure Branch")
        self.radio_tier3.setToolTip("Includes sibling subclauses in the same procedure branch for deep architectural context.")
        self.tier_group.addButton(self.radio_tier3, 3)
        row2.addWidget(self.radio_tier3)

        row2.addStretch()
        self.tier_group.buttonClicked.connect(self._generate_diff_preview)
        config_layout.addLayout(row2)

        layout.addWidget(config_group)

        # ---------------------------------------------------------------------
        # Markdown Preview Area
        # ---------------------------------------------------------------------
        preview_header = QHBoxLayout()
        preview_header.addWidget(QLabel("<b>Generated Prompt Preview for LLM:</b>"))
        preview_header.addStretch()
        layout.addLayout(preview_header)

        self.preview_browser = QTextEdit()
        self.preview_browser.setReadOnly(True)
        self.preview_browser.setStyleSheet("""
            QTextEdit {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.4;
                color: #1E293B;
                padding: 6px;
            }
        """)
        layout.addWidget(self.preview_browser)

        # ---------------------------------------------------------------------
        # Action Buttons
        # ---------------------------------------------------------------------
        btn_bar = QHBoxLayout()

        self.btn_copy = QPushButton("📋 Copy Prompt for LLM")
        self.btn_copy.setStyleSheet("font-weight: bold; background-color: #0284C7; color: white; padding: 6px 16px; border-radius: 4px;")
        self.btn_copy.clicked.connect(self._copy_prompt)
        btn_bar.addWidget(self.btn_copy)

        self.btn_save = QPushButton("💾 Save Markdown (.md)")
        self.btn_save.setStyleSheet("font-weight: bold; background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; padding: 6px 14px; border-radius: 4px;")
        self.btn_save.clicked.connect(self._save_markdown)
        btn_bar.addWidget(self.btn_save)

        btn_bar.addStretch()
        self.btn_close = QPushButton("Close")
        self.btn_close.clicked.connect(self.accept)
        btn_bar.addWidget(self.btn_close)

        layout.addLayout(btn_bar)

    def _load_available_versions(self):
        """Populates version dropdowns for the specification."""
        vers = self.db.get_versions_for_spec(self.spec_number)
        if not vers:
            return

        self.base_ver_combo.blockSignals(True)
        self.target_ver_combo.blockSignals(True)

        self.base_ver_combo.clear()
        self.target_ver_combo.clear()

        for v in vers:
            ver_str = v["version"]
            date_str = f" ({v['release_date']})" if v.get("release_date") else ""
            label = f"v{ver_str}{date_str}"
            self.base_ver_combo.addItem(label, v)
            self.target_ver_combo.addItem(label, v)

        # Default Target Version to current_version
        target_idx = 0
        for i in range(self.target_ver_combo.count()):
            v_data = self.target_ver_combo.itemData(i)
            if v_data and v_data.get("version") == self.current_version:
                target_idx = i
                break
        self.target_ver_combo.setCurrentIndex(target_idx)

        # Default Base Version to the preceding version
        base_idx = min(target_idx + 1, self.base_ver_combo.count() - 1)
        self.base_ver_combo.setCurrentIndex(base_idx)

        self.base_ver_combo.blockSignals(False)
        self.target_ver_combo.blockSignals(False)

    def _generate_diff_preview(self):
        """Generates the unified diff and updates the markdown preview."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()

        if not base_data or not target_data:
            self.preview_browser.setText("Select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        base_date = base_data.get("release_date", "")
        target_ver = target_data.get("version", "")
        target_date = target_data.get("release_date", "")

        tier = self.tier_group.checkedId()
        focus_mode = self.focus_combo.currentData() or "standards"

        # Fetch Clause Text
        base_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, base_ver, self.clause_number)
        target_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, target_ver, self.clause_number)

        base_text = base_clause.get("content", "") if base_clause else ""
        target_text = target_clause.get("content", "") if target_clause else ""

        base_lbl = f"TS {self.spec_number} v{base_ver} (Clause {self.clause_number})"
        target_lbl = f"TS {self.spec_number} v{target_ver} (Clause {self.clause_number})"
        diff_text = generate_unified_diff(base_text, target_text, base_lbl, target_lbl)

        # Fetch Hierarchy Context if requested
        hierarchy = None
        if tier >= 2:
            hierarchy = self.db.get_clause_hierarchy(self.spec_number, target_ver, self.clause_number)

        # Fetch Branch Context if requested
        branch_clauses = None
        if tier >= 3:
            branch_clauses = self.db.get_branch_clauses(self.spec_number, target_ver, self.clause_number)

        # Build Markdown Prompt
        self._generated_prompt = build_llm_clause_prompt(
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
            base_version=base_ver,
            base_date=base_date,
            target_version=target_ver,
            target_date=target_date,
            diff_text=diff_text,
            tier=tier,
            hierarchy=hierarchy,
            branch_clauses=branch_clauses,
            focus_mode=focus_mode,
        )

        self.preview_browser.setPlainText(self._generated_prompt)

    def _copy_prompt(self):
        if self._generated_prompt:
            QApplication.clipboard().setText(self._generated_prompt)
            self.btn_copy.setText("✅ Copied to Clipboard!")
            QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 Copy Prompt for LLM"))

    def _save_markdown(self):
        if not self._generated_prompt:
            return
        default_name = f"TS_{self.spec_number}_Clause_{self.clause_number}_Diff.md".replace(".", "_")
        path, _ = QFileDialog.getSaveFileName(self, "Save Clause Diff Markdown", default_name, "Markdown Files (*.md)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self._generated_prompt)
                QMessageBox.information(self, "Saved", f"Prompt saved successfully to:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to write file:\n{e}")