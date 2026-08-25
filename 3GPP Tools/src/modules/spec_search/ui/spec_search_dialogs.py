"""
Specification and Release Selection Dialog.
Allows selecting any TS/TR specification from 3gpp_data.db to index.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.core.nas_threads import find_cached_spec_file
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.specifications.core.database import SpecsDatabase


class SpecSearchVersionSelectDialog(QDialog):
    """Dialog to pick versions across any 3GPP specification."""

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
        self.selected_files_info: List[Dict[str, Any]] = []

        self.setWindowTitle("Select 3GPP Specification Versions to Index")
        self.resize(760, 500)
        self._setup_ui()
        self._load_available_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("Specification:"))

        self.spec_combo = QComboBox()
        self.spec_combo.setEditable(True)
        # Pre-populate prominent specifications
        for spec in ["24.501", "24.301", "38.331", "36.331", "38.413", "23.501", "23.502", "33.501", "29.500"]:
            self.spec_combo.addItem(f"TS {spec}", spec)

        self.spec_combo.currentIndexChanged.connect(self._load_available_versions)
        self.spec_combo.lineEdit().returnPressed.connect(self._on_spec_custom_typed)
        top_bar.addWidget(self.spec_combo)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        info_lbl = QLabel("Select versions to index into the Full-Text search engine (Ctrl+Click / Shift+Click):")
        info_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Version", "Filename", "Local Cache", "Search DB Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("📥 Fetch && Index Selected")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _on_spec_custom_typed(self):
        typed = self.spec_combo.currentText().replace("TS", "").strip()
        self.spec_combo.setCurrentText(typed)
        self._load_available_versions()

    def _load_available_versions(self):
        spec_num = self.spec_combo.currentData() or self.spec_combo.currentText().replace("TS", "").strip()
        spec_files = self.specs_db.search_files(spec_number=spec_num)

        imported_entries = {
            (v["spec_number"], v["version"]) for v in self.search_db.get_imported_versions()
        }

        spec_files = sorted(
            spec_files,
            key=lambda row: parse_version_tuple(row[5]),
            reverse=True,
        )

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(spec_files):
            _, s_num, _, _, filename, version, url = row_data
            self.table.insertRow(row_idx)

            v_item = QTableWidgetItem(f"v{version}")
            v_item.setData(
                Qt.UserRole,
                {
                    "spec_number": s_num,
                    "version": version,
                    "filename": filename,
                    "file_url": url,
                },
            )
            self.table.setItem(row_idx, 0, v_item)
            self.table.setItem(row_idx, 1, QTableWidgetItem(filename))

            cached_file = find_cached_spec_file(filename, s_num)
            if cached_file:
                cache_text = f"🟢 Cached ({cached_file.suffix[1:].upper()})"
                cache_item = QTableWidgetItem(cache_text)
                cache_item.setForeground(Qt.darkGreen)
            else:
                cache_item = QTableWidgetItem("🌐 Remote (FTP)")
                cache_item.setForeground(Qt.darkGray)
            self.table.setItem(row_idx, 2, cache_item)

            in_db = (s_num, version) in imported_entries
            db_item = QTableWidgetItem("✅ Indexed" if in_db else "⚪ Ready")
            if in_db:
                db_item.setForeground(Qt.blue)
            self.table.setItem(row_idx, 3, db_item)

    def _on_item_double_clicked(self, item: QTableWidgetItem):
        row = item.row()
        self.selected_files_info = [self.table.item(row, 0).data(Qt.UserRole)]
        self.accept()

    def _on_accept(self):
        selected_rows = sorted(list(set(item.row() for item in self.table.selectedItems())))
        if not selected_rows:
            QMessageBox.warning(self, "Selection Required", "Please select at least one version to index.")
            return

        self.selected_files_info = [
            self.table.item(row, 0).data(Qt.UserRole)
            for row in selected_rows
        ]
        self.accept()