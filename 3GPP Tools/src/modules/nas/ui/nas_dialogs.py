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
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import find_cached_spec_file
from modules.specifications.core.database import SpecsDatabase


def get_spec_title(spec_number: str) -> str:
    """Returns a descriptive title for 3GPP NAS and ASN.1 specifications."""
    spec_titles = {
        "38.331": "TS 38.331 (NR RRC)",
        "36.331": "TS 36.331 (LTE RRC)",
        "38.413": "TS 38.413 (NGAP)",
        "24.501": "TS 24.501 (5GS NAS)",
        "24.301": "TS 24.301 (EPS NAS)",
        "24.008": "TS 24.008 (Core Network)",
        "23.501": "TS 23.501 (5GS Architecture)",
        "23.502": "TS 23.502 (5GS Procedures)",
    }
    return spec_titles.get(spec_number, f"TS {spec_number}")


class NASVersionSelectDialog(QDialog):
    """Dialog allowing selection of NAS (24.501/24.301) and RRC/NGAP (38.331/36.331/38.413) versions."""

    def __init__(
            self,
            specs_db: SpecsDatabase,
            nas_db: NASDatabase,
            cache_dir: Path,
            parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.nas_db = nas_db
        self.cache_dir = cache_dir
        self.selected_files_info: List[Dict[str, Any]] = []

        self.setWindowTitle("Select 3GPP Specification Version(s) to Ingest")
        self.resize(740, 480)
        self._setup_ui()
        self._load_available_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("Specification:"))

        self.spec_combo = QComboBox()
        self.spec_combo.addItem("TS 38.331 (NR RRC)", "38.331")
        self.spec_combo.addItem("TS 36.331 (LTE RRC)", "36.331")
        self.spec_combo.addItem("TS 38.413 (NGAP)", "38.413")
        self.spec_combo.addItem("TS 24.501 (5GS NAS)", "24.501")
        self.spec_combo.addItem("TS 24.301 (EPS NAS)", "24.301")
        self.spec_combo.currentIndexChanged.connect(self._load_available_versions)
        top_bar.addWidget(self.spec_combo)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        info_lbl = QLabel(
            "Select one or more versions from the specification archive (use Ctrl+Click or Shift+Click).\n"
            "Files not currently cached locally will be downloaded automatically from the 3GPP FTP server."
        )
        info_lbl.setStyleSheet("color: #555; padding-bottom: 4px;")
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Version", "Filename", "Local Cache", "DB Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("📥 Fetch && Ingest Selected")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _load_available_versions(self):
        spec_num = self.spec_combo.currentData()
        spec_files = self.specs_db.search_files(spec_number=spec_num)

        imported_entries = {
            (v["spec_number"], v["version"]) for v in self.nas_db.get_imported_versions()
        }

        spec_files = sorted(
            spec_files,
            key=lambda row: parse_version_tuple(row[5]),
            reverse=True,
        )

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(spec_files):
            _, s_num, _, _, filename, version, url, specification_upload_date = row_data
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
            db_item = QTableWidgetItem("✅ Ingested" if in_db else "⚪ Ready")
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
            QMessageBox.warning(
                self, "Selection Required", "Please select at least one specification version to import."
            )
            return

        self.selected_files_info = [
            self.table.item(row, 0).data(Qt.UserRole)
            for row in selected_rows
        ]
        self.accept()