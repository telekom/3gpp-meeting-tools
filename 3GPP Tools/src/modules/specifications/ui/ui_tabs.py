# --- File: modules/specs_db/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QGroupBox, QComboBox)
from PyQt5.QtCore import pyqtSignal, Qt

from modules.specifications.core.database import SpecsDatabase


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)

    def __init__(self, db_path: Path):
        super().__init__()
        # Initialize a read-only connection for the UI to display data
        self.db = SpecsDatabase(db_path)
        self._setup_ui()
        self.refresh_table()  # Load initial data

    def _setup_ui(self):
        main_layout = QVBoxLayout()

        # --- TOP PANEL: Controls ---
        control_group = QGroupBox("Database Synchronization")
        control_layout = QHBoxLayout()

        self.force_meta_checkbox = QCheckBox("Force Scrape DynaReports (Metadata)")
        self.force_meta_checkbox.setToolTip("If unchecked, only fetches metadata for new specifications.")
        control_layout.addWidget(self.force_meta_checkbox)

        self.update_btn = QPushButton("🔄 Start Synchronization")
        self.update_btn.clicked.connect(self._on_update_clicked)
        control_layout.addWidget(self.update_btn)

        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        # --- MIDDLE PANEL: Search & Filter ---
        search_layout = QHBoxLayout()

        search_layout.addWidget(QLabel("🔍 Spec Number:"))
        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("e.g. 23.501")
        self.spec_search_input.returnPressed.connect(self.refresh_table)
        search_layout.addWidget(self.spec_search_input)

        search_layout.addWidget(QLabel("Release/Version:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15. or 16.2")
        self.version_search_input.returnPressed.connect(self.refresh_table)
        search_layout.addWidget(self.version_search_input)

        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self.refresh_table)
        search_layout.addWidget(self.search_btn)

        main_layout.addLayout(search_layout)

        # --- BOTTOM PANEL: Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Spec #", "Type", "Title", "Version / Download"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Title takes space

        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

    def _on_update_clicked(self):
        force_meta = self.force_meta_checkbox.isChecked()
        self.update_db_requested.emit(force_meta)

    def refresh_table(self):
        specs = self.db.get_all_specifications()  # Gets (number, title, type)
        self.table.setRowCount(0)

        for row_idx, (number, title, spec_type) in enumerate(specs):
            self.table.insertRow(row_idx)

            # 1. Spec Number (Concatenated with type)
            display_num = f"{spec_type} {number}" if spec_type else number
            self.table.setItem(row_idx, 0, QTableWidgetItem(display_num))

            # 2. Type and Title
            self.table.setItem(row_idx, 1, QTableWidgetItem(spec_type))
            self.table.setItem(row_idx, 2, QTableWidgetItem(title))

            # 3. Version Dropdown
            version_combo = QComboBox()
            versions = self.db.get_versions_for_spec(number)

            for ver, url, filename in versions:
                version_combo.addItem(f"v{ver}", userData=url)

            # Button to trigger download of selected version
            download_btn = QPushButton("⬇️")
            download_btn.clicked.connect(lambda _, c=version_combo: webbrowser.open(c.currentData()))

            # Layout the combo and button together
            cell_widget = QWidget()
            layout = QHBoxLayout(cell_widget)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.addWidget(version_combo)
            layout.addWidget(download_btn)

            self.table.setCellWidget(row_idx, 3, cell_widget)