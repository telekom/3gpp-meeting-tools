# --- File: modules/specs_db/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QGroupBox)
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
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Series", "Spec #", "Title", "File", "Version", "Action"])

        # Make the table look nice and stretch the Title column
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Title takes up remaining space
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # Read-only
        self.table.setAlternatingRowColors(True)
        main_layout.addWidget(self.table)

        self.setLayout(main_layout)

    def _on_update_clicked(self):
        force_meta = self.force_meta_checkbox.isChecked()
        self.update_db_requested.emit(force_meta)

    def refresh_table(self):
        """Queries the database and populates the table."""
        spec_query = self.spec_search_input.text().strip()
        version_query = self.version_search_input.text().strip()

        # Fetch results from database
        results = self.db.search_files(
            spec_number=spec_query if spec_query else None,
            release_version=version_query if version_query else None
        )

        self.table.setRowCount(0)  # Clear existing rows

        for row_idx, row_data in enumerate(results):
            # row_data format: (Series, Spec Number, Title, Filename, Version, URL)
            self.table.insertRow(row_idx)

            # Populate columns 0 to 4 with text
            for col_idx in range(5):
                item = QTableWidgetItem(str(row_data[col_idx] if row_data[col_idx] else ""))
                self.table.setItem(row_idx, col_idx, item)

            # Column 5: Action Button (Open Link)
            link_url = row_data[5]
            if link_url:
                open_btn = QPushButton("⬇️ Download")
                open_btn.setCursor(Qt.PointingHandCursor)
                # Capture the URL in the lambda to prevent late-binding issues
                open_btn.clicked.connect(lambda checked, url=link_url: webbrowser.open(url))
                self.table.setCellWidget(row_idx, 5, open_btn)