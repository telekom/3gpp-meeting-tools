# --- File: modules/specs_db/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QGroupBox, QComboBox, )
from PyQt5.QtCore import pyqtSignal, QTimer, Qt

from modules.specifications.core.database import SpecsDatabase


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)

        # --- NEW: Setup the debounce timer for live search ---
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)  # Ensures the timer only fires once per typing pause
        self.search_timer.setInterval(400)  # Waits 400ms after the last keystroke to search
        self.search_timer.timeout.connect(self.refresh_table)

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
        # ---> NEW: Restart the timer every time the text changes
        self.spec_search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.spec_search_input)

        search_layout.addWidget(QLabel("Release/Version:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15. or 16.2")
        # ---> NEW: Restart the timer every time the text changes
        self.version_search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.version_search_input)

        # (The Search button has been completely removed)

        main_layout.addLayout(search_layout)

        # --- BOTTOM PANEL: Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Spec #", "Type", "Title", "Version / Download"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Title takes up remaining space

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # Read-only
        self.table.setAlternatingRowColors(True)
        main_layout.addWidget(self.table)

        self.setLayout(main_layout)

    def _on_update_clicked(self):
        force_meta = self.force_meta_checkbox.isChecked()
        self.update_db_requested.emit(force_meta)

    def refresh_table(self):
        spec_query = self.spec_search_input.text().strip()
        version_query = self.version_search_input.text().strip()

        # Fetch results from database
        specs = self.db.search_files(
            spec_number=spec_query if spec_query else None,
            release_version=version_query if version_query else None
        )

        self.table.setRowCount(0)  # Clear existing rows

        # Because we changed the UI to Group by Spec earlier,
        # let's map the raw search_files rows back into unique groupings:
        grouped_specs = {}
        for row in specs:
            series, spec_num, title, filename, version, url = row
            if spec_num not in grouped_specs:
                grouped_specs[spec_num] = {
                    'title': title,
                    'type': "", # We extract Type separately in standard view, but Title is usually fine here
                    'versions': []
                }
            grouped_specs[spec_num]['versions'].append((version, url, filename))

        for row_idx, (spec_num, data) in enumerate(grouped_specs.items()):
            self.table.insertRow(row_idx)

            # 1. Spec Number
            self.table.setItem(row_idx, 0, QTableWidgetItem(spec_num))

            # 2. Type (Left blank or extracted if needed from Title)
            self.table.setItem(row_idx, 1, QTableWidgetItem(data['type']))

            # 3. Title
            self.table.setItem(row_idx, 2, QTableWidgetItem(data['title'] if data['title'] else "Unknown Title"))

            # 4. Version Dropdown & Download Button
            version_combo = QComboBox()
            for ver, url, fname in data['versions']:
                version_combo.addItem(f"v{ver}", userData=url)

            download_btn = QPushButton("⬇️")
            download_btn.setCursor(Qt.PointingHandCursor)
            download_btn.clicked.connect(lambda _, c=version_combo: webbrowser.open(c.currentData()))

            cell_widget = QWidget()
            layout = QHBoxLayout(cell_widget)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.addWidget(version_combo)
            layout.addWidget(download_btn)

            self.table.setCellWidget(row_idx, 3, cell_widget)