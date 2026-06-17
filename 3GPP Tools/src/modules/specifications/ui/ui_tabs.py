# --- File: modules/specs_db/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QGroupBox, QComboBox, QMenu, QAbstractItemView)
from PyQt5.QtCore import pyqtSignal, Qt, QTimer

from modules.specifications.core.database import SpecsDatabase


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
    # ---> NEW: Signal for targeted updates passing a list of specification numbers
    update_specific_requested = pyqtSignal(list, bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(400)
        self.search_timer.timeout.connect(self.refresh_table)

        self._setup_ui()
        self.refresh_table()

    def _setup_ui(self):
        main_layout = QVBoxLayout()

        # --- TOP PANEL: Controls ---
        control_group = QGroupBox("Database Synchronization")
        control_layout = QHBoxLayout()

        self.force_meta_checkbox = QCheckBox("Force Scrape DynaReports (Metadata)")
        self.force_meta_checkbox.setToolTip("If unchecked, only fetches metadata for new specifications.")
        control_layout.addWidget(self.force_meta_checkbox)

        self.update_btn = QPushButton("🔄 Full Synchronization")
        self.update_btn.clicked.connect(lambda: self.update_db_requested.emit(self.force_meta_checkbox.isChecked()))
        control_layout.addWidget(self.update_btn)

        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        # --- MIDDLE PANEL: Search & Filter ---
        search_layout = QHBoxLayout()

        search_layout.addWidget(QLabel("🔍 Spec Number:"))
        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("e.g. 23.501")
        self.spec_search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.spec_search_input)

        search_layout.addWidget(QLabel("Release/Version:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15. or 16.2")
        self.version_search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.version_search_input)

        self.count_label = QLabel("Showing 0 specifications")
        self.count_label.setStyleSheet("font-weight: bold; color: #555555;")

        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Filter Results</b>"))
        header_layout.addStretch()
        header_layout.addWidget(self.count_label)

        main_layout.addLayout(search_layout)
        main_layout.addSpacing(10)
        main_layout.addLayout(header_layout)

        # --- BOTTOM PANEL: Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Specification", "Title", "Version / Download"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)

        # ---> NEW: Enable Multiple Selection and Full Row Selection
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        # ---> NEW: Setup Right-Click Context Menu
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

    def _show_context_menu(self, position):
        """Displays the right-click menu for targeted updates."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        menu = QMenu()
        update_action = menu.addAction(f"🔄 Update selected ({len(selected_rows)} specifications)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if action == update_action:
            # Extract the raw specification numbers (e.g., "TS 23.501" -> "23.501")
            target_specs = []
            for index in selected_rows:
                display_text = self.table.item(index.row(), 0).text()
                spec_num = display_text.split(" ")[-1]  # Gets the last part after the space
                target_specs.append(spec_num)

            force_meta = self.force_meta_checkbox.isChecked()
            self.update_specific_requested.emit(target_specs, force_meta)

    def refresh_table(self):
        # ... (Keep the rest of your refresh_table method exactly the same) ...
        spec_query = self.spec_search_input.text().strip()
        version_query = self.version_search_input.text().strip()

        specs = self.db.search_files(
            spec_number=spec_query if spec_query else None,
            release_version=version_query if version_query else None
        )

        self.table.setRowCount(0)

        grouped_specs = {}
        for row in specs:
            series, spec_num, title, spec_type, filename, version, url = row
            if spec_num not in grouped_specs:
                grouped_specs[spec_num] = {
                    'title': title,
                    'type': spec_type if spec_type else "",
                    'versions': []
                }
            grouped_specs[spec_num]['versions'].append((version, url, filename))

        self.count_label.setText(f"Showing {len(grouped_specs)} specifications")

        for row_idx, (spec_num, data) in enumerate(grouped_specs.items()):
            self.table.insertRow(row_idx)

            display_num = f"{data['type']} {spec_num}".strip()
            self.table.setItem(row_idx, 0, QTableWidgetItem(display_num))
            self.table.setItem(row_idx, 1, QTableWidgetItem(data['title'] if data['title'] else "Unknown Title"))

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

            self.table.setCellWidget(row_idx, 2, cell_widget)