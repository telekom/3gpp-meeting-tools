# --- File: modules/specs_db/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QComboBox, QMenu, QAbstractItemView)
from PyQt5.QtCore import pyqtSignal, Qt, QTimer

from modules.specifications.core.database import SpecsDatabase

class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
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
        # Reduce margins to push the table closer to the edges
        main_layout.setContentsMargins(5, 10, 5, 5)

        # --- COMPACT TOP PANEL: Sync & Search on ONE line ---
        top_layout = QHBoxLayout()

        self.update_btn = QPushButton("🔄 Sync DB")
        self.update_btn.setToolTip("Run a full synchronization of the 3GPP database.")
        self.update_btn.clicked.connect(lambda: self.update_db_requested.emit(self.force_meta_checkbox.isChecked()))
        top_layout.addWidget(self.update_btn)

        self.force_meta_checkbox = QCheckBox("Force Metadata")
        self.force_meta_checkbox.setToolTip("If unchecked, only fetches metadata for new specifications.")
        top_layout.addWidget(self.force_meta_checkbox)

        top_layout.addSpacing(20) # Add a small visual gap

        top_layout.addWidget(QLabel("🔍 Spec:"))
        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("e.g. 23.501")
        self.spec_search_input.textChanged.connect(lambda text: self.search_timer.start())
        top_layout.addWidget(self.spec_search_input)

        top_layout.addWidget(QLabel("Ver:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15.")
        self.version_search_input.textChanged.connect(lambda text: self.search_timer.start())
        top_layout.addWidget(self.version_search_input)

        main_layout.addLayout(top_layout)

        # --- MIDDLE PANEL: Results Header ---
        self.count_label = QLabel("Showing 0 specifications")
        self.count_label.setStyleSheet("font-weight: bold; color: #555555;")

        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Database Results</b>"))
        header_layout.addStretch()
        header_layout.addWidget(self.count_label)

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

        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

    def _show_context_menu(self, position):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        menu = QMenu()
        update_action = menu.addAction(f"🔄 Update selected ({len(selected_rows)} specifications)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if action == update_action:
            target_specs = []
            for index in selected_rows:
                display_text = self.table.item(index.row(), 0).text()
                spec_num = display_text.split(" ")[-1]
                target_specs.append(spec_num)

            force_meta = self.force_meta_checkbox.isChecked()
            self.update_specific_requested.emit(target_specs, force_meta)

    def refresh_table(self):
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()

            if not spec_query and not version_query:
                self.table.setRowCount(0)
                self.count_label.setText("⌨️ Type a specification number (e.g., 23.501) to begin searching...")
                self.count_label.setStyleSheet("font-weight: bold; color: #555555;")
                return

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

            total_found = len(grouped_specs)
            rendered_specs = list(grouped_specs.items())[:100]

            if total_found > 100:
                self.count_label.setText(
                    f"⚠️ Showing top 100 of {total_found} specifications. Keep typing to narrow down.")
                self.count_label.setStyleSheet("font-weight: bold; color: #D83B01;")
            else:
                self.count_label.setText(f"Showing {total_found} specifications")
                self.count_label.setStyleSheet("font-weight: bold; color: #555555;")

            for row_idx, (spec_num, data) in enumerate(rendered_specs):
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

        except Exception as e:
            print(f"Error during refresh_table: {e}")