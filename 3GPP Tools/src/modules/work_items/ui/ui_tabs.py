from pathlib import Path

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QTableView, QHeaderView, QPushButton, QProgressBar,
                             QMessageBox, QLineEdit)

from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import WorkItemsScraperThread


class WorkItemsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["Code", "Acronym", "Name", "Latest WID", "Release", "Start Date", "End Date"]

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if role == Qt.DisplayRole or role == Qt.UserRole:
            key_map = {
                "Code": "code", "Acronym": "acronym", "Name": "name",
                "Latest WID": "latest_wid", "Release": "release",
                "Start Date": "start_date", "End Date": "end_date"
            }
            val = row.get(key_map.get(col_name, ""), "")
            return str(val).strip() if val is not None else ""
        elif role == Qt.TextAlignmentRole:
            if col_name == "Name": return Qt.AlignLeft | Qt.AlignVCenter
            return Qt.AlignCenter
        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        return None

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()


class WorkItemsTab(QWidget):
    def __init__(self, db_path: Path):
        super().__init__()
        self.db_path = db_path
        self.db = WorkItemsDatabase(db_path)

        # Debounce timer for search
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(400)
        self.search_timer.timeout.connect(self.refresh_table)

        self._setup_ui()
        self._populate_filters()
        self.refresh_table()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # --- HEADER & CONTROLS ---
        header_layout = QHBoxLayout()

        header_lbl = QLabel("<b>📋 3GPP Work Items (WIs)</b>")
        header_lbl.setStyleSheet("font-size: 16px; color: #333;")

        self.sync_btn = QPushButton("🔄 Sync 3GPP WIs")
        self.sync_btn.setStyleSheet("""
            QPushButton { font-weight: bold; background-color: #0078D7; color: white; padding: 5px 15px; border-radius: 4px; }
            QPushButton:hover { background-color: #005A9E; }
            QPushButton:disabled { background-color: #A0C0E0; }
        """)
        self.sync_btn.setToolTip("Click to download and synchronize 3GPP Work Items in parallel from the server.")
        self.sync_btn.clicked.connect(self._start_sync)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #666; font-style: italic;")

        header_layout.addWidget(header_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.status_lbl)
        header_layout.addWidget(self.progress_bar)
        header_layout.addWidget(self.sync_btn)

        main_layout.addLayout(header_layout)

        # --- INLINE SEARCH & FILTER BAR ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("<b>🔍 Local Search:</b>"))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Code, Acronym, or Name...")
        self.search_input.setToolTip("Filter the table instantly by typing keywords.")
        self.search_input.textChanged.connect(lambda text: self.search_timer.start())
        search_layout.addWidget(self.search_input)

        # ---> NEW: Use CheckableComboBox for multi-select
        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(170)  # Set your desired width in pixels here
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(150)  # Set your desired width in pixels here
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.wg_combo)

        main_layout.addLayout(search_layout)

        # --- RESULTS COUNTER ---
        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet("font-weight: bold; color: #555555; margin-top: 5px;")
        count_layout = QHBoxLayout()
        count_layout.addStretch()
        count_layout.addWidget(self.count_label)
        main_layout.addLayout(count_layout)

        # --- TABLE VIEW ---
        self.table = QTableView()
        self.table_model = WorkItemsTableModel()
        self.table.setModel(self.table_model)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(
            "QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; }"
            "QTableView::item:selected { background-color: #cce8ff; color: #000; }"
        )

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)

        main_layout.addWidget(self.table)

    def _populate_filters(self):
        """Fetches options from the DB and populates the UI dropdowns."""
        options = self.db.get_filter_options()

        # CheckableComboBox uses updateItems() to seamlessly populate the list
        self.release_combo.blockSignals(True)
        self.release_combo.updateItems(options.get('releases', []))
        self.release_combo.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.updateItems(options.get('groups', []))
        self.wg_combo.blockSignals(False)

    def refresh_table(self):
        """Pulls the latest data from the database using active filters."""
        search_term = self.search_input.text().strip()

        # Retrieve the selected items as lists
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()

        # Execute query
        data = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs
        )

        self.table_model.update_data(data)

        # Update counter
        count = len(data)
        self.count_label.setText(f"Showing {count} Work Items")

    def _start_sync(self):
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Syncing...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.scraper_thread = WorkItemsScraperThread(self.db_path, self)
        self.scraper_thread.progress.connect(self._update_progress)
        self.scraper_thread.finished_sync.connect(self._on_sync_finished)
        self.scraper_thread.start()

    def _update_progress(self, current: int, total: int, msg: str):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_lbl.setText(msg)

    def _on_sync_finished(self, success: bool, msg: str):
        self.sync_btn.setEnabled(True)
        self.sync_btn.setText("🔄 Sync 3GPP WIs")
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("")

        # Refresh dropdowns and table with new data
        self._populate_filters()
        self.refresh_table()

        if success:
            QMessageBox.information(self, "Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Sync Failed", msg)