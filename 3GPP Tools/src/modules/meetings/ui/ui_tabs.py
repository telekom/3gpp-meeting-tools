# --- File: modules/meetings/ui/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLineEdit, QComboBox, QTableView, QHeaderView,
                             QMessageBox, QMenu)
from PyQt5.QtCore import Qt, pyqtSignal, QAbstractTableModel, QModelIndex, QPoint

from modules.meetings.core.meetings_db import MeetingsDatabase


class MeetingsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["WG", "Meeting", "Name", "Location", "Start Date", "End Date", ""]

    def data(self, index, role):
        if not index.isValid():
            return None

        row_data = self._data[index.row()]

        if role == Qt.DisplayRole:
            col = index.column()
            if col == 0:
                return row_data.get("wg_name", "")
            elif col == 1:
                return row_data.get("meeting_number", "")
            elif col == 2:
                return row_data.get("name", "")
            elif col == 3:
                return row_data.get("location", "")
            elif col == 4:
                return row_data.get("start_date", "")
            elif col == 5:
                return row_data.get("end_date", "")
            elif col == 6:
                return "⋮"  # Kebab menu column

        elif role == Qt.TextAlignmentRole:
            if index.column() in [0, 1, 4, 5, 6]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.UserRole:
            return row_data  # Return full dict for Kebab menu actions

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


class MeetingsTab(QWidget):
    update_db_requested = pyqtSignal()

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self._setup_ui()
        self.refresh_table()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # --- Top Filter Bar ---
        filter_layout = QHBoxLayout()

        self.wg_filter = QComboBox()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.currentIndexChanged.connect(self.refresh_table)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search by Meeting No, Name, or Location...")
        self.search_input.textChanged.connect(self.refresh_table)

        self.update_btn = QPushButton("🔄 Sync Meetings")
        self.update_btn.clicked.connect(self.update_db_requested.emit)

        filter_layout.addWidget(self.wg_filter)
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.update_btn)
        layout.addLayout(filter_layout)

        # --- Table View ---
        self.table = QTableView()
        self.table_model = MeetingsTableModel()
        self.table.setModel(self.table_model)

        # Table Styling
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; }
            QTableView::item:selected { background-color: #cce8ff; color: #000; }
        """)

        # Column Sizing
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 60)  # WG
        header.resizeSection(1, 80)  # Meeting Num
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Name
        header.resizeSection(3, 150)  # Location
        header.resizeSection(4, 100)  # Start
        header.resizeSection(5, 100)  # End
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 40)  # Kebab

        # Connect Context Menu for Kebab clicks
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_kebab_menu)
        self.table.clicked.connect(self.on_table_clicked)

        layout.addWidget(self.table)
        self._populate_filters()

    def _populate_filters(self):
        """Loads available Working Groups into the dropdown."""
        wgs = self.db.get_working_groups()
        self.wg_filter.blockSignals(True)
        self.wg_filter.clear()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.addItems(wgs)
        self.wg_filter.blockSignals(False)

    def refresh_table(self):
        """Fetches data from DB based on current filters."""
        wg = self.wg_filter.currentText() if self.wg_filter.currentIndex() > 0 else None
        search = self.search_input.text().strip()

        data = self.db.search_meetings(wg_name=wg, search_term=search)
        self.table_model.update_data(data)

    def on_table_clicked(self, index):
        """Trigger the menu on left click if the kebab column is clicked."""
        if index.column() == 6:
            # Calculate position to show menu right under the cell
            rect = self.table.visualRect(index)
            pos = self.table.viewport().mapToGlobal(rect.bottomRight())
            self._show_menu(index, pos)

    def show_kebab_menu(self, pos: QPoint):
        """Trigger the menu on right click anywhere on the row."""
        index = self.table.indexAt(pos)
        if index.isValid():
            global_pos = self.table.viewport().mapToGlobal(pos)
            self._show_menu(index, global_pos)

    def _show_menu(self, index: QModelIndex, global_pos: QPoint):
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data: return

        menu = QMenu(self)

        # Info Action
        info_action = menu.addAction("ℹ️ Meeting Info")
        info_action.triggered.connect(lambda: self.show_meeting_info(row_data))

        menu.addSeparator()

        # Web Links
        url_key = row_data.get("url_key")
        docs_url = row_data.get("docs_folder_url")

        if url_key:
            page_action = menu.addAction("🌐 Open Main Folder (FTP)")
            page_action.triggered.connect(lambda: webbrowser.open(url_key))

        if docs_url:
            docs_action = menu.addAction("📂 Open Documents Folder")
            docs_action.triggered.connect(lambda: webbrowser.open(docs_url))

        menu.exec_(global_pos)

    def show_meeting_info(self, data: dict):
        """Displays a clean popup with all meeting data."""
        info = (
            f"<b>Meeting:</b> {data.get('wg_name')} {data.get('meeting_number')}<br><br>"
            f"<b>Name:</b> {data.get('name') or 'N/A'}<br>"
            f"<b>Location:</b> {data.get('location') or 'N/A'}<br>"
            f"<b>Dates:</b> {data.get('start_date')} to {data.get('end_date')}<br><br>"
            f"<b>First TDoc:</b> {data.get('first_tdoc') or 'Unknown'}<br>"
            f"<b>Last TDoc:</b> {data.get('last_tdoc') or 'Unknown'}"
        )
        QMessageBox.information(self, "Meeting Details", info)