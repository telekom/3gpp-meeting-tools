# --- File: modules/meetings/ui/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLineEdit, QComboBox, QTableView, QHeaderView,
                             QMessageBox, QMenu, QLabel, QCheckBox, QDateEdit, QSplitter)
from PyQt5.QtCore import Qt, pyqtSignal, QAbstractTableModel, QModelIndex, QPoint, QDate

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
                return "⋮"

        elif role == Qt.TextAlignmentRole:
            if index.column() in [0, 1, 4, 5, 6]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.UserRole:
            return row_data

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
    # --- SIGNALS MUST BE DEFINED HERE AT THE CLASS LEVEL ---
    update_db_requested = pyqtSignal()
    update_specific_requested = pyqtSignal(list)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self._setup_ui()
        self.refresh_table()

    def _setup_ui(self):
        main_layout = QHBoxLayout(self)
        self.splitter = QSplitter(Qt.Horizontal)

        # --- Left Side: Table View ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.table = QTableView()
        self.table_model = MeetingsTableModel()
        self.table.setModel(self.table_model)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; }
            QTableView::item:selected { background-color: #cce8ff; color: #000; }
        """)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 60)
        header.resizeSection(1, 80)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.resizeSection(3, 150)
        header.resizeSection(4, 90)
        header.resizeSection(5, 90)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 40)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_kebab_menu)
        self.table.clicked.connect(self.on_table_clicked)

        left_layout.addWidget(self.table)
        self.splitter.addWidget(left_widget)

        # --- Right Side: Filter Panel ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)

        title_lbl = QLabel("<b>Filter & Sync</b>")
        title_lbl.setStyleSheet("font-size: 14px; margin-bottom: 10px;")
        right_layout.addWidget(title_lbl)

        right_layout.addWidget(QLabel("Working Group:"))
        self.wg_filter = QComboBox()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.currentIndexChanged.connect(self.refresh_table)
        right_layout.addWidget(self.wg_filter)

        right_layout.addWidget(QLabel("Search (No. or Name):"))
        self.search_input = QLineEdit()
        self.search_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.search_input)

        right_layout.addWidget(QLabel("Location:"))
        self.location_input = QLineEdit()
        self.location_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.location_input)

        right_layout.addSpacing(10)
        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.toggled.connect(self._toggle_date_inputs)
        self.enable_dates_cb.toggled.connect(self.refresh_table)
        right_layout.addWidget(self.enable_dates_cb)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)
        right_layout.addWidget(QLabel("Start Date (From):"))
        right_layout.addWidget(self.date_from)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)
        right_layout.addWidget(QLabel("End Date (To):"))
        right_layout.addWidget(self.date_to)

        right_layout.addStretch()

        self.update_btn = QPushButton("🔄 Sync All Meetings")
        self.update_btn.setStyleSheet("padding: 8px; font-weight: bold;")
        self.update_btn.clicked.connect(self.update_db_requested.emit)
        right_layout.addWidget(self.update_btn)

        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([750, 250])

        main_layout.addWidget(self.splitter)
        self._populate_filters()

    def _toggle_date_inputs(self, checked):
        self.date_from.setEnabled(checked)
        self.date_to.setEnabled(checked)

    def _populate_filters(self):
        wgs = self.db.get_working_groups()
        self.wg_filter.blockSignals(True)
        self.wg_filter.clear()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.addItems(wgs)
        self.wg_filter.blockSignals(False)

    def refresh_table(self):
        wg = self.wg_filter.currentText()
        search = self.search_input.text().strip()
        location = self.location_input.text().strip()

        date_from = None
        date_to = None

        if self.enable_dates_cb.isChecked():
            date_from = self.date_from.date().toString("yyyy-MM-dd")
            date_to = self.date_to.date().toString("yyyy-MM-dd")

        data = self.db.search_meetings(
            wg_name=wg,
            search_term=search,
            location=location,
            date_from=date_from,
            date_to=date_to
        )
        self.table_model.update_data(data)

    def on_table_clicked(self, index):
        if index.column() == 6:
            rect = self.table.visualRect(index)
            pos = self.table.viewport().mapToGlobal(rect.bottomRight())
            self._show_menu(index, pos)

    def show_kebab_menu(self, pos: QPoint):
        index = self.table.indexAt(pos)
        if index.isValid():
            global_pos = self.table.viewport().mapToGlobal(pos)
            self._show_menu(index, global_pos)

    def _show_menu(self, index: QModelIndex, global_pos: QPoint):
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data: return

        menu = QMenu(self)

        info_action = menu.addAction("ℹ️ Meeting Info")
        info_action.triggered.connect(lambda: self.show_meeting_info(row_data))

        # --- THE SPECIFIC UPDATE ACTION IS NOW PROPERLY WIRED ---
        update_action = menu.addAction("🔄 Sync this Meeting")
        update_action.triggered.connect(lambda: self.update_specific_requested.emit([
            {"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}
        ]))

        menu.addSeparator()

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
        info = (
            f"<b>Meeting:</b> {data.get('wg_name')} {data.get('meeting_number')}<br><br>"
            f"<b>Name:</b> {data.get('name') or 'N/A'}<br>"
            f"<b>Location:</b> {data.get('location') or 'N/A'}<br>"
            f"<b>Dates:</b> {data.get('start_date')} to {data.get('end_date')}<br><br>"
            f"<b>First TDoc:</b> {data.get('first_tdoc') or 'Unknown'}<br>"
            f"<b>Last TDoc:</b> {data.get('last_tdoc') or 'Unknown'}"
        )
        QMessageBox.information(self, "Meeting Details", info)