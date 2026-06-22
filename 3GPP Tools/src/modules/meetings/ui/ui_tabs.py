# --- File: modules/meetings/ui/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLineEdit, QComboBox, QTableView, QHeaderView,
                             QMenu, QLabel, QCheckBox, QDateEdit, QSplitter, QDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QAbstractTableModel, QModelIndex, QDate

from modules.meetings.core.meetings_db import MeetingsDatabase
# Reuse the exact HoverMenuButton from the specifications module
from modules.specifications.ui.components import HoverMenuButton


# ==========================================
# --- HELPER & DIALOG: MEETING INFO ---
# ==========================================
def _format_meeting_info(data: dict) -> str:
    """Safely formats all database parameters into a clean HTML block, handling None values."""
    if not data: return ""

    def clean(val): return str(val) if val else "N/A"

    return (
        f"<b>Working Group:</b> {clean(data.get('wg_name'))}<br>"
        f"<b>Meeting Number:</b> {clean(data.get('meeting_number'))}<br>"
        f"<b>Meeting Name:</b> {clean(data.get('name'))}<br>"
        f"<b>Location:</b> {clean(data.get('location'))}<br>"
        f"<b>Dates:</b> {clean(data.get('start_date'))} to {clean(data.get('end_date'))}<br><hr>"
        f"<b>First TDoc:</b> {clean(data.get('first_tdoc'))}<br>"
        f"<b>Last TDoc:</b> {clean(data.get('last_tdoc'))}<br><hr>"
        f"<b>Main FTP Link:</b> {clean(data.get('url_key'))}<br>"
        f"<b>Docs Folder Link:</b> {clean(data.get('docs_folder_url'))}<br>"
        f"<b>Database ID:</b> {clean(data.get('id'))}"
    )


class MeetingInfoDialog(QDialog):
    """A silent QDialog to show meeting info without triggering the Windows alert sound."""

    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        title_str = f"{data.get('wg_name', '')} {data.get('meeting_number', '')}".strip()
        self.setWindowTitle(f"Meeting Details: {title_str}")
        self.setMinimumWidth(450)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; }")

        layout = QVBoxLayout(self)

        info_label = QLabel(_format_meeting_info(data))
        info_label.setWordWrap(True)
        info_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(info_label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)


# ==========================================
# --- TABLE MODEL ---
# ==========================================
class MeetingsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["", "WG", "Meeting", "Name", "Location", "Start Date", "End Date"]

    def data(self, index, role):
        if not index.isValid():
            return None

        row_data = self._data[index.row()]

        if role == Qt.DisplayRole:
            col = index.column()
            if col == 0:
                return ""  # Left empty for the HoverMenuButton
            elif col == 1:
                return row_data.get("wg_name", "")
            elif col == 2:
                return row_data.get("meeting_number", "")
            elif col == 3:
                return row_data.get("name", "")
            elif col == 4:
                return row_data.get("location", "")
            elif col == 5:
                return row_data.get("start_date", "")
            elif col == 6:
                return row_data.get("end_date", "")

        elif role == Qt.TextAlignmentRole:
            if index.column() in [0, 1, 2, 5, 6]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.UserRole:
            return row_data

        elif role == Qt.ToolTipRole:
            return _format_meeting_info(row_data)

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


# ==========================================
# --- MAIN UI TAB ---
# ==========================================
class MeetingsTab(QWidget):
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
        self.table.setMouseTracking(True)
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
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.resizeSection(0, 40)  # Kebab
        header.resizeSection(1, 60)  # WG
        header.resizeSection(2, 80)  # Meeting Num
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Name
        header.resizeSection(4, 150)  # Location
        header.resizeSection(5, 90)  # Start
        header.resizeSection(6, 90)  # End

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

        # --- Rebuild the Hover Menus for the updated data ---
        for row_idx, row_data in enumerate(data):
            self._inject_hover_menu(row_idx, row_data)

    def _inject_hover_menu(self, row_idx: int, row_data: dict):
        """Creates and embeds the reusable HoverMenuButton into Column 0."""
        action_btn = HoverMenuButton("⋮")
        action_btn.setFixedSize(24, 24)
        action_btn.setCursor(Qt.PointingHandCursor)
        action_btn.setStyleSheet("""
            QPushButton { border: none; background: transparent; color: #555; font-size: 20px; font-weight: bold; padding-bottom: 4px; }
            QPushButton:hover { color: #0078D7; }
            QPushButton::menu-indicator { image: none; width: 0px; }
        """)

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #FAFAFA; border: 1px solid #CCC; } 
            QMenu::item { padding: 5px 20px 5px 15px; color: #333333; } 
            QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
        """)

        info_action = menu.addAction("ℹ️ Meeting Info")
        info_action.triggered.connect(lambda _, d=row_data: self.show_meeting_info(d))

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

        action_btn.setMenu(menu)

        # Center the button nicely in the cell
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(action_btn)

        self.table.setIndexWidget(self.table_model.index(row_idx, 0), container)

    def show_meeting_info(self, data: dict):
        """Silently displays the info dialog."""
        dialog = MeetingInfoDialog(data, self)
        dialog.exec_()