# --- File: modules/meetings/ui/ui_tabs.py ---
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLineEdit, QComboBox, QTableView, QHeaderView,
                             QMenu, QLabel, QCheckBox, QDateEdit, QSplitter,
                             QMessageBox, QFrame)
from PyQt5.QtCore import Qt, pyqtSignal, QDate, QPoint

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.specifications.ui.components import HoverMenuButton
from core.network.session import NetworkConfigDialog

# --- IMPORTS FROM REFACTORED FILES ---
from modules.meetings.ui.models import MeetingsTableModel
from modules.meetings.ui.dialogs import MeetingInfoDialog


# ==========================================
# --- MAIN UI TAB ---
# ==========================================
class MeetingsTab(QWidget):
    update_db_requested = pyqtSignal(bool, bool, bool)
    update_specific_requested = pyqtSignal(list, bool, bool, bool)

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
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(
            "QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; } QTableView::item:selected { background-color: #cce8ff; color: #000; }")

        # --- FIXED: Adjusted Column Widths for the new Layout ---
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.resizeSection(0, 40)   # Action Button
        header.resizeSection(1, 60)   # WG
        header.resizeSection(2, 90)   # Meeting Number
        header.setSectionResizeMode(3, QHeaderView.Stretch) # Location gets the remaining space
        header.resizeSection(4, 90)   # Start Date
        header.resizeSection(5, 90)   # End Date
        header.resizeSection(6, 110)  # First TDoc
        header.resizeSection(7, 110)  # Last TDoc

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_right_click_menu)

        left_layout.addWidget(self.table)
        self.splitter.addWidget(left_widget)

        # --- Right Side: Filter & Sync Panel ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)

        # 1. Filters
        title_lbl = QLabel("<b>Filter & Search</b>")
        title_lbl.setStyleSheet("font-size: 14px; margin-bottom: 5px;")
        right_layout.addWidget(title_lbl)

        right_layout.addWidget(QLabel("Working Group:"))
        self.wg_filter = QComboBox()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.currentTextChanged.connect(self.refresh_table)

        self.adhoc_filter = QComboBox()
        self.adhoc_filter.addItems(["All Meetings", "Regular", "Ad-Hoc / BIS"])
        self.adhoc_filter.currentTextChanged.connect(self.refresh_table)

        self.type_filter = QComboBox()
        self.type_filter.addItems(["All Types", "In-Person", "Electronic"])
        self.type_filter.currentTextChanged.connect(self.refresh_table)

        right_layout.addWidget(self.wg_filter)
        right_layout.addWidget(self.adhoc_filter)
        right_layout.addWidget(self.type_filter)

        right_layout.addWidget(QLabel("Search (No. or Name):"))
        self.search_input = QLineEdit()
        self.search_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.search_input)

        right_layout.addWidget(QLabel("Location:"))
        self.location_input = QLineEdit()
        self.location_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.location_input)

        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.toggled.connect(self._toggle_date_inputs)
        self.enable_dates_cb.toggled.connect(self.refresh_table)
        right_layout.addWidget(self.enable_dates_cb)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)
        right_layout.addWidget(self.date_from)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)
        right_layout.addWidget(self.date_to)
        self.enable_dates_cb.setChecked(True)

        # Separator line
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        right_layout.addWidget(line)

        # 2. Sync Configuration
        sync_lbl = QLabel("<b>Scrape Configuration</b>")
        sync_lbl.setStyleSheet("font-size: 14px; margin-top: 5px;")
        right_layout.addWidget(sync_lbl)

        self.chk_wg = QCheckBox("1. Check for New Folders")
        self.chk_wg.setChecked(True)
        self.chk_docs = QCheckBox("2. Deep Scrape 'Docs/'")
        self.chk_docs.setChecked(True)
        self.chk_dyna = QCheckBox("3. Update Metadata")
        self.chk_dyna.setChecked(True)

        right_layout.addWidget(self.chk_wg)
        right_layout.addWidget(self.chk_docs)
        right_layout.addWidget(self.chk_dyna)

        right_layout.addStretch()

        # 3. Actions
        self.update_btn = QPushButton("🔄 Sync All Meetings")
        self.update_btn.setStyleSheet("padding: 8px; font-weight: bold;")
        self.update_btn.clicked.connect(lambda: self.update_db_requested.emit(
            self.chk_wg.isChecked(), self.chk_docs.isChecked(), self.chk_dyna.isChecked()
        ))
        right_layout.addWidget(self.update_btn)

        self.delete_all_btn = QPushButton("🗑️ Clear All Meetings")
        self.delete_all_btn.setStyleSheet("padding: 8px; font-weight: bold; color: #D83B01;")
        self.delete_all_btn.clicked.connect(self._confirm_delete_all)
        right_layout.addWidget(self.delete_all_btn)

        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([750, 250])
        main_layout.addWidget(self.splitter)
        self._populate_filters()

    # --- NETWORK LOGIC ---
    def _open_network_config(self):
        NetworkConfigDialog(self).exec_()

    # --- DELETION LOGIC ---
    def _confirm_delete_all(self):
        if QMessageBox.question(self, 'Confirm Clear', "Delete ALL meetings? Cannot be undone.",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_all_meetings()
            self.refresh_table()

    def _confirm_delete_specific(self, targets: list):
        if QMessageBox.question(self, 'Confirm', f"Delete {len(targets)} meeting(s)?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_specific_meetings(targets)
            self.refresh_table()

    def _emit_multi_delete(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets: self._confirm_delete_specific(targets)

    # --- INTERACTION LOGIC ---
    def _toggle_date_inputs(self, checked):
        self.date_from.setEnabled(checked)
        self.date_to.setEnabled(checked)

    def _populate_filters(self):
        current_wg = self.wg_filter.currentText()
        wgs = self.db.get_working_groups()

        self.wg_filter.blockSignals(True)
        self.wg_filter.clear()
        self.wg_filter.addItem("All WGs")
        self.wg_filter.addItems(wgs)

        idx = self.wg_filter.findText(current_wg)
        if idx >= 0:
            self.wg_filter.setCurrentIndex(idx)

        self.wg_filter.blockSignals(False)

    def refresh_table(self):
        wg = self.wg_filter.currentText()
        date_from = self.date_from.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None
        date_to = self.date_to.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None

        adhoc_val = self.adhoc_filter.currentText()
        type_val = self.type_filter.currentText()

        data = self.db.search_meetings(
            wg_name=wg,
            search_term=self.search_input.text().strip(),
            location=self.location_input.text().strip(),
            date_from=date_from,
            date_to=date_to,
            adhoc_filter=adhoc_val,
            type_filter=type_val
        )

        self.table_model.update_data(data)
        for row_idx, row_data in enumerate(data):
            self._inject_hover_menu(row_idx, row_data)

    def _emit_multi_sync(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets: self.update_specific_requested.emit(targets, self.chk_wg.isChecked(), self.chk_docs.isChecked(),
                                                        self.chk_dyna.isChecked())

    def _populate_dynamic_menu(self, menu: QMenu, row_data: dict, row_idx: int):
        menu.clear()
        selected_rows = self.table.selectionModel().selectedRows()
        if len(selected_rows) > 1 and any(r.row() == row_idx for r in selected_rows):
            menu.addAction(f"🔄 Sync selected ({len(selected_rows)} meetings)").triggered.connect(
                lambda _, rows=selected_rows: self._emit_multi_sync(rows))
            menu.addSeparator()
            menu.addAction(f"🗑️ Delete selected ({len(selected_rows)} meetings)").triggered.connect(
                lambda _, rows=selected_rows: self._emit_multi_delete(rows))
        else:
            menu.addAction("ℹ️ Meeting Info").triggered.connect(lambda _, d=row_data: self.show_meeting_info(d))
            menu.addAction("🔄 Sync this Meeting").triggered.connect(lambda: self.update_specific_requested.emit(
                [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}], self.chk_wg.isChecked(),
                self.chk_docs.isChecked(), self.chk_dyna.isChecked()))
            menu.addSeparator()

            raw_url = row_data.get("url_key", "")
            if raw_url:
                full_ftp_url = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url}"
                menu.addAction("🌐 Open Main Folder (FTP)").triggered.connect(
                    lambda _, u=full_ftp_url: webbrowser.open(u))

            docs_url = row_data.get("docs_folder_url")
            if docs_url:
                menu.addAction("📂 Open Documents Folder").triggered.connect(lambda _, u=docs_url: webbrowser.open(u))

            wg_name = row_data.get("wg_name", "")
            meeting_name = row_data.get("name", "")
            start_date = row_data.get("start_date", "")
            end_date = row_data.get("end_date", "")
            is_elec = row_data.get("is_electronic", 0)

            if self.db.is_active_sync_meeting(wg_name, start_date, end_date, is_elec):
                menu.addSeparator()

                # Handle the special SA3-LI edge case
                sync_wg = "SA3LI" if wg_name == "SA3" and "LI" in meeting_name.upper() else wg_name
                sync_base_url = f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}"

                menu.addAction("🔄 Open SYNC folder (FTP)").triggered.connect(
                    lambda _, u=sync_base_url: webbrowser.open(u))
                menu.addAction("📂 Open SYNC Documents folder").triggered.connect(
                    lambda _, u=f"{sync_base_url}/Docs": webbrowser.open(u))

            menu.addSeparator()
            menu.addAction("🗑️ Delete this Meeting").triggered.connect(lambda: self._confirm_delete_specific(
                [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}]))

    def _inject_hover_menu(self, row_idx: int, row_data: dict):
        action_btn = HoverMenuButton("⋮")
        action_btn.setFixedSize(24, 24)
        action_btn.setCursor(Qt.PointingHandCursor)
        action_btn.setStyleSheet(
            "QPushButton { border: none; background: transparent; color: #555; font-size: 20px; font-weight: bold; padding-bottom: 4px; } QPushButton:hover { color: #0078D7; } QPushButton::menu-indicator { image: none; width: 0px; }")

        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background-color: #FAFAFA; border: 1px solid #CCC; } QMenu::item { padding: 5px 20px 5px 15px; color: #333333; } QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }")
        menu.aboutToShow.connect(lambda m=menu, d=row_data, i=row_idx: self._populate_dynamic_menu(m, d, i))
        action_btn.setMenu(menu)

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(action_btn)
        self.table.setIndexWidget(self.table_model.index(row_idx, 0), container)

    def show_right_click_menu(self, pos: QPoint):
        index = self.table.indexAt(pos)
        if index.isValid():
            row_data = self.table_model.data(index, Qt.UserRole)
            if not row_data: return
            menu = QMenu(self)
            menu.setStyleSheet(
                "QMenu { background-color: #FAFAFA; border: 1px solid #CCC; } QMenu::item { padding: 5px 20px 5px 15px; color: #333333; } QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }")
            self._populate_dynamic_menu(menu, row_data, index.row())
            menu.exec_(self.table.viewport().mapToGlobal(pos))

    def show_meeting_info(self, data: dict):
        MeetingInfoDialog(data, self).exec_()