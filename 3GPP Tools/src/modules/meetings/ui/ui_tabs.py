# --- File: modules/meetings/ui/ui_tabs.py ---
import json
import os
import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, pyqtSignal, QDate, QPoint
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLineEdit, QComboBox, QTableView, QHeaderView,
                             QMenu, QLabel, QCheckBox, QDateEdit, QSplitter,
                             QMessageBox, QFrame, QFileDialog, QApplication)

import core.utils.paths
from core.network.session import NetworkConfigDialog
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.ui.dialogs import MeetingInfoDialog
from modules.meetings.ui.models import MeetingsTableModel
from modules.meetings.ui.tdocs_window import TDocsWindow
from modules.specifications.ui.components import HoverMenuButton
from modules.meetings.core.compare_manager import ComparisonManager
from modules.word_tools.core.word_comparator import WordComparatorThread
from modules.meetings.core.tdocs_threads import TDocActionThread


# ==========================================
# --- MAIN UI TAB ---
# ==========================================
class MeetingsTab(QWidget):
    update_db_requested = pyqtSignal(bool, bool, bool)
    update_specific_requested = pyqtSignal(list, bool, bool, bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = MeetingsDatabase(db_path)

        # --- Local Cache Configuration Setup ---
        self.config_file = core.utils.paths.get_project_root() / "meetings_config.json"
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        self.cache_dir = self._load_settings()

        self.tdoc_windows = {}  # <-- NEW: Keeps track of open windows by mtg_id
        self._active_dl_threads = {}  # <-- NEW: Keeps threads alive during execution

        self._setup_ui()
        self.refresh_table()

    # --- SETTINGS LOGIC ---
    def _load_settings(self) -> str:
        fallback = str(Path.home() / "3GPP_Delegate_Helper" / "cache")
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('download_dir', fallback)
            except Exception:
                pass
        return fallback

    def _save_settings(self):
        try:
            current_dir = self.dl_dir_input.text().strip()
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['download_dir'] = current_dir

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            self.cache_dir = current_dir
        except Exception as e:
            print(f"Error saving config: {e}")

    def _save_last_meeting(self, mtg_info: dict):
        """Silently logs the last accessed meeting to the config file."""
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['last_mtg_id'] = mtg_info.get("mtg_id")
            data['last_mtg_number'] = mtg_info.get("meeting_number")

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving last meeting state: {e}")

    def _browse_cache_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Cache Directory", self.dl_dir_input.text())
        if directory:
            # Normalize path slashes for the OS
            normalized_dir = str(Path(directory))
            self.dl_dir_input.setText(normalized_dir)
            self._save_settings()

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
        self.table.verticalHeader().setDefaultSectionSize(36)
        self.table.setStyleSheet(
            "QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; } QTableView::item:selected { background-color: #cce8ff; color: #000; }")

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)  # NEW TDocs Button Column
        header.resizeSection(0, 40)  # Action Button
        header.resizeSection(1, 90)  # TDocs Button
        header.resizeSection(2, 60)  # WG
        header.resizeSection(3, 90)  # Meeting Number
        header.setSectionResizeMode(4, QHeaderView.Stretch)  # Location gets the remaining space
        header.resizeSection(5, 90)  # Start Date
        header.resizeSection(6, 90)  # End Date
        header.resizeSection(7, 110)  # First TDoc
        header.resizeSection(8, 110)  # Last TDoc

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_right_click_menu)

        left_layout.addWidget(self.table)
        self.splitter.addWidget(left_widget)

        # ==========================================
        # ---> NEW: GLOBAL COMPARISON CART UI <---
        # ==========================================
        self.cart_frame = QFrame()
        self.cart_frame.setStyleSheet("""
                    QFrame { background-color: #E8F2FB; border: 1px solid #B0D0F0; border-radius: 8px; }
                    QLabel { color: #333; border: none; }
                """)
        cart_layout = QHBoxLayout(self.cart_frame)
        cart_layout.setContentsMargins(15, 10, 15, 10)

        cart_layout.addWidget(QLabel("<b>⚖️ Comparison Cart:</b>"))

        self.lbl_slot_a = QLabel("<i style='color:#777;'>[Slot A Empty]</i>")
        self.lbl_slot_b = QLabel("<i style='color:#777;'>[Slot B Empty]</i>")

        cart_layout.addSpacing(10)
        cart_layout.addWidget(self.lbl_slot_a)
        cart_layout.addWidget(QLabel(" <b>VS</b> "))
        cart_layout.addWidget(self.lbl_slot_b)
        cart_layout.addStretch()

        self.btn_compare = QPushButton("⚖️ Compare in Word")
        self.btn_compare.setEnabled(False)
        self.btn_compare.setStyleSheet(
            "QPushButton { font-weight: bold; background-color: #0078D7; color: white; padding: 5px 15px; border-radius: 4px; } QPushButton:disabled { background-color: #A0C0E0; }")
        self.btn_compare.clicked.connect(self._run_comparison)

        self.btn_clear_cart = QPushButton("✖ Clear")
        self.btn_clear_cart.setStyleSheet("QPushButton { color: #555; padding: 5px 10px; }")
        self.btn_clear_cart.clicked.connect(ComparisonManager.get_instance().clear_cart)

        cart_layout.addWidget(self.btn_compare)
        cart_layout.addWidget(self.btn_clear_cart)

        left_layout.addWidget(self.cart_frame)

        # Wire up the Singleton signals
        ComparisonManager.get_instance().cart_updated.connect(self._update_cart_ui)

        # ==========================================
        # ---> END OF: GLOBAL COMPARISON CART UI <---
        # ==========================================

        # --- Right Side: Filter & Sync Panel ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)

        self.btn_open_last = QPushButton("🚀 Open Last Meeting")
        self.btn_open_last.setStyleSheet("""
                    QPushButton {
                        font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; font-weight: bold;
                        background-color: #0078D7; color: white; border: none;
                        padding: 8px; border-radius: 6px; margin-bottom: 10px;
                    }
                    QPushButton:hover { 
                        background-color: #005A9E; 
                    }
                    QPushButton:pressed { 
                        background-color: #004578; 
                    }
                """)
        self.btn_open_last.clicked.connect(self._open_last_meeting)
        right_layout.addWidget(self.btn_open_last)

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

        # --- NEW: GLOBAL TDOC SEARCH ---
        right_layout.addWidget(QLabel("Global TDoc Search:"))
        global_search_layout = QHBoxLayout()
        self.global_tdoc_input = QLineEdit()
        self.global_tdoc_input.setPlaceholderText("e.g., S2-2605740")
        self.global_tdoc_input.returnPressed.connect(self._on_global_tdoc_search)

        self.global_tdoc_btn = QPushButton("🔍 Find TDoc")
        self.global_tdoc_btn.setCursor(Qt.PointingHandCursor)
        self.global_tdoc_btn.setStyleSheet("""
                    QPushButton { 
                        font-weight: bold; background-color: #0078D7; color: white; 
                        padding: 6px 15px; border-radius: 4px; 
                    } 
                    QPushButton:hover { background-color: #005A9E; }
                    QPushButton:pressed { background-color: #004578; }
                    QPushButton:disabled { background-color: #A0C0E0; }
                """)
        self.global_tdoc_btn.clicked.connect(self._on_global_tdoc_search)

        global_search_layout.addWidget(self.global_tdoc_input)
        global_search_layout.addWidget(self.global_tdoc_btn)
        right_layout.addLayout(global_search_layout)

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

        # 2. Compact Sync Configuration
        self.scrape_toggle_btn = QPushButton("⚙️ Scrape Configuration (Click to Expand)")
        self.scrape_toggle_btn.setCheckable(True)
        self.scrape_toggle_btn.setCursor(Qt.PointingHandCursor)
        self.scrape_toggle_btn.setStyleSheet("""
                    QPushButton { text-align: left; padding: 5px; border: none; font-weight: bold; color: #555; }
                    QPushButton:hover { color: #0078D7; }
                """)

        self.scrape_frame = QFrame()
        self.scrape_frame.setVisible(False)  # Hidden by default
        scrape_layout = QVBoxLayout(self.scrape_frame)
        scrape_layout.setContentsMargins(15, 0, 0, 5)

        self.chk_wg = QCheckBox("Check for New Folders")
        self.chk_wg.setChecked(True)
        self.chk_dyna = QCheckBox("Update Metadata")
        self.chk_dyna.setChecked(True)
        self.chk_docs = QCheckBox("Deep Scrape 'Docs/'")
        self.chk_docs.setChecked(True)

        scrape_layout.addWidget(self.chk_wg)
        scrape_layout.addWidget(self.chk_dyna)
        scrape_layout.addWidget(self.chk_docs)

        self.scrape_toggle_btn.toggled.connect(self.scrape_frame.setVisible)

        right_layout.addWidget(self.scrape_toggle_btn)
        right_layout.addWidget(self.scrape_frame)

        # --- NEW: Local Cache GUI Element ---
        right_layout.addWidget(QLabel("Local Cache Directory:"))
        cache_layout = QHBoxLayout()
        self.dl_dir_input = QLineEdit()
        self.dl_dir_input.setText(self.cache_dir)
        self.dl_dir_input.editingFinished.connect(self._save_settings)

        browse_btn = QPushButton("...")
        browse_btn.setFixedWidth(30)
        browse_btn.clicked.connect(self._browse_cache_dir)

        cache_layout.addWidget(self.dl_dir_input)
        cache_layout.addWidget(browse_btn)
        right_layout.addLayout(cache_layout)
        # ------------------------------------

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

    def _update_cart_ui(self, slot_a: dict, slot_b: dict):
        self.lbl_slot_a.setText(
            f"<b style='color:#005A9E;'>{slot_a['name']}</b>" if slot_a else "<i style='color:#777;'>[Slot A Empty]</i>")
        self.lbl_slot_b.setText(
            f"<b style='color:#005A9E;'>{slot_b['name']}</b>" if slot_b else "<i style='color:#777;'>[Slot B Empty]</i>")
        self.btn_compare.setEnabled(bool(slot_a and slot_b))

    def _run_comparison(self):
        mgr = ComparisonManager.get_instance()
        if mgr.slot_a and mgr.slot_b:
            self.btn_compare.setText("⏳ Comparing...")
            self.btn_compare.setEnabled(False)

            # Fire the existing Word Comparator Thread!
            self.cmp_thread = WordComparatorThread(mgr.slot_a['path'], mgr.slot_b['path'])
            self.cmp_thread.ui_log_msg.connect(self._handle_compare_log)
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setText("⚖️ Compare in Word"))
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setEnabled(True))
            self.cmp_thread.start()

    def _handle_compare_log(self, msg: str, level: int):
        import logging
        if level == logging.ERROR:
            print(f"🔴 {msg}")
        elif level == logging.WARNING:
            print(f"🟠 {msg}")
        else:
            print(f"🔵 {msg}")

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
            location=None,  # <--- THE FIX: Set this to None since we removed the text box!
            date_from=date_from,
            date_to=date_to,
            adhoc_filter=adhoc_val,
            type_filter=type_val
        )

        self.table_model.update_data(data)
        for row_idx, row_data in enumerate(data):
            self._inject_hover_menu(row_idx, row_data)
            self._inject_tdocs_button(row_idx, row_data)

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

            # --- LOCAL CACHE & TDOCS DOWNLOAD INTEGRATION ---
            folder_name = row_data.get("folder_name")
            if not folder_name:
                folder_name = row_data.get("meeting_number", "")

            if folder_name:
                # Bulletproof fallback: Read directly from the UI text box if self.cache_dir is missing
                current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.cache_dir
                local_path = Path(current_cache) / folder_name

                # 1. Existing: Open Local Cache
                if local_path.exists() and local_path.is_dir():
                    menu.addAction("📁 Open Local Cache Folder").triggered.connect(
                        lambda _, p=str(local_path): os.startfile(p) if hasattr(os, 'startfile') else webbrowser.open(
                            f"file:///{p}")
                    )

                # 2. UPGRADED: Trigger the inline TDocs button directly!
                mtg_id = row_data.get("mtg_id")
                if mtg_id:
                    # Grab the container widget we injected into Column 1
                    container = self.table.indexWidget(self.table_model.index(row_idx, 1))
                    if container:
                        # Find the actual QPushButton inside the container
                        tdocs_btn = container.findChild(QPushButton)
                        if tdocs_btn and tdocs_btn.isEnabled():
                            menu.addAction("📗 Open TDocs List").triggered.connect(tdocs_btn.click)

                # 3. NEW: Bulk Cache TDocs
                if docs_url:
                    menu.addAction("📥 Cache TDocs (Docs/)").triggered.connect(
                        lambda _, u=docs_url, p=local_path: self._start_tdocs_caching(u, p)
                    )
            # ------------------------------------------------

            wg_name = row_data.get("wg_name", "")
            meeting_name = row_data.get("name", "")
            start_date = row_data.get("start_date", "")
            end_date = row_data.get("end_date", "")
            is_elec = row_data.get("is_electronic", 0)

            if self.db.is_active_sync_meeting(wg_name, start_date, end_date, is_elec):
                menu.addSeparator()

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

    def _start_tdocs_download(self, mtg_id: str, local_path: Path):
        # Notify the user that the download has started
        self.update_btn.setText("⏳ Opening TDocs...")
        self.update_btn.setEnabled(False)

        # Initialize and start the background worker
        self.dl_thread = TDocsDownloaderThread(mtg_id, local_path)
        self.dl_thread.finished.connect(self._on_tdocs_download_finished)
        self.dl_thread.start()

    def _on_tdocs_download_finished(self, success: bool, result: str):
        # Reset the UI button
        self.update_btn.setText("🔄 Sync All Meetings")
        self.update_btn.setEnabled(True)

        if success:
            # Result contains the filepath. Open it natively in Excel.
            try:
                if hasattr(os, 'startfile'):
                    os.startfile(result)
                else:
                    webbrowser.open(f"file:///{result}")
            except Exception as e:
                QMessageBox.warning(self, "Open Error", f"Could not open the downloaded file:\n{e}")
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to download TDocs List:\n{result}")

    # ==========================================
    # --- NEW: TDOC BUTTON & WINDOW LOGIC ---
    # ==========================================
    def _get_tdoc_list_path(self, row_data: dict) -> Path:
        mtg_id = row_data.get("mtg_id")
        if not mtg_id: return None

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")

        agenda_dir = Path(current_cache) / folder_name / "Agenda"

        # FIXED: Dynamically search the folder for existing files!
        if agenda_dir.exists() and agenda_dir.is_dir():
            for file_path in agenda_dir.iterdir():
                filename = file_path.name.lower()
                # Catch both the new names and the old server fallback names
                if (filename.startswith("tdoc_list_meeting_") or filename.startswith(
                        "tdocs_list_")) and filename.endswith(".xlsx"):
                    return file_path

        # If no file exists, return a fallback path for the downloader to use
        return agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"

    def _download_and_open_tdocs(self, row_data: dict, btn: QPushButton):
        mtg_id = row_data.get("mtg_id")

        # Shortened text and unified CSS
        btn.setText("⏳ Fetching")
        btn.setStyleSheet("""
                    QPushButton {
                        font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; font-weight: bold;
                        border-radius: 12px; padding: 2px 6px;
                        color: #B85C00; background-color: #FFF4CE; border: 1px solid #F3C74C;
                    }
                """)
        btn.setEnabled(False)

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        local_path = Path(current_cache) / folder_name

        thread = TDocsDownloaderThread(mtg_id, local_path, self)
        self._active_dl_threads[mtg_id] = thread
        thread.finished.connect(
            lambda success, res, m_id: self._on_inline_download_finished(success, res, m_id, row_data))
        thread.start()

    def _inject_tdocs_button(self, row_idx: int, row_data: dict):
        mtg_id = row_data.get("mtg_id")
        filepath = self._get_tdoc_list_path(row_data)

        btn = QPushButton()
        btn.setFixedHeight(24)
        btn.setCursor(Qt.PointingHandCursor)

        # STATE 1: Unobtainable (No 3GPP ID)
        if not mtg_id:
            btn.setText("N/A")
            btn.setToolTip("Missing 3GPP Portal ID (Cannot fetch TDocs)")
            btn.setStyleSheet("""
                QPushButton {
                    font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; font-weight: bold;
                    border-radius: 12px; padding: 2px 6px;
                    color: #7A7A7A; background-color: #F0F0F0; border: 1px solid #D1D1D1;
                }
            """)
            btn.setEnabled(False)

        # STATE 2: Cached Locally (Ready to Open)
        elif filepath and filepath.exists():
            # Using a bold unicode checkmark that inherits our green CSS color
            btn.setText("✔ Open")
            btn.setToolTip("TDocs are cached locally. Click to view table.")
            btn.setStyleSheet("""
                QPushButton {
                    font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; font-weight: bold;
                    border-radius: 12px; padding: 2px 6px;
                    color: #0C6B0C; background-color: #E6F4E6; border: 1px solid #A3DDA3;
                }
                QPushButton:hover {
                    background-color: #D1EED1; border: 1px solid #0C6B0C; color: #0C6B0C;
                }
            """)
            btn.clicked.connect(lambda _, d=row_data, f=str(filepath): self._open_tdocs_window(d, f))

        # STATE 3: Missing (Needs Download)
        else:
            # Shortened text to "Get" so it easily fits the column width!
            btn.setText("⬇ Get")
            btn.setToolTip("Not cached. Click to download TDocs List from 3GPP Portal.")
            btn.setStyleSheet("""
                QPushButton {
                    font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; font-weight: bold;
                    border-radius: 12px; padding: 2px 6px;
                    color: #005A9E; background-color: #E1F0FF; border: 1px solid #99C9FF;
                }
                QPushButton:hover {
                    background-color: #CCE4FF; border: 1px solid #005A9E; color: #005A9E;
                }
            """)
            btn.clicked.connect(lambda _, d=row_data, b=btn: self._download_and_open_tdocs(d, b))

        container = QWidget()
        container.setStyleSheet("background-color: transparent;")

        layout = QHBoxLayout(container)
        # Tightly pack the button inside the cell so it doesn't get clipped
        layout.setContentsMargins(2, 0, 2, 0)
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(btn)

        self.table.setIndexWidget(self.table_model.index(row_idx, 1), container)

    def _open_tdocs_window(self, mtg_info: dict, filepath: str):
        self._save_last_meeting(mtg_info)
        mtg_id = mtg_info.get("mtg_id")

        # 1. Bring to front if already open
        if mtg_id in self.tdoc_windows and self.tdoc_windows[mtg_id].isVisible():
            self.tdoc_windows[mtg_id].raise_()
            self.tdoc_windows[mtg_id].activateWindow()
            return

        # 2. Parse and Create Window
        tdocs_data = TDocsParser.parse_tdocs_excel(filepath)
        if not tdocs_data:
            QMessageBox.warning(self, "Parse Error", "Could not read data from the Excel file.")
            return

        # FIXED: Pass the filepath as the third argument!
        window = TDocsWindow(mtg_info, tdocs_data, filepath)
        self.tdoc_windows[mtg_id] = window
        window.show()

    def _on_inline_download_finished(self, success: bool, result: str, mtg_id: str, row_data: dict):
        if mtg_id in self._active_dl_threads:
            del self._active_dl_threads[mtg_id]

        self.refresh_table()  # This resets the button icons from ⏳ to 📊

        if success:
            self._open_tdocs_window(row_data, result)
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to download TDocs List:\n{result}")

    def _open_last_meeting(self):
        try:
            if not self.config_file.exists():
                QMessageBox.information(self, "No History",
                                        "No recent meeting history found. Please open a meeting first.")
                return

            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            last_id = data.get("last_mtg_id")
            last_num = data.get("last_mtg_number")

            if not last_id or not last_num:
                QMessageBox.information(self, "No History",
                                        "No recent meeting history found. Please open a meeting first.")
                return

            # Search the DB using the meeting number as a highly optimized search term
            results = self.db.search_meetings(search_term=last_num)

            # Find the exact match in the returned results
            target_meeting = next((m for m in results if m.get("mtg_id") == last_id), None)

            if not target_meeting:
                QMessageBox.warning(self, "Not Found",
                                    f"Meeting '{last_num}' could not be found in the database.\nIt may have been cleared or the database was updated.")
                return

            filepath = self._get_tdoc_list_path(target_meeting)

            if filepath and filepath.exists():
                # TDocs are already cached, jump straight to the window!
                self._open_tdocs_window(target_meeting, str(filepath))
            else:
                # Needs to be downloaded. We pass a dummy UI button in memory to absorb
                # the "⏳ Fetching" state changes safely, preventing crashes if the row is currently filtered out!
                dummy_btn = QPushButton()
                self._download_and_open_tdocs(target_meeting, dummy_btn)

        except Exception as e:
            QMessageBox.critical(self, "Launch Error", f"Could not open last meeting:\n{e}")

    # ==========================================
    # --- GLOBAL TDOC SEARCH LOGIC ---
    # ==========================================
    def _on_global_tdoc_search(self):
        tdoc_str = self.global_tdoc_input.text().strip()
        if not tdoc_str:
            return

        # --- FIX 1: Prevent UI Double-Firing ---
        if not self.global_tdoc_btn.isEnabled():
            return

        self.global_tdoc_input.setEnabled(False)
        self.global_tdoc_btn.setEnabled(False)
        self.global_tdoc_btn.setText("⏳ Searching...")
        QApplication.processEvents()  # Force UI update

        # 1. Find the parent meeting in the database
        meeting = self.db.find_meeting_by_tdoc(tdoc_str)

        # Restore UI State immediately after DB lookup
        self.global_tdoc_btn.setText("🔍 Find TDoc")
        self.global_tdoc_input.setEnabled(True)
        self.global_tdoc_btn.setEnabled(True)

        if not meeting:
            QMessageBox.warning(
                self,
                "Not Found",
                f"Could not find a meeting containing TDoc '{tdoc_str}'.\n\nEnsure you have fully synced the database and enabled the 'Deep Scrape Docs' option."
            )
            return

        # 2. Extract clean base TDoc and revision targets
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid Format", "Could not parse the provided TDoc number.")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()

        # --- FIX 2: Prevent overlapping background threads for the same file ---
        if target_filename in self._active_dl_threads:
            return

        # 3. Open the TDocs table for this meeting (downloading the Excel file if necessary)
        filepath = self._get_tdoc_list_path(meeting)
        if filepath and filepath.exists():
            self._open_tdocs_window(meeting, str(filepath))
        else:
            dummy_btn = QPushButton()  # Dummy button to absorb loading states
            self._download_and_open_tdocs(meeting, dummy_btn)

        # 4. Trigger the background download and auto-open of the specific TDoc file
        self._download_global_tdoc(meeting, base_tdoc, target_filename, has_rev=bool(match.group(2)))

    def _download_global_tdoc(self, meeting: dict, base_tdoc: str, target_filename: str, has_rev: bool):
        docs_url = meeting.get("docs_folder_url")
        if not docs_url:
            return

        if not docs_url.startswith("http"):
            docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip('/')

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.cache_dir
        folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "")
        meeting_dir = Path(current_cache) / folder_name

        # If it's a revision, we assume it's in the Revisions folder.
        # If the revision folder doesn't exist, the TDocActionThread will safely handle the 404 error.
        dl_url = docs_url
        if has_rev:
            raw_url = meeting.get("url_key", "")
            main_ftp = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            dl_url = main_ftp.rstrip('/') + '/INBOX/Revisions/'

        # Launch the action thread!
        thread = TDocActionThread(base_tdoc, target_filename, dl_url, meeting_dir, open_file=True)
        self._active_dl_threads[target_filename] = thread

        # Connect cleanup and error handling
        thread.finished_action.connect(
            lambda t, s, m, th=thread: self._on_global_tdoc_download_finished(target_filename, s, m, th)
        )
        thread.start()

    def _on_global_tdoc_download_finished(self, tdoc_name: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc_name in self._active_dl_threads:
            del self._active_dl_threads[tdoc_name]

        if not success:
            QMessageBox.warning(self, f"Download Failed: {tdoc_name}", msg)