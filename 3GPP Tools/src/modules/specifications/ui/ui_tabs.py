# --- File: src/modules/specifications/ui/ui_tabs.py ---
import json
import logging
import os
import re
import webbrowser
import zipfile
from pathlib import Path

from PyQt5 import sip
from PyQt5.QtCore import QTimer, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.components import HoverMenuButton
from modules.specifications.ui.dialogs import (
    AdvancedSyncDialog,
    SpecInfoDialog,
    SpecsConfigDialog,
    TableFilterDialog,
    TargetedSyncDialog,
)
from modules.specifications.ui.threads import SpecDownloadThread


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
    update_specific_requested = pyqtSignal(list, bool)
    log_msg = pyqtSignal(str, int)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self._download_threads = []

        self.config_file = db_path.parent / "specs_config.json"

        # Defaults
        self.default_dl_dir = str(Path.home() / "3GPP_SA2_Meeting_Helper" / "specs")
        self.download_dir = self.default_dl_dir
        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}
        self.saved_search = ""
        self.saved_version = ""
        self.saved_downloaded_only = False

        # Load persisted settings
        self._load_settings()

        # Timers
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(350)
        self.search_timer.timeout.connect(self.refresh_table)

        self.save_settings_timer = QTimer()
        self.save_settings_timer.setSingleShot(True)
        self.save_settings_timer.setInterval(800)
        self.save_settings_timer.timeout.connect(self._save_settings)

        self._setup_ui()
        self.refresh_table()

    # ==========================================
    # --- CONFIG / PERSISTENCE ---
    # ==========================================
    def _load_settings(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.download_dir = data.get("download_dir", self.default_dl_dir)
                    self.saved_search = data.get("search_query", "")
                    self.saved_version = data.get("version_query", "")
                    self.saved_downloaded_only = data.get("downloaded_only", False)
                    saved_filters = data.get("table_filters", {})
                    if isinstance(saved_filters, dict):
                        self.table_filters.update(saved_filters)
            except Exception as e:
                print(f"Error loading specs_config.json: {e}")
        else:
            self.download_dir = self.default_dl_dir

    def _save_settings(self):
        try:
            data = {
                "download_dir": self.download_dir,
                "search_query": self.spec_search_input.text().strip(),
                "version_query": self.version_search_input.text().strip(),
                "downloaded_only": self.downloaded_only_cb.isChecked(),
                "table_filters": self.table_filters,
            }
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving specs_config.json: {e}")

    # ==========================================
    # --- UI SETUP ---
    # ==========================================
    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # --- ROW 1: Network Sync & Configuration Toolbar ---
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(8)

        toolbar_layout.addWidget(QLabel("<b>🌐 Network Sync:</b>"))

        self.sync_menu_btn = QPushButton("🌐 Sync & Fetch ▾")
        self.sync_menu_btn.setCursor(Qt.PointingHandCursor)
        self.sync_menu_btn.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #0066CC;
                color: #FFFFFF;
                border: 1px solid #0055AA;
                border-radius: 4px;
                padding: 5px 12px;
            }
            QPushButton:hover {
                background-color: #0052A3;
            }
        """)

        sync_menu = QMenu(self)
        sync_menu.setStyleSheet("""
            QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
            QMenu::item { padding: 6px 20px 6px 15px; color: #333333; }
            QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
        """)

        quick_action = sync_menu.addAction("🎯 Quick Fetch (By Spec or Series)...")
        quick_action.setToolTip("Instantly download specific specs or series (e.g., '23.501' or '23') directly.")
        quick_action.triggered.connect(self._open_targeted_sync)

        scoped_action = sync_menu.addAction("⚙️ Scoped Sync (By WG / Series)...")
        scoped_action.setToolTip("Sync only specific portions of the database based on Series, Tech, or Group.")
        scoped_action.triggered.connect(self._open_advanced_sync)

        full_action = sync_menu.addAction("🔄 Full Database Sync (All Series)...")
        full_action.setToolTip("Synchronize the entire 3GPP specification database from the FTP server.")
        full_action.triggered.connect(lambda: self.update_db_requested.emit(self.force_meta_action.isChecked()))

        sync_menu.addSeparator()

        self.force_meta_action = QAction("Force Re-download Metadata", self)
        self.force_meta_action.setCheckable(True)
        self.force_meta_action.setChecked(False)
        self.force_meta_action.setToolTip("Force the scraper to re-download HTML metadata even if it already exists.")
        sync_menu.addAction(self.force_meta_action)

        sync_menu.addSeparator()
        config_action = sync_menu.addAction("⚙️ Download Settings...")
        config_action.triggered.connect(self._open_settings_dialog)

        open_folder_action = sync_menu.addAction("📂 Open Base Download Directory")
        open_folder_action.triggered.connect(self._open_download_dir)

        self.sync_menu_btn.setMenu(sync_menu)
        toolbar_layout.addWidget(self.sync_menu_btn)

        # Toolbar quick-action buttons
        self.settings_btn = QPushButton("⚙️ Settings")
        self.settings_btn.setToolTip("Configure the local specifications download folder and paths.")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.clicked.connect(self._open_settings_dialog)
        toolbar_layout.addWidget(self.settings_btn)

        self.open_dir_btn = QPushButton("↗️ Open Folder")
        self.open_dir_btn.setToolTip("Open the local specifications storage directory in Explorer.")
        self.open_dir_btn.setCursor(Qt.PointingHandCursor)
        self.open_dir_btn.clicked.connect(self._open_download_dir)
        toolbar_layout.addWidget(self.open_dir_btn)

        self.bg_sync_label = QLabel("⏳ Fetching deep metadata in background...")
        self.bg_sync_label.setStyleSheet("color: #E65100; font-weight: bold; font-style: italic;")
        self.bg_sync_label.setVisible(False)
        toolbar_layout.addWidget(self.bg_sync_label)

        toolbar_layout.addStretch()
        main_layout.addLayout(toolbar_layout)

        # --- ROW 2: Local Search & Main Filters ---
        search_layout = QHBoxLayout()
        search_layout.setSpacing(6)

        search_layout.addWidget(QLabel("<b>🔍 Search:</b>"))

        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("Spec Number or Title...")
        self.spec_search_input.setText(self.saved_search)
        self.spec_search_input.setClearButtonEnabled(True)
        self.spec_search_input.setToolTip("Filter instantly by specification number (e.g., '23.501') or title keywords.")
        self.spec_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.spec_search_input, stretch=2)

        search_layout.addWidget(QLabel("Ver:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15.")
        self.version_search_input.setText(self.saved_version)
        self.version_search_input.setClearButtonEnabled(True)
        self.version_search_input.setToolTip("Filter the table by release version (e.g., '15' or '16.2').")
        self.version_search_input.setFixedWidth(75)
        self.version_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.version_search_input)

        self.downloaded_only_cb = QCheckBox("💾 Downloaded Only")
        self.downloaded_only_cb.setToolTip("Show only specifications that currently have local cached files.")
        self.downloaded_only_cb.setChecked(self.saved_downloaded_only)
        self.downloaded_only_cb.toggled.connect(self._on_search_changed)
        search_layout.addWidget(self.downloaded_only_cb)

        self.filter_btn = QPushButton("⚙️ Table Filters")
        self.filter_btn.setToolTip("Apply advanced filters (Series, Technology, Group) to the local table.")
        self.filter_btn.setCursor(Qt.PointingHandCursor)
        self.filter_btn.clicked.connect(self._open_table_filters)
        search_layout.addWidget(self.filter_btn)

        self.clear_filter_btn = QPushButton("❌")
        self.clear_filter_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filter_btn.setToolTip("Clear all active table filters.")
        self.clear_filter_btn.setFixedWidth(26)
        self.clear_filter_btn.setVisible(False)
        self.clear_filter_btn.clicked.connect(self._clear_table_filters)
        search_layout.addWidget(self.clear_filter_btn)

        main_layout.addLayout(search_layout)

        # --- ROW 3: Quick Series Preset Chips & Status Banner ---
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(4)

        chips_label = QLabel("Presets:")
        chips_label.setStyleSheet("color: #718096; font-size: 11px; font-weight: bold;")
        chips_layout.addWidget(chips_label)

        self.chip_buttons = {}
        preset_chips = [
            ("All", ""),
            ("23 (SA2)", "23"),
            ("38 (RAN)", "38"),
            ("24 (CT)", "24"),
            ("36 (LTE)", "36"),
        ]

        for label, series_val in preset_chips:
            chip = QPushButton(label)
            chip.setCursor(Qt.PointingHandCursor)
            chip.clicked.connect(lambda _, s=series_val: self._on_chip_clicked(s))
            self.chip_buttons[series_val] = chip
            chips_layout.addWidget(chip)

        chips_layout.addStretch()

        self.count_badge = QLabel("0 specifications")
        self.count_badge.setStyleSheet("""
            QLabel {
                padding: 2px 10px;
                font-size: 11px;
                font-weight: bold;
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #E2E8F0;
                border-radius: 10px;
            }
        """)
        chips_layout.addWidget(self.count_badge)

        main_layout.addLayout(chips_layout)

        # --- ROW 4: Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Specification", "Title", "Version / Documents"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.verticalHeader().setStyleSheet("QHeaderView::section { color: #94A3B8; font-size: 11px; }")

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

        self._update_filter_ui()
        self._update_chip_styles()

    # ==========================================
    # --- EVENT HANDLERS & FILTERS ---
    # ==========================================
    def _open_settings_dialog(self):
        dialog = SpecsConfigDialog(self.download_dir, self)
        if dialog.exec_():
            new_path = dialog.get_download_path()
            if new_path and new_path != self.download_dir:
                self.download_dir = new_path
                self._save_settings()
                self.refresh_table()

    def _on_search_changed(self):
        self.search_timer.start()
        self.save_settings_timer.start()

    def _on_chip_clicked(self, series_val: str):
        if self.table_filters.get("series") == series_val and series_val != "":
            self.table_filters["series"] = ""
        else:
            self.table_filters["series"] = series_val

        self._update_chip_styles()
        self._update_filter_ui()
        self.save_settings_timer.start()
        self.refresh_table()

    def _update_chip_styles(self):
        current_series = self.table_filters.get("series", "")
        for series_val, chip in self.chip_buttons.items():
            is_active = (current_series == series_val) or (not current_series and series_val == "")
            if is_active:
                chip.setStyleSheet("""
                    QPushButton {
                        padding: 2px 8px;
                        font-size: 11px;
                        font-weight: bold;
                        background-color: #0066CC;
                        color: white;
                        border: 1px solid #0055AA;
                        border-radius: 10px;
                    }
                """)
            else:
                chip.setStyleSheet("""
                    QPushButton {
                        padding: 2px 8px;
                        font-size: 11px;
                        background-color: #F0F4F8;
                        color: #2D3748;
                        border: 1px solid #D2E3FC;
                        border-radius: 10px;
                    }
                    QPushButton:hover {
                        background-color: #E2E8F0;
                    }
                """)

    def set_bg_sync_active(self, is_active: bool):
        self.bg_sync_label.setVisible(is_active)
        if not is_active:
            self.refresh_table()

    def _open_download_dir(self):
        target_dir = Path(self.download_dir)
        if not target_dir.exists():
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Directory Error", f"Could not create directory:\n{e}")
                return
        try:
            os.startfile(str(target_dir))
        except Exception as e:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _open_advanced_sync(self):
        dialog = AdvancedSyncDialog(self.db, self)
        if dialog.exec_():
            target_specs = dialog.matching_specs
            if target_specs:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(target_specs, force_meta)

    def _open_targeted_sync(self):
        dialog = TargetedSyncDialog(self)
        if dialog.exec_():
            targets = dialog.get_targets()
            if targets:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(targets, force_meta)

    def _open_table_filters(self):
        dialog = TableFilterDialog(self.db, self.table_filters, self)
        if dialog.exec_():
            self.table_filters = dialog.get_filters()
            self._update_filter_ui()
            self._update_chip_styles()
            self.save_settings_timer.start()
            self.refresh_table()

    def _clear_table_filters(self):
        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}
        self._update_filter_ui()
        self._update_chip_styles()
        self.save_settings_timer.start()
        self.refresh_table()

    def _update_filter_ui(self):
        active_count = 0
        if self.table_filters.get("series"):
            active_count += 1
        if self.table_filters.get("tech"):
            active_count += 1
        if self.table_filters.get("group"):
            active_count += 1
        if self.table_filters.get("spec_type", "Any") != "Any":
            active_count += 1

        if active_count > 0:
            self.filter_btn.setText(f"⚙️ Filters ({active_count})")
            self.filter_btn.setStyleSheet(
                "background-color: #E1F0FF; color: #0078D7; font-weight: bold; border: 1px solid #0078D7;"
            )
            self.clear_filter_btn.setVisible(True)
        else:
            self.filter_btn.setText("⚙️ Table Filters")
            self.filter_btn.setStyleSheet("")
            self.clear_filter_btn.setVisible(False)

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
                widget = self.table.cellWidget(index.row(), 0)
                num_label = widget.findChild(QLabel, "specNumberLabel")
                if num_label:
                    target_specs.append(num_label.text().strip())
            force_meta = self.force_meta_action.isChecked()
            self.update_specific_requested.emit(target_specs, force_meta)

    def _show_spec_info(self, spec_num: str):
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
            if details:
                details["number"] = spec_num

        dialog = SpecInfoDialog(details or {}, self)
        dialog.exec_()

    def _open_spec_folder(self, spec_num: str):
        target_dir = Path(self.download_dir) / spec_num
        if target_dir.exists() and any(target_dir.iterdir()):
            try:
                os.startfile(str(target_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")
        else:
            QMessageBox.information(
                self,
                "Folder Not Found",
                f"No downloaded files found for {spec_num}.\nDownload a document first to create the local folder.",
            )

    def _open_web_report(self, spec_num: str):
        clean_number = spec_num.replace(".", "")
        url = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        webbrowser.open(url)

    # ==========================================
    # --- TABLE REFRESH & RENDERING ---
    # ==========================================
    def refresh_table(self):
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()
            downloaded_only = self.downloaded_only_cb.isChecked()
            base_dl_dir = Path(self.download_dir)

            is_filtered = any(
                [
                    self.table_filters["series"],
                    self.table_filters["tech"],
                    self.table_filters["group"],
                    self.table_filters["spec_type"] != "Any",
                ]
            )

            # Query database
            specs = self.db.search_files(
                spec_number=spec_query if spec_query else None,
                release_version=version_query if version_query else None,
                **self.table_filters,
            )

            self.table.setRowCount(0)

            grouped_specs = {}
            for row in specs:
                series, spec_num, title, spec_type, filename, version, url, upload_date = row
                if filename:
                    part_match = re.search(r"\d{4,5}-(\d{1,2})(?:[-_.]|$)", filename)
                    if part_match and "-" not in spec_num:
                        spec_num = f"{spec_num}-{part_match.group(1)}"

                if spec_num not in grouped_specs:
                    grouped_specs[spec_num] = {
                        "title": title,
                        "type": spec_type if spec_type else "TS",
                        "versions": [],
                    }
                grouped_specs[spec_num]["versions"].append((version, url, filename, upload_date))

            # Filter for Downloaded Only if enabled
            if downloaded_only and base_dl_dir.exists():
                filtered_grouped = {}
                for s_num, s_data in grouped_specs.items():
                    s_dir = base_dl_dir / s_num
                    if s_dir.exists() and any(s_dir.iterdir()):
                        filtered_grouped[s_num] = s_data
                grouped_specs = filtered_grouped

            total_found = len(grouped_specs)
            rendered_specs = list(grouped_specs.items())[:100]

            # Dynamic pill badge status
            if not spec_query and not version_query and not is_filtered and not downloaded_only:
                self.count_badge.setText(f"Top {len(rendered_specs)} specifications")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #EBF8FF; color: #2B6CB0; border: 1px solid #BEE3F8; border-radius: 10px;
                """)
            elif total_found > 100:
                self.count_badge.setText(f"Showing 100 of {total_found} specs")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #FFF3E0; color: #D83B01; border: 1px solid #FFB74D; border-radius: 10px;
                """)
            else:
                self.count_badge.setText(f"{total_found} specifications")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #F1F5F9; color: #475569; border: 1px solid #E2E8F0; border-radius: 10px;
                """)

            for row_idx, (spec_num, data) in enumerate(rendered_specs):
                self.table.insertRow(row_idx)
                spec_target_dir = base_dl_dir / spec_num
                dir_has_files = spec_target_dir.exists() and any(spec_target_dir.iterdir())

                # --- COLUMN 0: Action Button + Badged Spec Number ---
                spec_widget = QWidget()
                spec_layout = QHBoxLayout(spec_widget)
                spec_layout.setContentsMargins(6, 0, 6, 0)
                spec_layout.setSpacing(6)

                action_btn = HoverMenuButton("⋮")
                action_btn.setFixedSize(22, 22)
                action_btn.setToolTip("Specification Actions")
                action_btn.setCursor(Qt.PointingHandCursor)
                action_btn.setStyleSheet("""
                    QPushButton { border: none; background: transparent; color: #718096; font-size: 18px; font-weight: bold; padding-bottom: 3px; }
                    QPushButton:hover { color: #0078D7; }
                    QPushButton::menu-indicator { image: none; width: 0px; }
                """)

                menu = QMenu(self)
                menu.setStyleSheet("""
                    QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
                    QMenu::item { padding: 5px 20px 5px 15px; color: #333333; }
                    QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                    QMenu::item:disabled { color: #AAAAAA; }
                """)

                info_action = menu.addAction("ℹ️  View Details")
                info_action.triggered.connect(lambda _, s=spec_num: self._show_spec_info(s))

                web_action = menu.addAction("🌐  Open 3GPP Web Report")
                web_action.triggered.connect(lambda _, s=spec_num: self._open_web_report(s))

                menu.addSeparator()

                folder_action = menu.addAction("📂  Open Local Folder")
                folder_action.triggered.connect(lambda _, s=spec_num: self._open_spec_folder(s))

                def _update_menu_state(act=folder_action, path=spec_target_dir):
                    if path.exists() and any(path.iterdir()):
                        act.setText("📂  Open Local Folder")
                        act.setEnabled(True)
                    else:
                        act.setText("📁  Folder Not Created")
                        act.setEnabled(False)

                menu.aboutToShow.connect(_update_menu_state)
                _update_menu_state()
                action_btn.setMenu(menu)

                # Spec Type Badge (TS vs TR)
                spec_type = (data["type"] or "TS").upper()
                type_badge = QLabel(spec_type)
                if spec_type == "TR":
                    type_badge.setStyleSheet("""
                        background-color: #F3E8FF; color: #6B21A8; border: 1px solid #E9D5FF;
                        border-radius: 3px; padding: 1px 5px; font-size: 10px; font-weight: bold;
                    """)
                else:
                    type_badge.setStyleSheet("""
                        background-color: #EBF8FF; color: #2B6CB0; border: 1px solid #BEE3F8;
                        border-radius: 3px; padding: 1px 5px; font-size: 10px; font-weight: bold;
                    """)

                spec_label = QLabel(spec_num)
                spec_label.setObjectName("specNumberLabel")
                spec_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #1A202C;")

                spec_layout.addWidget(action_btn)
                spec_layout.addWidget(type_badge)
                spec_layout.addWidget(spec_label)
                spec_layout.addStretch()
                self.table.setCellWidget(row_idx, 0, spec_widget)

                # --- COLUMN 1: Title ---
                self.table.setItem(row_idx, 1, QTableWidgetItem(data["title"] if data["title"] else "Unknown Title"))

                # --- COLUMN 2: Documents Action Bar ---
                version_combo = QComboBox()
                version_combo.setFixedWidth(195)
                version_combo.setFixedHeight(26)

                def parse_ver(v_str):
                    return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split(".")]

                sorted_versions = sorted(data["versions"], key=lambda x: parse_ver(x[0]), reverse=True)

                for ver, url, fname, u_date in sorted_versions:
                    zip_path = spec_target_dir / fname
                    is_dl = zip_path.exists()
                    status = "✅ " if is_dl else ""
                    date_label = f" ({u_date})" if u_date else ""
                    display_text = f"{status}v{ver}{date_label}"

                    version_combo.addItem(
                        display_text,
                        userData={
                            "url": url,
                            "fname": fname,
                            "spec_num": spec_num,
                            "is_downloaded": is_dl,
                            "upload_date": u_date,
                        },
                    )

                doc_action_btn = QPushButton()
                doc_action_btn.setFixedWidth(145)
                doc_action_btn.setFixedHeight(26)
                doc_action_btn.setCursor(Qt.PointingHandCursor)

                doc_menu = QMenu(self)
                doc_menu.setStyleSheet("""
                    QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
                    QMenu::item { padding: 6px 22px 6px 14px; color: #333333; font-size: 12px; }
                    QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                    QMenu::item:disabled { color: #AAAAAA; }
                    QMenu::separator { height: 1px; background-color: #E2E8F0; margin: 4px 0; }
                """)
                doc_action_btn.setMenu(doc_menu)

                def _update_btn_state(index_ignore=0, c=version_combo, btn=doc_action_btn, menu=doc_menu):
                    c_data = c.currentData()
                    if not c_data:
                        return

                    current_dir = Path(self.download_dir) / c_data["spec_num"]
                    stem = Path(c_data["fname"]).stem
                    zip_exists = (current_dir / c_data["fname"]).exists()

                    word_exists = any(current_dir.glob(f"{stem}*.doc*"))
                    pdf_exists = any(current_dir.glob(f"{stem}*.pdf"))
                    html_exists = any(current_dir.glob(f"{stem}*.html"))
                    txt_exists = any(current_dir.glob(f"{stem}*.txt"))
                    dir_ready = current_dir.exists() and any(current_dir.iterdir())

                    if word_exists:
                        btn.setText("📝 Open Word ▾")
                        btn.setStyleSheet("""
                            QPushButton {
                                font-size: 11px;
                                font-weight: bold;
                                background-color: #E8F5E9;
                                color: #2E7D32;
                                border: 1px solid #2E7D32;
                                border-radius: 4px;
                            }
                            QPushButton:hover { background-color: #C8E6C9; }
                        """)
                    elif zip_exists:
                        btn.setText("⚙️ Extract Word ▾")
                        btn.setStyleSheet("""
                            QPushButton {
                                font-size: 11px;
                                font-weight: bold;
                                background-color: #FFF3E0;
                                color: #E65100;
                                border: 1px solid #FFB74D;
                                border-radius: 4px;
                            }
                            QPushButton:hover { background-color: #FFE0B2; }
                        """)
                    else:
                        btn.setText("⬇️ Get Word ▾")
                        btn.setStyleSheet("""
                            QPushButton {
                                font-size: 11px;
                                font-weight: bold;
                                background-color: #EBF3FB;
                                color: #0066CC;
                                border: 1px solid #B0D0F0;
                                border-radius: 4px;
                            }
                            QPushButton:hover { background-color: #D6E8FA; border-color: #0066CC; }
                        """)

                    menu.clear()

                    word_label = "📝 Open Word Document" if word_exists else ("⚙️ Extract Word Document" if zip_exists else "⬇️ Download & Open Word")
                    act_word = menu.addAction(f"{word_label} {'✅' if word_exists else ''}".strip())
                    act_word.triggered.connect(lambda: self._handle_document_action(c, "word", btn))

                    pdf_label = "📕 Open PDF" if pdf_exists else ("⚙️ Convert to PDF" if (word_exists or zip_exists) else "⬇️ Get & Convert to PDF")
                    act_pdf = menu.addAction(f"{pdf_label} {'✅' if pdf_exists else ''}".strip())
                    act_pdf.triggered.connect(lambda: self._handle_document_action(c, "pdf", btn))

                    html_label = "🌐 Open HTML" if html_exists else ("⚙️ Convert to HTML" if (word_exists or zip_exists) else "⬇️ Get & Convert to HTML")
                    act_html = menu.addAction(f"{html_label} {'✅' if html_exists else ''}".strip())
                    act_html.triggered.connect(lambda: self._handle_document_action(c, "html", btn))

                    txt_label = "📄 Open TXT" if txt_exists else ("⚙️ Convert to TXT" if (word_exists or zip_exists) else "⬇️ Get & Convert to TXT")
                    act_txt = menu.addAction(f"{txt_label} {'✅' if txt_exists else ''}".strip())
                    act_txt.triggered.connect(lambda: self._handle_document_action(c, "txt", btn))

                    menu.addSeparator()

                    zip_label = "📦 Show ZIP Archive ✅" if zip_exists else "📥 Download ZIP Archive"
                    act_zip = menu.addAction(zip_label)
                    act_zip.triggered.connect(lambda: self._handle_zip_action(c, btn))

                    if dir_ready:
                        act_dir = menu.addAction("📂 Open Local Folder")
                        act_dir.setEnabled(True)
                        act_dir.triggered.connect(lambda _, s=c_data["spec_num"]: self._open_spec_folder(s))
                    else:
                        act_dir = menu.addAction("📁 Folder Not Created")
                        act_dir.setEnabled(False)

                version_combo.currentIndexChanged.connect(_update_btn_state)
                _update_btn_state()

                cell_widget = QWidget()
                layout = QHBoxLayout(cell_widget)
                layout.setContentsMargins(4, 0, 4, 0)
                layout.setSpacing(6)

                layout.addWidget(version_combo)
                layout.addWidget(doc_action_btn)
                layout.addStretch()

                self.table.setCellWidget(row_idx, 2, cell_widget)

        except Exception as e:
            print(f"Error during refresh_table: {e}")

    # ==========================================
    # --- DOCUMENT FETCH & EXTRACTION ---
    # ==========================================
    def _handle_document_action(self, combo: QComboBox, doc_type: str, btn: QPushButton):
        c_data = combo.currentData()
        if not c_data:
            return

        spec_dl_dir = Path(self.download_dir) / c_data["spec_num"]
        zip_path = spec_dl_dir / c_data["fname"]
        stem = Path(c_data["fname"]).stem

        def _process_and_open():
            extracted_docs = []

            if zip_path.exists():
                try:
                    with zipfile.ZipFile(zip_path, "r") as z:
                        for member in z.namelist():
                            if "__MACOSX" in member or member.startswith("._"):
                                continue

                            if member.lower().endswith((".doc", ".docx")):
                                target_file = spec_dl_dir / Path(member).name
                                if not target_file.exists():
                                    target_file.write_bytes(z.read(member))

                                if target_file not in extracted_docs:
                                    extracted_docs.append(target_file)
                except Exception as e:
                    QMessageBox.warning(self, "Extraction Error", f"Failed to extract archive:\n{e}")
                    return

            if not extracted_docs:
                extracted_docs = list(spec_dl_dir.glob(f"{stem}*.doc*"))

            if not extracted_docs:
                QMessageBox.warning(self, "Not Found", "No Word documents found on disk or inside the zip archive.")
                return

            try:
                from modules.word_tools.core.word_converter import WordConverterThread
            except ImportError as e:
                QMessageBox.warning(self, "Import Error", f"Could not import word_converter:\n{e}")
                return

            for doc_path in extracted_docs:
                try:
                    if doc_type == "word":
                        os.startfile(str(doc_path))

                    elif doc_type in ("pdf", "html", "txt"):
                        target_ext = f".{doc_type}"
                        target_path = doc_path.with_suffix(target_ext)

                        if not target_path.exists():
                            orig_text = btn.text() if not sip.isdeleted(btn) else "Converting..."
                            if not sip.isdeleted(btn):
                                btn.setText("⏳ Converting...")
                                btn.setEnabled(False)

                            conv_thread = WordConverterThread(str(doc_path), doc_type)
                            conv_thread.ui_log_msg.connect(self._handle_converter_log)

                            def on_success(p, c=combo, b=btn, txt=orig_text):
                                try:
                                    os.startfile(p)
                                except Exception as e:
                                    print(f"Error opening converted file: {e}")

                                try:
                                    if not sip.isdeleted(c):
                                        c.currentIndexChanged.emit(c.currentIndex())
                                    if not sip.isdeleted(b):
                                        b.setText(txt)
                                        b.setEnabled(True)
                                except RuntimeError:
                                    pass

                            conv_thread.finished_path.connect(on_success)

                            def cleanup(t=conv_thread, b=btn, txt=orig_text):
                                if t in self._download_threads:
                                    self._download_threads.remove(t)
                                try:
                                    if not sip.isdeleted(b) and not b.isEnabled():
                                        b.setText(txt)
                                        b.setEnabled(True)
                                except RuntimeError:
                                    pass

                            conv_thread.finished.connect(cleanup)
                            self._download_threads.append(conv_thread)
                            conv_thread.start()

                        else:
                            os.startfile(str(target_path))

                except Exception as e:
                    QMessageBox.warning(self, "Open Error", f"Could not open/convert {doc_type.upper()}:\n{e}")

        word_exists = any(spec_dl_dir.glob(f"{stem}*.doc*"))
        target_exists = False

        if doc_type == "word":
            target_exists = word_exists
        elif doc_type == "pdf":
            target_exists = any(spec_dl_dir.glob(f"{stem}*.pdf"))
        elif doc_type == "html":
            target_exists = any(spec_dl_dir.glob(f"{stem}*.html"))
        elif doc_type == "txt":
            target_exists = any(spec_dl_dir.glob(f"{stem}*.txt"))

        if target_exists or word_exists or zip_path.exists():
            _process_and_open()
        else:
            spec_dl_dir.mkdir(parents=True, exist_ok=True)

            idx = combo.currentIndex()
            orig_text = combo.itemText(idx)
            combo.setItemText(idx, "⏳ Downloading...")
            combo.setEnabled(False)

            if not sip.isdeleted(btn):
                btn.setText("⏳ Downloading...")
                btn.setEnabled(False)

            thread = SpecDownloadThread(c_data["url"], zip_path)
            thread.ui_log_msg.connect(self._handle_converter_log)

            def _on_success(zp):
                clean_text = orig_text.replace("✅ ", "").replace("⚙️ ", "").replace("⬇️ ", "").strip()
                try:
                    if not sip.isdeleted(combo):
                        combo.setItemText(idx, f"✅ {clean_text}")
                        c_data_inner = combo.itemData(idx)
                        if c_data_inner:
                            c_data_inner["is_downloaded"] = True
                            combo.setItemData(idx, c_data_inner)
                        combo.setEnabled(True)
                    if not sip.isdeleted(btn):
                        btn.setEnabled(True)
                except RuntimeError:
                    pass

                _process_and_open()

                try:
                    if not sip.isdeleted(combo):
                        combo.currentIndexChanged.emit(idx)
                except RuntimeError:
                    pass

            def _on_err(err):
                try:
                    if not sip.isdeleted(combo):
                        combo.setItemText(idx, "❌ Error")
                        combo.setEnabled(True)
                    if not sip.isdeleted(btn):
                        btn.setEnabled(True)
                        combo.currentIndexChanged.emit(idx)
                except RuntimeError:
                    pass
                QMessageBox.critical(self, "Download Failed", f"Network error:\n{err}")

            thread.finished_success.connect(_on_success)
            thread.error.connect(_on_err)

            self._download_threads.append(thread)
            thread.finished.connect(
                lambda t=thread: self._download_threads.remove(t) if t in self._download_threads else None
            )
            thread.start()

    def _handle_zip_action(self, combo: QComboBox, btn: QPushButton):
        c_data = combo.currentData()
        if not c_data:
            return

        spec_dl_dir = Path(self.download_dir) / c_data["spec_num"]
        zip_path = spec_dl_dir / c_data["fname"]

        if zip_path.exists():
            try:
                os.startfile(str(spec_dl_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")
            return

        spec_dl_dir.mkdir(parents=True, exist_ok=True)

        idx = combo.currentIndex()
        orig_text = combo.itemText(idx)
        combo.setItemText(idx, "⏳ Downloading...")
        combo.setEnabled(False)

        if not sip.isdeleted(btn):
            btn.setText("⏳ Downloading...")
            btn.setEnabled(False)

        thread = SpecDownloadThread(c_data["url"], zip_path)
        thread.ui_log_msg.connect(self._handle_converter_log)

        def _on_success(zp):
            clean_text = orig_text.replace("✅ ", "").replace("⚙️ ", "").replace("⬇️ ", "").strip()
            try:
                if not sip.isdeleted(combo):
                    combo.setItemText(idx, f"✅ {clean_text}")
                    c_data_inner = combo.itemData(idx)
                    if c_data_inner:
                        c_data_inner["is_downloaded"] = True
                        combo.setItemData(idx, c_data_inner)
                    combo.setEnabled(True)
                if not sip.isdeleted(btn):
                    btn.setEnabled(True)
                    combo.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass

        def _on_err(err):
            try:
                if not sip.isdeleted(combo):
                    combo.setItemText(idx, "❌ Error")
                    combo.setEnabled(True)
                if not sip.isdeleted(btn):
                    btn.setEnabled(True)
                    combo.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass
            QMessageBox.critical(self, "Download Failed", f"Network error:\n{err}")

        thread.finished_success.connect(_on_success)
        thread.error.connect(_on_err)

        self._download_threads.append(thread)
        thread.finished.connect(
            lambda t=thread: self._download_threads.remove(t) if t in self._download_threads else None
        )
        thread.start()

    def _handle_converter_log(self, msg: str, level: int):
        self.log_msg.emit(msg, level)