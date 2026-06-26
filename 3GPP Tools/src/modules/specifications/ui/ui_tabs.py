# --- File: modules/specs_db/ui_tabs.py ---
import json
import os
import re
import zipfile
import webbrowser
from pathlib import Path

from PyQt5.QtCore import pyqtSignal, Qt, QTimer
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QComboBox, QMenu, QAbstractItemView,
                             QFileDialog, QMessageBox, QFrame)

from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.components import HoverMenuButton
from modules.specifications.ui.threads import SpecDownloadThread
from modules.specifications.ui.dialogs import SpecInfoDialog, AdvancedSyncDialog, TableFilterDialog, TargetedSyncDialog


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
    update_specific_requested = pyqtSignal(list, bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self._download_threads = []

        self.config_file = db_path.parent / "specs_config.json"
        self.default_dl_dir = self._load_settings()

        self.table_filters = {'series': '', 'tech': '', 'group': '', 'spec_type': 'Any'}

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(400)
        self.search_timer.timeout.connect(self.refresh_table)

        self._setup_ui()
        self.refresh_table()

    def _load_settings(self) -> str:
        fallback = str(Path.home() / "3GPP_SA2_Meeting_Helper" / "specs")
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
            data = {'download_dir': current_dir}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 10, 5, 5)

        # --- ROW 1: Download Path ---
        dir_layout = QHBoxLayout()
        self.dl_dir_input = QLineEdit(self.default_dl_dir)
        self.dl_dir_input.editingFinished.connect(self._save_settings)
        browse_btn = QPushButton("📂 Browse...")
        browse_btn.clicked.connect(self._browse_dir)
        open_dir_btn = QPushButton("↗️ Open")
        open_dir_btn.setToolTip("Open the base download folder in Windows Explorer")
        open_dir_btn.setCursor(Qt.PointingHandCursor)
        open_dir_btn.clicked.connect(self._open_download_dir)

        dir_layout.addWidget(QLabel("💾 Download Path:"))
        dir_layout.addWidget(self.dl_dir_input)
        dir_layout.addWidget(browse_btn)
        dir_layout.addWidget(open_dir_btn)
        main_layout.addLayout(dir_layout)

        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        line1.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line1)

        # --- ROW 2: Network Actions ---
        sync_layout = QHBoxLayout()
        sync_layout.addWidget(QLabel("<b>🌐 Network Sync:</b>"))

        self.full_sync_btn = QPushButton("🔄 Full Sync")
        self.full_sync_btn.clicked.connect(lambda: self.update_db_requested.emit(self.force_meta_checkbox.isChecked()))
        sync_layout.addWidget(self.full_sync_btn)

        self.target_sync_btn = QPushButton("🎯 Quick Fetch")
        self.target_sync_btn.clicked.connect(self._open_targeted_sync)
        sync_layout.addWidget(self.target_sync_btn)

        self.adv_sync_btn = QPushButton("⚙️ Filtered Sync")
        self.adv_sync_btn.clicked.connect(self._open_advanced_sync)
        sync_layout.addWidget(self.adv_sync_btn)

        self.force_meta_checkbox = QCheckBox("Force Metadata")
        sync_layout.addWidget(self.force_meta_checkbox)

        self.bg_sync_label = QLabel("⏳ Fetching deep metadata in background...")
        self.bg_sync_label.setStyleSheet("color: #E65100; font-weight: bold; font-style: italic;")
        self.bg_sync_label.setVisible(False)
        sync_layout.addWidget(self.bg_sync_label)

        sync_layout.addStretch()

        main_layout.addLayout(sync_layout)

        # --- ROW 3: Local Table Search ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("<b>🔍 Local Search:</b>"))

        self.spec_search_input = QLineEdit()
        self.spec_search_input.setPlaceholderText("Spec Number or Title...")
        self.spec_search_input.textChanged.connect(lambda text: self.search_timer.start())
        search_layout.addWidget(self.spec_search_input)

        search_layout.addWidget(QLabel("Ver:"))
        self.version_search_input = QLineEdit()
        self.version_search_input.setPlaceholderText("e.g. 15.")
        self.version_search_input.setFixedWidth(60)
        self.version_search_input.textChanged.connect(lambda text: self.search_timer.start())
        search_layout.addWidget(self.version_search_input)

        self.filter_btn = QPushButton("⚙️ Table Filters")
        self.filter_btn.setCursor(Qt.PointingHandCursor)
        self.filter_btn.clicked.connect(self._open_table_filters)

        self.clear_filter_btn = QPushButton("❌")
        self.clear_filter_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filter_btn.setToolTip("Clear Active Filters")
        self.clear_filter_btn.setFixedWidth(24)
        self.clear_filter_btn.setVisible(False)
        self.clear_filter_btn.clicked.connect(self._clear_table_filters)

        search_layout.addWidget(self.filter_btn)
        search_layout.addWidget(self.clear_filter_btn)
        main_layout.addLayout(search_layout)

        # --- Results Header ---
        self.count_label = QLabel("Showing 0 specifications")
        self.count_label.setStyleSheet("font-weight: bold; color: #555555; margin-top: 5px;")

        header_layout = QHBoxLayout()
        header_layout.addStretch()
        header_layout.addWidget(self.count_label)
        main_layout.addLayout(header_layout)

        # --- Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Specification", "Title", "Version / Documents"])
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

    def set_bg_sync_active(self, is_active: bool):
        self.bg_sync_label.setVisible(is_active)
        if not is_active:
            self.refresh_table()

    def _browse_dir(self):
        new_dir = QFileDialog.getExistingDirectory(self, "Select Download Directory", self.dl_dir_input.text())
        if new_dir:
            self.dl_dir_input.setText(new_dir)
            self._save_settings()
            self.refresh_table()

    def _open_download_dir(self):
        target_dir = Path(self.dl_dir_input.text().strip())
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
                force_meta = self.force_meta_checkbox.isChecked()
                self.update_specific_requested.emit(target_specs, force_meta)

    def _open_targeted_sync(self):
        dialog = TargetedSyncDialog(self)
        if dialog.exec_():
            targets = dialog.get_targets()
            if targets:
                force_meta = self.force_meta_checkbox.isChecked()
                # Emits the exact same signal used by the right-click menu and Filtered Sync!
                self.update_specific_requested.emit(targets, force_meta)

    def _show_context_menu(self, position):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows: return

        menu = QMenu()
        update_action = menu.addAction(f"🔄 Update selected ({len(selected_rows)} specifications)")
        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if action == update_action:
            target_specs = []
            for index in selected_rows:
                widget = self.table.cellWidget(index.row(), 0)
                display_text = widget.findChild(QLabel).text()
                spec_num = display_text.split(" ")[-1]
                target_specs.append(spec_num)
            force_meta = self.force_meta_checkbox.isChecked()
            self.update_specific_requested.emit(target_specs, force_meta)

    def _show_spec_info(self, spec_num: str):
        details = self.db.get_spec_details(spec_num)

        # --- FIX: Fallback if the database strictly stored the base number ---
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
            if details:
                # Graft the part number back onto the details dictionary for the UI display
                details['number'] = spec_num

        dialog = SpecInfoDialog(details or {}, self)
        dialog.exec_()

    def _open_spec_folder(self, spec_num: str):
        target_dir = Path(self.dl_dir_input.text().strip()) / spec_num
        if target_dir.exists():
            try:
                os.startfile(str(target_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _open_web_report(self, spec_num: str):
        clean_number = spec_num.replace('.', '')
        url = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        webbrowser.open(url)

    def _handle_document_action(self, combo: QComboBox, doc_type: str, btn: QPushButton):
        c_data = combo.currentData()
        if not c_data: return

        spec_dl_dir = Path(self.dl_dir_input.text().strip()) / c_data['spec_num']
        zip_path = spec_dl_dir / c_data['fname']
        stem = Path(c_data['fname']).stem

        def _process_and_open():
            extracted_docs = []

            # 1. Flat Extraction (If zip exists and hasn't been extracted yet)
            if zip_path.exists():
                try:
                    with zipfile.ZipFile(zip_path, 'r') as z:
                        for member in z.namelist():
                            if '__MACOSX' in member or member.startswith('._'): continue

                            if member.lower().endswith(('.doc', '.docx')):
                                target_file = spec_dl_dir / Path(member).name
                                if not target_file.exists():
                                    target_file.write_bytes(z.read(member))

                                if target_file not in extracted_docs:
                                    extracted_docs.append(target_file)
                except Exception as e:
                    QMessageBox.warning(self, "Extraction Error", f"Failed to extract archive:\n{e}")
                    return

            # Fallback: Find existing word docs on disk if zip didn't provide them (e.g. if zip was deleted manually)
            if not extracted_docs:
                extracted_docs = list(spec_dl_dir.glob(f"{stem}*.doc*"))

            if not extracted_docs:
                QMessageBox.warning(self, "Not Found", "No Word documents found on disk or inside the zip archive.")
                return

            # Lazy-load Converter
            try:
                from modules.word_tools.core.word_converter import WordConverterThread
            except ImportError as e:
                QMessageBox.warning(self, "Import Error", f"Could not import your existing word_converter:\n{e}")
                return

            # 2. Smart Conversion & Opening
            for doc_path in extracted_docs:
                try:
                    if doc_type == 'word':
                        os.startfile(str(doc_path))

                    elif doc_type in ('pdf', 'html'):
                        target_ext = f".{doc_type}"
                        target_path = doc_path.with_suffix(target_ext)

                        if not target_path.exists():
                            orig_text = btn.text()
                            btn.setText("⏳ Converting...")
                            btn.setEnabled(False)

                            conv_thread = WordConverterThread(str(doc_path), doc_type)

                            def on_success(p, c=combo, b=btn, txt=orig_text):
                                try:
                                    os.startfile(p)
                                except Exception as e:
                                    print(f"Error opening converted file: {e}")

                                # Instantly force the row UI to update so the button turns Green ✅
                                c.currentIndexChanged.emit(c.currentIndex())
                                b.setText(txt)
                                b.setEnabled(True)

                            conv_thread.finished_path.connect(on_success)

                            def cleanup(t=conv_thread, b=btn, txt=orig_text):
                                if t in self._download_threads:
                                    self._download_threads.remove(t)
                                if not b.isEnabled():
                                    b.setText(txt)
                                    b.setEnabled(True)

                            conv_thread.finished.connect(cleanup)
                            self._download_threads.append(conv_thread)
                            conv_thread.start()

                        else:
                            os.startfile(str(target_path))

                except Exception as e:
                    QMessageBox.warning(self, "Open Error", f"Could not open/convert {doc_type.upper()}:\n{e}")

        # --- The Workflow Logic Gates ---
        word_exists = any(spec_dl_dir.glob(f"{stem}*.doc*"))
        target_exists = False

        if doc_type == 'word':
            target_exists = word_exists
        elif doc_type == 'pdf':
            target_exists = any(spec_dl_dir.glob(f"{stem}*.pdf"))
        elif doc_type == 'html':
            target_exists = any(spec_dl_dir.glob(f"{stem}*.html"))

        # If the target exists, OR we have the Word doc to convert, OR we have the zip to extract... process locally!
        if target_exists or word_exists or zip_path.exists():
            _process_and_open()
        else:
            # We have nothing. Download the zip from the internet first.
            spec_dl_dir.mkdir(parents=True, exist_ok=True)

            idx = combo.currentIndex()
            orig_text = combo.itemText(idx)
            combo.setItemText(idx, "⏳ Downloading...")
            combo.setEnabled(False)

            thread = SpecDownloadThread(c_data['url'], zip_path)

            def _on_success(zp):
                # Clean any formatting prefixes
                clean_text = orig_text.replace('✅ ', '').replace('⚙️ ', '').replace('⬇️ ', '').strip()
                combo.setItemText(idx, f"✅ {clean_text}")
                combo.setEnabled(True)
                _process_and_open()

            def _on_err(err):
                combo.setItemText(idx, "❌ Error")
                combo.setEnabled(True)
                QMessageBox.critical(self, "Download Failed", f"Network error:\n{err}")

            thread.finished_success.connect(_on_success)
            thread.error.connect(_on_err)

            self._download_threads.append(thread)
            thread.finished.connect(
                lambda t=thread: self._download_threads.remove(t) if t in self._download_threads else None)
            thread.start()

    def _open_table_filters(self):
        dialog = TableFilterDialog(self.db, self.table_filters, self)
        if dialog.exec_():
            self.table_filters = dialog.get_filters()
            self._update_filter_ui()
            self.refresh_table()

    def _clear_table_filters(self):
        self.table_filters = {'series': '', 'tech': '', 'group': '', 'spec_type': 'Any'}
        self._update_filter_ui()
        self.refresh_table()

    def _update_filter_ui(self):
        active_count = 0
        if self.table_filters['series']: active_count += 1
        if self.table_filters['tech']: active_count += 1
        if self.table_filters['group']: active_count += 1
        if self.table_filters['spec_type'] != 'Any': active_count += 1

        if active_count > 0:
            self.filter_btn.setText(f"⚙️ Filters ({active_count})")
            self.filter_btn.setStyleSheet(
                "background-color: #E1F0FF; color: #0078D7; font-weight: bold; border: 1px solid #0078D7;")
            self.clear_filter_btn.setVisible(True)
        else:
            self.filter_btn.setText("⚙️ Table Filters")
            self.filter_btn.setStyleSheet("")
            self.clear_filter_btn.setVisible(False)

    def refresh_table(self):
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()
            base_dl_dir = Path(self.dl_dir_input.text().strip())

            is_filtered = any([
                self.table_filters['series'], self.table_filters['tech'],
                self.table_filters['group'], self.table_filters['spec_type'] != 'Any'
            ])

            if not spec_query and not version_query and not is_filtered:
                self.table.setRowCount(0)
                self.count_label.setText("⌨️ Type a specification number or apply a Filter to begin...")
                self.count_label.setStyleSheet("font-weight: bold; color: #555555;")
                return

            specs = self.db.search_files(
                spec_number=spec_query if spec_query else None,
                release_version=version_query if version_query else None,
                **self.table_filters
            )

            self.table.setRowCount(0)

            grouped_specs = {}
            for row in specs:
                series, spec_num, title, spec_type, filename, version, url = row

                # --- NEW LOGIC: Extract part number from filename ---
                # This safely corrects multipart specs like 23.801-01 missing their suffix in the DB
                if filename:
                    part_match = re.search(r'^\d{4,5}-(\d{2,3})(?:[-_.]|$)', filename)
                    if part_match and "-" not in spec_num:
                        spec_num = f"{spec_num}-{part_match.group(1)}"
                # ----------------------------------------------------

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
                spec_target_dir = base_dl_dir / spec_num

                spec_widget = QWidget()
                spec_layout = QHBoxLayout(spec_widget)
                spec_layout.setContentsMargins(5, 0, 5, 0)

                action_btn = HoverMenuButton("⋮")
                action_btn.setFixedSize(24, 24)
                action_btn.setToolTip("Specification Actions")
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
                    if path.exists():
                        act.setText("📂  Open Local Folder")
                        act.setEnabled(True)
                    else:
                        act.setText("📁  Folder Not Created")
                        act.setEnabled(False)

                menu.aboutToShow.connect(_update_menu_state)
                action_btn.setMenu(menu)

                display_num = f"{data['type']} {spec_num}".strip()
                spec_label = QLabel(display_num)

                spec_layout.addWidget(action_btn)
                spec_layout.addWidget(spec_label)
                spec_layout.addStretch()
                self.table.setCellWidget(row_idx, 0, spec_widget)

                self.table.setItem(row_idx, 1, QTableWidgetItem(data['title'] if data['title'] else "Unknown Title"))

                # --- COLUMN 2: Documents Action Bar ---
                version_combo = QComboBox()

                # Type-safe semantic version parser
                def parse_ver(v_str):
                    return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split('.')]

                sorted_versions = sorted(data['versions'], key=lambda x: parse_ver(x[0]), reverse=True)

                for ver, url, fname in sorted_versions:
                    zip_path = spec_target_dir / fname
                    is_dl = zip_path.exists()
                    status = "✅ " if is_dl else ""
                    version_combo.addItem(f"{status}v{ver}", userData={
                        'url': url, 'fname': fname, 'spec_num': spec_num, 'is_downloaded': is_dl
                    })

                # Action Buttons
                word_btn = QPushButton("📝 Word")
                pdf_btn = QPushButton("📕 PDF")
                html_btn = QPushButton("🌐 HTML")

                for b in (word_btn, pdf_btn, html_btn):
                    b.setCursor(Qt.PointingHandCursor)

                # ---> 3-STATE DYNAMIC UI CHECKER
                def _update_btn_state(index_ignore=0, c=version_combo, wb=word_btn, pb=pdf_btn, hb=html_btn):
                    c_data = c.currentData()
                    if not c_data: return

                    # ---> FIX: Resolve the directory dynamically from the combo box data!
                    # This prevents Python loop closures from pointing to the wrong folder.
                    current_dir = Path(self.dl_dir_input.text().strip()) / c_data['spec_num']

                    stem = Path(c_data['fname']).stem
                    zip_exists = (current_dir / c_data['fname']).exists()

                    word_exists = any(current_dir.glob(f"{stem}*.doc*"))
                    pdf_exists = any(current_dir.glob(f"{stem}*.pdf"))
                    html_exists = any(current_dir.glob(f"{stem}*.html"))

                    def style_btn(btn, exists, icon, name):
                        if exists:
                            # State 1: File is ready locally
                            btn.setText(f"{icon} {name} ✅")
                            btn.setStyleSheet("""
                                                QPushButton { padding: 4px 6px; font-size: 11px; font-weight: bold; background-color: #E8F5E9; color: #2E7D32; border: 1px solid #2E7D32; border-radius: 3px; } 
                                                QPushButton:hover { background-color: #C8E6C9; }
                                            """)
                        elif word_exists or zip_exists:
                            # State 2: Can be processed entirely offline (Zip->Extract, or Word->Convert)
                            action = "Extract" if name == "Word" else "Convert"
                            btn.setText(f"⚙️ {action}")
                            btn.setStyleSheet("""
                                                QPushButton { padding: 4px 6px; font-size: 11px; font-weight: bold; background-color: #FFF3E0; color: #E65100; border: 1px solid #FFB74D; border-radius: 3px; } 
                                                QPushButton:hover { background-color: #FFE0B2; }
                                            """)
                        else:
                            # State 3: Requires Network Download
                            btn.setText(f"⬇️ Get {name}")
                            btn.setStyleSheet("""
                                                QPushButton { padding: 4px 6px; font-size: 11px; background-color: transparent; color: #555; border: 1px solid #CCC; border-radius: 3px; } 
                                                QPushButton:hover { background-color: #E1F0FF; color: #0078D7; border: 1px solid #0078D7; }
                                            """)

                    style_btn(wb, word_exists, "📝", "Word")
                    style_btn(pb, pdf_exists, "📕", "PDF")
                    style_btn(hb, html_exists, "🌐", "HTML")

                # Bind the UI style updater
                version_combo.currentIndexChanged.connect(_update_btn_state)
                _update_btn_state()

                # Bind the specific buttons to the specific document extensions
                word_btn.clicked.connect(
                    lambda _, c=version_combo, b=word_btn: self._handle_document_action(c, 'word', b))
                pdf_btn.clicked.connect(lambda _, c=version_combo, b=pdf_btn: self._handle_document_action(c, 'pdf', b))
                html_btn.clicked.connect(
                    lambda _, c=version_combo, b=html_btn: self._handle_document_action(c, 'html', b))

                cell_widget = QWidget()
                layout = QHBoxLayout(cell_widget)
                layout.setContentsMargins(0, 0, 0, 0)
                layout.setSpacing(4)

                layout.addWidget(version_combo)
                layout.addWidget(word_btn)
                layout.addWidget(pdf_btn)
                layout.addWidget(html_btn)

                self.table.setCellWidget(row_idx, 2, cell_widget)

        except Exception as e:
            print(f"Error during refresh_table: {e}")