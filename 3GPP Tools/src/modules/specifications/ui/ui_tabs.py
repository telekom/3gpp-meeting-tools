# --- File: modules/specs_db/ui_tabs.py ---
import json
import os
import zipfile
from pathlib import Path

from PyQt5.QtCore import pyqtSignal, Qt, QTimer, QThread
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QLineEdit, QComboBox, QMenu, QAbstractItemView,
                             QDialog, QFormLayout, QFileDialog, QMessageBox)

from core.network.session import NetworkSession
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.utils.utils import open_extracted_documents


class SpecInfoDialog(QDialog):
    """Dynamic Popup to display all Database information about a Specification."""

    def __init__(self, details: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Specification Details: {details.get('number', '')}")
        self.setMinimumWidth(450)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; }")

        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        for key, value in details.items():
            if key in ('id', 'series_id') or not value:
                continue

            display_key = key.replace('_', ' ').title()
            val_label = QLabel(str(value))
            val_label.setWordWrap(True)
            val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

            key_label = QLabel(f"<b>{display_key}:</b>")
            key_label.setStyleSheet("color: #444;")
            form.addRow(key_label, val_label)

        layout.addLayout(form)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)


class SpecDownloadThread(QThread):
    """Background worker to download and unzip files without freezing the GUI."""
    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, url: str, zip_path: Path):
        super().__init__()
        self.url = url
        self.zip_path = zip_path

    def run(self):
        try:
            NetworkSession.download_file(self.url, self.zip_path)
            extract_dir = self.zip_path.with_suffix('')
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            self.finished_success.emit(extract_dir)
        except Exception as e:
            self.error.emit(str(e))


class AdvancedSyncDialog(QDialog):
    def __init__(self, db: SpecsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Advanced Filtered Sync")
        self.setModal(True)
        self.resize(450, 250)
        self.matching_specs = []

        layout = QVBoxLayout(self)
        info_label = QLabel("Note: Filters apply to specifications already discovered in your local database. "
                            "To discover brand new specifications, run a 'Full Sync' first.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666666; font-style: italic; margin-bottom: 10px;")
        layout.addWidget(info_label)

        form = QFormLayout()
        self.series_input = QLineEdit()
        self.series_input.setPlaceholderText("e.g. 23, 24")
        self.tech_input = QLineEdit()
        self.tech_input.setPlaceholderText("e.g. 5G")
        self.group_input = QLineEdit()
        self.group_input.setPlaceholderText("e.g. SA2")

        type_layout = QHBoxLayout()
        self.ts_cb = QCheckBox("TS")
        self.ts_cb.setChecked(True)
        self.tr_cb = QCheckBox("TR")
        self.tr_cb.setChecked(True)
        type_layout.addWidget(self.ts_cb)
        type_layout.addWidget(self.tr_cb)
        type_layout.addStretch()

        form.addRow("Series:", self.series_input)
        form.addRow("Radio Tech:", self.tech_input)
        form.addRow("Working Group:", self.group_input)
        form.addRow("Type:", type_layout)
        layout.addLayout(form)

        self.count_label = QLabel("Matching specifications: 0")
        self.count_label.setStyleSheet("font-weight: bold; color: #395396; margin-top: 10px;")
        layout.addWidget(self.count_label)

        btn_layout = QHBoxLayout()
        self.sync_btn = QPushButton("🚀 Start Sync")
        self.sync_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.sync_btn)
        layout.addLayout(btn_layout)

        self.series_input.textChanged.connect(self.update_count)
        self.tech_input.textChanged.connect(self.update_count)
        self.group_input.textChanged.connect(self.update_count)
        self.ts_cb.stateChanged.connect(self.update_count)
        self.tr_cb.stateChanged.connect(self.update_count)

        self.update_count()

    def update_count(self):
        series = self.series_input.text().strip()
        tech = self.tech_input.text().strip()
        group = self.group_input.text().strip()
        types = []
        if self.ts_cb.isChecked(): types.append("TS")
        if self.tr_cb.isChecked(): types.append("TR")

        self.matching_specs = self.db.get_filtered_specs(series, tech, group, types)
        count = len(self.matching_specs)
        self.count_label.setText(f"Matching specifications in local DB: {count}")
        self.sync_btn.setEnabled(count > 0)

class TableFilterDialog(QDialog):
    """Compact dialog for filtering the currently displayed table results."""
    def __init__(self, current_filters: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filter Specifications")
        self.setModal(True)
        self.resize(350, 200)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.series_input = QLineEdit(current_filters.get('series', ''))
        self.series_input.setPlaceholderText("e.g. 23, 24")
        self.tech_input = QLineEdit(current_filters.get('tech', ''))
        self.tech_input.setPlaceholderText("e.g. 5G")
        self.group_input = QLineEdit(current_filters.get('group', ''))
        self.group_input.setPlaceholderText("e.g. SA2")

        type_layout = QHBoxLayout()
        active_types = current_filters.get('types', ['TS', 'TR'])
        self.ts_cb = QCheckBox("TS")
        self.ts_cb.setChecked('TS' in active_types)
        self.tr_cb = QCheckBox("TR")
        self.tr_cb.setChecked('TR' in active_types)
        type_layout.addWidget(self.ts_cb)
        type_layout.addWidget(self.tr_cb)
        type_layout.addStretch()

        form.addRow("Series:", self.series_input)
        form.addRow("Radio Tech:", self.tech_input)
        form.addRow("Working Group:", self.group_input)
        form.addRow("Type:", type_layout)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        apply_btn = QPushButton("✅ Apply Filters")
        apply_btn.clicked.connect(self.accept)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_and_accept)

        btn_layout.addStretch()
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(apply_btn)
        layout.addLayout(btn_layout)

    def _clear_and_accept(self):
        self.series_input.clear()
        self.tech_input.clear()
        self.group_input.clear()
        self.ts_cb.setChecked(True)
        self.tr_cb.setChecked(True)
        self.accept()

    def get_filters(self) -> dict:
        types = []
        if self.ts_cb.isChecked(): types.append("TS")
        if self.tr_cb.isChecked(): types.append("TR")
        return {
            'series': self.series_input.text().strip(),
            'tech': self.tech_input.text().strip(),
            'group': self.group_input.text().strip(),
            'types': types
        }

class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
    update_specific_requested = pyqtSignal(list, bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self._download_threads = []

        self.config_file = db_path.parent / "specs_config.json"
        self.default_dl_dir = self._load_settings()

        self.table_filters = {'series': '', 'tech': '', 'group': '', 'types': ['TS', 'TR']}

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

    def _open_download_dir(self):
        """Safely opens the configured download directory in File Explorer."""
        target_dir = Path(self.dl_dir_input.text().strip())

        # If the user typed a new path but hasn't downloaded anything yet, create it.
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

    def _save_settings(self):
        try:
            current_dir = self.dl_dir_input.text().strip()
            data = {'download_dir': current_dir}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving specifications config: {e}")

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 10, 5, 5)

        # --- Download Directory Settings ---
        dir_layout = QHBoxLayout()

        self.dl_dir_input = QLineEdit(self.default_dl_dir)
        self.dl_dir_input.editingFinished.connect(self._save_settings)

        browse_btn = QPushButton("📂 Browse...")
        browse_btn.clicked.connect(self._browse_dir)

        # ---> NEW: Compact "Open" Button
        open_dir_btn = QPushButton("↗️ Open")
        open_dir_btn.setToolTip("Open the base download folder in Windows Explorer")
        open_dir_btn.setCursor(Qt.PointingHandCursor)
        open_dir_btn.clicked.connect(self._open_download_dir)

        dir_layout.addWidget(QLabel("💾 Download Path:"))
        dir_layout.addWidget(self.dl_dir_input)
        dir_layout.addWidget(browse_btn)
        dir_layout.addWidget(open_dir_btn)  # Added right next to Browse
        main_layout.addLayout(dir_layout)

        # --- COMPACT TOP PANEL ---
        top_layout = QHBoxLayout()

        self.full_sync_btn = QPushButton("🔄 Full Sync")
        self.full_sync_btn.clicked.connect(lambda: self.update_db_requested.emit(self.force_meta_checkbox.isChecked()))
        top_layout.addWidget(self.full_sync_btn)

        self.adv_sync_btn = QPushButton("⚙️ Filtered Sync")
        self.adv_sync_btn.clicked.connect(self._open_advanced_sync)
        top_layout.addWidget(self.adv_sync_btn)

        self.force_meta_checkbox = QCheckBox("Force Metadata")
        top_layout.addWidget(self.force_meta_checkbox)

        top_layout.addSpacing(15)

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

        self.filter_btn = QPushButton("⚙️ Filters")
        self.filter_btn.setCursor(Qt.PointingHandCursor)
        self.filter_btn.clicked.connect(self._open_table_filters)

        self.clear_filter_btn = QPushButton("❌")
        self.clear_filter_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filter_btn.setToolTip("Clear Active Filters")
        self.clear_filter_btn.setFixedWidth(24)
        self.clear_filter_btn.setVisible(False)
        self.clear_filter_btn.clicked.connect(self._clear_table_filters)

        top_layout.addWidget(self.filter_btn)
        top_layout.addWidget(self.clear_filter_btn)

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

    def _browse_dir(self):
        new_dir = QFileDialog.getExistingDirectory(self, "Select Download Directory", self.dl_dir_input.text())
        if new_dir:
            self.dl_dir_input.setText(new_dir)
            self._save_settings()
            self.refresh_table()

    def _open_advanced_sync(self):
        dialog = AdvancedSyncDialog(self.db, self)
        if dialog.exec_():
            target_specs = dialog.matching_specs
            if target_specs:
                force_meta = self.force_meta_checkbox.isChecked()
                self.update_specific_requested.emit(target_specs, force_meta)

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
        dialog = SpecInfoDialog(details, self)
        dialog.exec_()

    def _open_spec_folder(self, spec_num: str):
        """Opens the base folder for the specification (e.g. .../specs/23.501)"""
        target_dir = Path(self.dl_dir_input.text().strip()) / spec_num
        if target_dir.exists():
            try:
                os.startfile(str(target_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    # ---> UPGRADED: Now accepts the folder_btn to dynamically reveal it
    def _handle_download_action(self, combo: QComboBox, btn: QPushButton, folder_btn: QPushButton):
        c_data = combo.currentData()
        if not c_data: return

        dl_dir = Path(self.dl_dir_input.text().strip())
        spec_dl_dir = dl_dir / c_data['spec_num']

        # Action: OPEN DOCUMENTS
        if c_data['is_downloaded']:
            zip_path = spec_dl_dir / c_data['fname']
            extracted_dir = zip_path.with_suffix('')

            if not extracted_dir.exists() and zip_path.exists():
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(extracted_dir)
                except Exception as e:
                    QMessageBox.warning(self, "Extraction Error", f"Could not extract archive:\n{e}")
                    return

            if extracted_dir.exists():
                opened_files = open_extracted_documents(extracted_dir)
                if not opened_files:
                    try:
                        os.startfile(str(extracted_dir))
                    except Exception as e:
                        QMessageBox.warning(self, "Open Error", f"Could not open directory:\n{e}")
            else:
                try:
                    os.startfile(str(spec_dl_dir))
                except Exception as e:
                    QMessageBox.warning(self, "Open Error", f"Could not open location:\n{e}")
            return

        # Action: DOWNLOAD
        spec_dl_dir.mkdir(parents=True, exist_ok=True)
        # ---> INSTANT UX: Reveal the folder button instantly because we just created the directory!
        folder_btn.setVisible(True)

        zip_path = spec_dl_dir / c_data['fname']

        btn.setText("⏳ Downloading...")
        btn.setEnabled(False)

        thread = SpecDownloadThread(c_data['url'], zip_path)
        thread.finished_success.connect(lambda p: self._on_download_success(combo, btn))
        thread.error.connect(lambda e: self._on_download_error(btn, e))

        self._download_threads.append(thread)
        thread.finished.connect(lambda t=thread: self._download_threads.remove(t))
        thread.start()

    def _handle_open_folder_action(self, combo: QComboBox):
        """Opens the extracted folder for the specific version, or the base spec folder if not extracted."""
        c_data = combo.currentData()
        if not c_data: return

        dl_dir = Path(self.dl_dir_input.text().strip())
        spec_dl_dir = dl_dir / c_data['spec_num']
        zip_path = spec_dl_dir / c_data['fname']
        extracted_dir = zip_path.with_suffix('')

        # Prefer opening the fully extracted specific version folder if it exists
        target_path = extracted_dir if extracted_dir.exists() else spec_dl_dir

        if not target_path.exists():
            QMessageBox.information(self, "Not Found", "Folder does not exist yet. Please download the file first.")
            return

        try:
            os.startfile(str(target_path))
        except Exception as e:
            QMessageBox.warning(self, "Open Error", f"Could not open directory:\n{e}")

    def _on_download_success(self, combo: QComboBox, btn: QPushButton):
        c_data = combo.currentData()
        c_data['is_downloaded'] = True

        idx = combo.currentIndex()
        text = combo.itemText(idx)
        if not text.startswith("✅"):
            combo.setItemText(idx, f"✅ {text}")

        btn.setText("📂 Open")
        btn.setEnabled(True)

    def _on_download_error(self, btn: QPushButton, error_msg: str):
        btn.setText("❌ Error")
        btn.setEnabled(True)
        QMessageBox.critical(self, "Download Failed", f"Network error during download:\n{error_msg}")

    def _open_table_filters(self):
        dialog = TableFilterDialog(self.table_filters, self)
        if dialog.exec_():
            self.table_filters = dialog.get_filters()
            self._update_filter_ui()
            self.refresh_table()

    def _clear_table_filters(self):
        self.table_filters = {'series': '', 'tech': '', 'group': '', 'types': ['TS', 'TR']}
        self._update_filter_ui()
        self.refresh_table()

    def _update_filter_ui(self):
        """Changes the button color and shows/hides the clear button if filters are active."""
        active_count = 0
        if self.table_filters['series']: active_count += 1
        if self.table_filters['tech']: active_count += 1
        if self.table_filters['group']: active_count += 1
        if len(self.table_filters['types']) < 2: active_count += 1

        if active_count > 0:
            self.filter_btn.setText(f"⚙️ Filters ({active_count})")
            self.filter_btn.setStyleSheet(
                "background-color: #E1F0FF; color: #0078D7; font-weight: bold; border: 1px solid #0078D7;")
            self.clear_filter_btn.setVisible(True)
        else:
            self.filter_btn.setText("⚙️ Filters")
            self.filter_btn.setStyleSheet("")
            self.clear_filter_btn.setVisible(False)

    def refresh_table(self):
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()
            base_dl_dir = Path(self.dl_dir_input.text().strip())

            # Only require typing if NO filters are set to prevent loading 4000 rows
            is_filtered = any([self.table_filters['series'], self.table_filters['tech'], self.table_filters['group']])

            if not spec_query and not version_query and not is_filtered:
                self.table.setRowCount(0)
                self.count_label.setText("⌨️ Type a specification number or apply a Filter to begin...")
                self.count_label.setStyleSheet("font-weight: bold; color: #555555;")
                return

            # ---> UPGRADED: Pass the unpacked dictionary directly to the new search_files
            specs = self.db.search_files(
                spec_number=spec_query if spec_query else None,
                release_version=version_query if version_query else None,
                **self.table_filters
            )

            # ... (The rest of the refresh_table rendering loop remains EXACTLY the same) ...
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

                # Resolve the directory early so we can check if it exists
                spec_target_dir = base_dl_dir / spec_num

                # --- 1. COLUMN 0: Info Button + Folder Button + Spec Number ---
                spec_widget = QWidget()
                spec_layout = QHBoxLayout(spec_widget)
                spec_layout.setContentsMargins(5, 0, 5, 0)

                info_btn = QPushButton("ⓘ")
                info_btn.setFixedSize(24, 24)
                info_btn.setToolTip("View full specification details")
                info_btn.setCursor(Qt.PointingHandCursor)
                info_btn.setStyleSheet("""
                                QPushButton { border: none; background: transparent; color: #0078D7; font-size: 18px; padding: 0px; margin: 0px; }
                                QPushButton:hover { color: #004A85; }
                            """)
                info_btn.clicked.connect(lambda _, s=spec_num: self._show_spec_info(s))

                # ---> NEW: Specification-Level Folder Button
                spec_folder_btn = QPushButton("📂")
                spec_folder_btn.setFixedSize(24, 24)
                spec_folder_btn.setToolTip("Open specification folder")
                spec_folder_btn.setCursor(Qt.PointingHandCursor)
                spec_folder_btn.setStyleSheet("""
                                QPushButton { border: none; background: transparent; font-size: 16px; padding: 0px; margin: 0px; }
                                QPushButton:hover { font-size: 18px; }
                            """)
                spec_folder_btn.setVisible(spec_target_dir.exists())  # Only show if downloaded!
                spec_folder_btn.clicked.connect(lambda _, s=spec_num: self._open_spec_folder(s))

                display_num = f"{data['type']} {spec_num}".strip()
                spec_label = QLabel(display_num)

                spec_layout.addWidget(info_btn)
                spec_layout.addWidget(spec_folder_btn)
                spec_layout.addWidget(spec_label)
                spec_layout.addStretch()
                self.table.setCellWidget(row_idx, 0, spec_widget)

                # --- 2. COLUMN 1: Title ---
                self.table.setItem(row_idx, 1, QTableWidgetItem(data['title'] if data['title'] else "Unknown Title"))

                # --- 3. COLUMN 2: Version Dropdown & Download Button ---
                version_combo = QComboBox()

                def parse_ver(v_str):
                    return [int(x) if x.isdigit() else x for x in str(v_str).split('.')]

                sorted_versions = sorted(data['versions'], key=lambda x: parse_ver(x[0]), reverse=True)

                for ver, url, fname in sorted_versions:
                    zip_path = spec_target_dir / fname
                    extracted_dir = zip_path.with_suffix('')
                    is_dl = zip_path.exists() or extracted_dir.exists()

                    status = "✅ " if is_dl else ""
                    version_combo.addItem(f"{status}v{ver}", userData={
                        'url': url, 'fname': fname, 'spec_num': spec_num, 'is_downloaded': is_dl
                    })

                download_btn = QPushButton()
                download_btn.setCursor(Qt.PointingHandCursor)

                def _update_btn_state(index_ignore=0, c=version_combo, b=download_btn):
                    c_data = c.currentData()
                    if c_data and c_data.get('is_downloaded'):
                        b.setText("📄 Open Docs")
                        b.setStyleSheet("font-weight: bold; color: #1E88E5;")
                    else:
                        b.setText("⬇️ Download")
                        b.setStyleSheet("")

                version_combo.currentIndexChanged.connect(_update_btn_state)
                _update_btn_state()

                # ---> Note: We pass 'spec_folder_btn' into the action so it can be revealed when downloaded!
                download_btn.clicked.connect(
                    lambda _, c=version_combo, b=download_btn, fb=spec_folder_btn: self._handle_download_action(c, b,
                                                                                                                fb))

                cell_widget = QWidget()
                layout = QHBoxLayout(cell_widget)
                layout.setContentsMargins(0, 0, 0, 0)
                layout.addWidget(version_combo)
                layout.addWidget(download_btn)

                self.table.setCellWidget(row_idx, 2, cell_widget)

        except Exception as e:
            print(f"Error during refresh_table: {e}")