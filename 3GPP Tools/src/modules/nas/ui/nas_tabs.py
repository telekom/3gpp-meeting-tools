import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import (
    NASFetchAndImportThread,
    find_cached_spec_file,
)
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase


class NASVersionSelectDialog(QDialog):
    """Dialog allowing selection of indexed TS 24.501 versions with natural numerical sorting."""

    def __init__(
            self,
            specs_db: SpecsDatabase,
            nas_db: NASDatabase,
            cache_dir: Path,
            parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.nas_db = nas_db
        self.cache_dir = cache_dir
        self.selected_file_info: Optional[Dict[str, Any]] = None

        self.setWindowTitle("Select TS 24.501 Version to Ingest")
        self.resize(680, 420)
        self._setup_ui()
        self._load_available_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        info_lbl = QLabel(
            "Select a TS 24.501 version from the specification archive.\n"
            "If not cached locally, it will be downloaded automatically from the 3GPP FTP server."
        )
        info_lbl.setStyleSheet("color: #555; padding-bottom: 5px;")
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Version", "Filename", "Local Cache", "NAS DB Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("📥 Fetch && Ingest")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _load_available_versions(self):
        spec_files = self.specs_db.search_files(spec_number="24.501")
        imported_versions = {v["version"] for v in self.nas_db.get_imported_versions()}

        # Sort numerically descending (v20.0.0 -> v19.7.0 -> ... -> v2.0.0)
        spec_files = sorted(
            spec_files,
            key=lambda row: parse_version_tuple(row[5]),
            reverse=True,
        )

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(spec_files):
            _, spec_num, _, _, filename, version, url = row_data
            self.table.insertRow(row_idx)

            v_item = QTableWidgetItem(f"v{version}")
            v_item.setData(
                Qt.UserRole,
                {
                    "spec_number": spec_num,
                    "version": version,
                    "filename": filename,
                    "url": url,
                },
            )
            self.table.setItem(row_idx, 0, v_item)
            self.table.setItem(row_idx, 1, QTableWidgetItem(filename))

            cached_file = find_cached_spec_file(filename, spec_num)
            if cached_file:
                cache_text = f"🟢 Cached ({cached_file.suffix[1:].upper()})"
                cache_item = QTableWidgetItem(cache_text)
                cache_item.setForeground(Qt.darkGreen)
            else:
                cache_item = QTableWidgetItem("🌐 Remote (FTP)")
                cache_item.setForeground(Qt.darkGray)
            self.table.setItem(row_idx, 2, cache_item)

            in_db = version in imported_versions
            db_item = QTableWidgetItem("✅ Ingested" if in_db else "⚪ Ready")
            if in_db:
                db_item.setForeground(Qt.blue)
            self.table.setItem(row_idx, 3, db_item)

    def _on_item_double_clicked(self, item: QTableWidgetItem):
        self._on_accept()

    def _on_accept(self):
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "Selection Required", "Please select a specification version to import.")
            return
        row = self.table.currentRow()
        self.selected_file_info = self.table.item(row, 0).data(Qt.UserRole)
        self.accept()


class NASTab(QWidget):
    log_msg = pyqtSignal(str, int)

    def __init__(self, nas_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.nas_db_path = Path(nas_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = Path.home() / "3GPP_Delegate_Helper" / "specs"

        self.db = NASDatabase(self.nas_db_path)
        self.specs_db = (
            SpecsDatabase(self.specs_db_path)
            if self.specs_db_path and self.specs_db_path.exists()
            else None
        )

        self.selected_version_ids: List[int] = []
        self.current_selected_message_name: Optional[str] = None
        self._updating_checks: bool = False
        self._setup_ui()
        self.refresh_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Toolbar
        toolbar = QHBoxLayout()

        self.fetch_btn = QPushButton("📥 Import from Specs DB")
        self.fetch_btn.clicked.connect(self._on_fetch_from_specs_db_clicked)
        toolbar.addWidget(self.fetch_btn)

        self.import_file_btn = QPushButton("📁 Import Local .docx")
        self.import_file_btn.clicked.connect(self._on_import_local_file_clicked)
        toolbar.addWidget(self.import_file_btn)

        self.clear_ver_btn = QPushButton("🗑️ Clear Version")
        self.clear_ver_btn.clicked.connect(self._on_clear_version_clicked)
        toolbar.addWidget(self.clear_ver_btn)

        self.wipe_db_btn = QPushButton("⚠️ Wipe NAS DB")
        self.wipe_db_btn.setStyleSheet("color: #D32F2F; font-weight: bold;")
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        toolbar.addWidget(self.wipe_db_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Main Splitter
        main_splitter = QSplitter(Qt.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        ver_group = QGroupBox("Specification Versions (Check Multiple or 'All')")
        ver_layout = QVBoxLayout(ver_group)
        self.version_list = QListWidget()
        self.version_list.itemChanged.connect(self._on_version_item_changed)
        ver_layout.addWidget(self.version_list)
        left_layout.addWidget(ver_group)

        msg_group = QGroupBox("NAS Messages")
        msg_layout = QVBoxLayout(msg_group)
        self.msg_search = QLineEdit()
        self.msg_search.setPlaceholderText("Filter messages (e.g. REGISTRATION)...")
        self.msg_search.textChanged.connect(self._filter_messages)
        msg_layout.addWidget(self.msg_search)

        self.msg_list = QListWidget()
        self.msg_list.itemClicked.connect(self._on_message_clicked)
        msg_layout.addWidget(self.msg_list)
        left_layout.addWidget(msg_group)

        main_splitter.addWidget(left_widget)

        right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)
        self.matrix_title = QLabel("Select a Message to View Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        matrix_layout.addWidget(self.matrix_title)

        self.matrix_table = QTableView()
        self.matrix_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.matrix_table.clicked.connect(self._on_table_cell_clicked)
        matrix_layout.addWidget(self.matrix_table)
        right_splitter.addWidget(matrix_widget)

        # Escaped ampersand (&&) renders as literal '&'
        inspector_group = QGroupBox("Clause 9 Structure && Coding Inspector")
        inspector_layout = QVBoxLayout(inspector_group)
        self.inspector_text = QTextEdit()
        self.inspector_text.setReadOnly(True)
        self.inspector_text.setPlaceholderText(
            "Click on an Information Element above to inspect its Clause 9 details...")
        inspector_layout.addWidget(self.inspector_text)
        right_splitter.addWidget(inspector_group)

        right_splitter.setSizes([450, 200])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([280, 670])

        layout.addWidget(main_splitter)

    def refresh_versions(self):
        self._updating_checks = True
        self.version_list.clear()
        versions = self.db.get_imported_versions()

        if versions:
            all_item = QListWidgetItem("All Versions")
            all_item.setData(Qt.UserRole, -1)
            all_item.setFlags(all_item.flags() | Qt.ItemIsUserCheckable)
            all_item.setCheckState(Qt.Checked)
            self.version_list.addItem(all_item)

            for v in versions:
                item = QListWidgetItem(f"TS {v['spec_number']} v{v['version']}")
                item.setData(Qt.UserRole, v["id"])
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.version_list.addItem(item)

            self.selected_version_ids = [v["id"] for v in versions]
        else:
            self.selected_version_ids = []

        self._updating_checks = False
        self._populate_messages()

    def _on_version_item_changed(self, item: QListWidgetItem):
        """Handles native checkbox clicks and synchronizes the 'All Versions' master toggle."""
        if self._updating_checks:
            return

        self._updating_checks = True
        item_id = item.data(Qt.UserRole)
        is_checked = item.checkState() == Qt.Checked

        # 1. Master "All Versions" Toggle
        if item_id == -1:
            for i in range(1, self.version_list.count()):
                self.version_list.item(i).setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
        else:
            # 2. Individual Version Checkbox
            all_item = self.version_list.item(0)
            if all_item:
                total_versions = self.version_list.count() - 1
                checked_versions = sum(
                    1
                    for i in range(1, self.version_list.count())
                    if self.version_list.item(i).checkState() == Qt.Checked
                )
                if checked_versions == total_versions and total_versions > 0:
                    all_item.setCheckState(Qt.Checked)
                else:
                    all_item.setCheckState(Qt.Unchecked)

        # 3. Update active version IDs
        self.selected_version_ids = [
            self.version_list.item(i).data(Qt.UserRole)
            for i in range(1, self.version_list.count())
            if self.version_list.item(i).checkState() == Qt.Checked
        ]

        self._updating_checks = False

        # Repopulates messages and preserves the active selection automatically
        self._populate_messages()

    def _populate_messages(self):
        """Populates the message list while maintaining the currently selected message."""
        target_msg_name = self.current_selected_message_name
        self.msg_list.clear()

        if not self.selected_version_ids:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
            self.current_selected_message_name = None
            return

        messages = self.db.get_messages_list(self.selected_version_ids)
        target_item = None
        filter_text = self.msg_search.text().strip().lower()

        for m in messages:
            msg_name = m["message_name"]
            item_text = f"{msg_name} ({m['clause']})"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, msg_name)

            if filter_text and filter_text not in item_text.lower():
                item.setHidden(True)

            self.msg_list.addItem(item)

            if target_msg_name and msg_name == target_msg_name:
                target_item = item

        # If previous selection still exists in the newly chosen version set, restore it
        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        else:
            self.current_selected_message_name = None
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")

    def _filter_messages(self, text: str):
        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def _on_message_clicked(self, item: QListWidgetItem):
        msg_name = item.data(Qt.UserRole)
        self.current_selected_message_name = msg_name
        self.matrix_title.setText(f"Message: {msg_name}")

        df = self.db.get_message_evolution_df(msg_name, self.selected_version_ids)
        model = NASEvolutionMatrixModel(df)
        self.matrix_table.setModel(model)
        self.matrix_table.resizeColumnsToContents()

    def _on_table_cell_clicked(self, index):
        model = self.matrix_table.model()
        if not model:
            return

        type_ref_idx = model.index(index.row(), 2)
        type_ref = model.data(type_ref_idx, Qt.DisplayRole)

        if type_ref:
            match = re.search(r"((?:9|D\.6)(?:\.[0-9A-Za-z]+)+)", str(type_ref))
            if match:
                cl = match.group(1).strip()
                ie_def = self.db.get_ie_definition(cl)
                if ie_def and ie_def.get("raw_description"):
                    self.inspector_text.setHtml(ie_def["raw_description"])
                    return

        self.inspector_text.setPlainText(
            f"Type / Reference: {type_ref}\n(No Clause 9 definition found for this reference)")

    def _on_fetch_from_specs_db_clicked(self):
        if not self.specs_db:
            QMessageBox.warning(
                self,
                "Specs DB Unavailable",
                "The 3GPP Specifications database (3gpp_data.db) is not configured or reachable.",
            )
            return

        dialog = NASVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)
        if dialog.exec_() == QDialog.Accepted and dialog.selected_file_info:
            info = dialog.selected_file_info
            self._start_ingestion_thread(
                spec_number=info["spec_number"],
                version=info["version"],
                filename=info["filename"],
                file_url=info["url"],
            )

    def _on_import_local_file_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select TS 24.501 Specification (.docx)",
            "",
            "Word Files (*.docx)",
        )
        if not file_path:
            return

        p = Path(file_path)
        self._start_ingestion_thread(
            spec_number="24.501",
            version="",
            filename=p.name,
            file_url="",
            local_docx_path=p,
        )

    def _start_ingestion_thread(
            self,
            spec_number: str,
            version: str,
            filename: str,
            file_url: str,
            local_docx_path: Optional[Path] = None,
    ):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.import_file_btn.setEnabled(False)

        self.thread = NASFetchAndImportThread(
            nas_db_path=self.nas_db_path,
            spec_number=spec_number,
            version=version,
            filename=filename,
            file_url=file_url,
            cache_dir=self.cache_dir,
            local_docx_path=local_docx_path,
        )
        self.thread.progress.connect(self._on_import_progress)
        self.thread.finished_success.connect(self._on_import_success)
        self.thread.error.connect(self._on_import_error)
        self.thread.start()

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        self.log_msg.emit(msg, logging.INFO)

    def _on_import_success(self, spec_number: str, version: str, count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        self.log_msg.emit(
            f"✅ Successfully ingested TS {spec_number} v{version} ({count} messages).",
            logging.INFO,
        )
        self.refresh_versions()

    def _on_import_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        QMessageBox.critical(self, "Ingestion Error", err)
        self.log_msg.emit(f"❌ Ingestion failed: {err}", logging.ERROR)

    def _on_clear_version_clicked(self):
        checked_items = [
            self.version_list.item(i)
            for i in range(1, self.version_list.count())
            if self.version_list.item(i).checkState() == Qt.Checked
        ]
        if not checked_items:
            QMessageBox.warning(self, "Select Version", "Please check at least one specific version to clear.")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete {len(checked_items)} checked specification version(s)?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for item in checked_items:
                v_text = item.text().replace("TS 24.501 v", "")
                self.db.clear_version("24.501", v_text)
            self.current_selected_message_name = None
            self.refresh_versions()

    def _on_wipe_db_clicked(self):
        reply = QMessageBox.critical(
            self,
            "Confirm Wipe",
            "This will delete ALL imported NAS specifications and tables. Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.db.wipe_database()
            self.current_selected_message_name = None
            self.refresh_versions()
            self.msg_list.clear()
            self.matrix_table.setModel(None)
            self.inspector_text.clear()
            self.log_msg.emit("🧹 NAS Database wiped.", logging.INFO)