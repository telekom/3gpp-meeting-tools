import json
import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_WARNING,
    BUTTON_STYLE_TOOLBAR_DANGER,
    COMBOBOX_STYLE_TOOLBAR,
)
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import NASFetchAndImportThread
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.ui.nas_components import NASInspectorWidget, NASVersionTreeWidget
from modules.nas.ui.nas_dialogs import NASVersionSelectDialog
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase

RE_SPEC_FROM_COL = re.compile(r"(?:24|29|36|38)\.[0-9]{3}")
RE_CLAUSE_FROM_REF = re.compile(r"((?:9|8|7|6|5|D\.6)(?:\.[0-9A-Za-z]+)+)")


class NASMessageListWorker(QThread):
    """Background worker to query message lists without freezing the UI."""

    results_ready = pyqtSignal(list, int)  # (messages, request_id)

    def __init__(
            self,
            db: NASDatabase,
            ie_query: str,
            version_ids: List[int],
            search_desc: bool,
            request_id: int,
    ):
        super().__init__()
        self.db = db
        self.ie_query = ie_query
        self.version_ids = version_ids
        self.search_desc = search_desc
        self.request_id = request_id

    def run(self):
        try:
            if self.ie_query:
                messages = self.db.get_messages_by_ie_search(
                    ie_query=self.ie_query,
                    version_ids=self.version_ids,
                    search_descriptions=self.search_desc,
                )
            else:
                messages = self.db.get_messages_list(self.version_ids)
            self.results_ready.emit(messages, self.request_id)
        except Exception as e:
            logging.error(f"Error querying messages in background: {e}")
            self.results_ready.emit([], self.request_id)


class NASTab(QWidget):
    """Main UI tab for Protocol Evolution Matrix, Release Ingestion, and Definition Inspection."""

    log_msg = pyqtSignal(str, int)

    def __init__(self, nas_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.nas_db_path = Path(nas_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = self.nas_db_path.parent / "nas_config.json"

        self._reverse_lookup_worker: Optional[ReverseLookupWorker] = None
        self._reverse_lookup_request_id: int = 0
        self._msg_query_worker: Optional[NASMessageListWorker] = None
        self._msg_query_req_id: int = 0

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
        self._current_message_spec: Optional[str] = None
        self._loading_config: bool = False
        self._initialized_filters: bool = False

        self._msg_search_timer = QTimer(self)
        self._msg_search_timer.setSingleShot(True)
        self._msg_search_timer.setInterval(200)
        self._msg_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._ie_search_timer = QTimer(self)
        self._ie_search_timer.setSingleShot(True)
        self._ie_search_timer.setInterval(200)
        self._ie_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._setup_ui()
        self.refresh_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Toolbar
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.fetch_btn = QPushButton("📥 Import from Specs DB")
        self.fetch_btn.setCursor(Qt.PointingHandCursor)
        self.fetch_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.fetch_btn.clicked.connect(self._on_fetch_from_specs_db_clicked)
        toolbar.addWidget(self.fetch_btn)

        self.import_file_btn = QPushButton("📁 Import Local .docx")
        self.import_file_btn.setCursor(Qt.PointingHandCursor)
        self.import_file_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.import_file_btn.clicked.connect(self._on_import_local_file_clicked)
        toolbar.addWidget(self.import_file_btn)

        self.clear_ver_btn = QPushButton("🗑️ Clear Version")
        self.clear_ver_btn.setCursor(Qt.PointingHandCursor)
        self.clear_ver_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.clear_ver_btn.clicked.connect(self._on_clear_version_clicked)
        toolbar.addWidget(self.clear_ver_btn)

        self.wipe_db_btn = QPushButton("⚠️ Wipe DB")
        self.wipe_db_btn.setCursor(Qt.PointingHandCursor)
        self.wipe_db_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        toolbar.addWidget(self.wipe_db_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        main_splitter = QSplitter(Qt.Horizontal)

        # Left Panel (Tree + Message List)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        ver_group = QGroupBox("Specification Versions && Releases")
        ver_layout = QVBoxLayout(ver_group)
        self.version_tree = NASVersionTreeWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.structure_changed.connect(self._save_config)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        ver_layout.addWidget(self.version_tree)
        left_layout.addWidget(ver_group)

        msg_group = QGroupBox("Protocol Messages && SIBs")
        msg_layout = QVBoxLayout(msg_group)

        self.msg_search = QLineEdit()
        self.msg_search.setPlaceholderText("Filter message/SIB name (e.g. SIB1, Reconfig)...")
        self.msg_search.setClearButtonEnabled(True)
        self.msg_search.textChanged.connect(lambda: self._msg_search_timer.start())
        msg_layout.addWidget(self.msg_search)

        ie_search_bar_layout = QHBoxLayout()
        ie_search_bar_layout.setContentsMargins(0, 0, 0, 0)
        ie_search_bar_layout.setSpacing(4)

        self.ie_search = QLineEdit()
        self.ie_search.setPlaceholderText("Filter by IE / Field (e.g. RadioBearer, CellGroup)...")
        self.ie_search.setClearButtonEnabled(True)
        self.ie_search.textChanged.connect(lambda: self._ie_search_timer.start())
        self.ie_search.setContextMenuPolicy(Qt.CustomContextMenu)
        self.ie_search.customContextMenuRequested.connect(self._on_ie_search_context_menu)
        ie_search_bar_layout.addWidget(self.ie_search)

        self.deep_search_btn = QPushButton("📖 Desc")
        self.deep_search_btn.setCheckable(True)
        self.deep_search_btn.setToolTip("Extended Search: Also search inside Clause 9 / ASN.1 field descriptions")
        self.deep_search_btn.setFixedHeight(26)
        self.deep_search_btn.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: bold;
                color: #64748B;
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 6px;
            }
            QPushButton:hover {
                color: #1E5C99;
                background-color: #EBF3FC;
                border-color: #1E5C99;
            }
            QPushButton:checked {
                color: #1E5C99;
                background-color: #EBF3FC;
                border: 1.5px solid #1E5C99;
            }
        """)
        self.deep_search_btn.toggled.connect(self._on_deep_search_toggled)
        ie_search_bar_layout.addWidget(self.deep_search_btn)
        msg_layout.addLayout(ie_search_bar_layout)

        self.msg_list = QListWidget()
        self.msg_list.itemClicked.connect(self._on_message_clicked)
        msg_layout.addWidget(self.msg_list)
        left_layout.addWidget(msg_group)
        main_splitter.addWidget(left_widget)

        # Right Panel (Matrix Table + Inspector)
        right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)

        matrix_header_layout = QHBoxLayout()
        self.matrix_title = QLabel("Select a Message to View Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        matrix_header_layout.addWidget(self.matrix_title)
        matrix_header_layout.addStretch()

        # Interface Filter Dropdown styled using the application-wide COMBOBOX_STYLE_TOOLBAR
        self.interface_combo = QComboBox()
        list_view = QListView(self.interface_combo)
        list_view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        list_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.interface_combo.setView(list_view)
        self.interface_combo.setMaxVisibleItems(25)
        self.interface_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.interface_combo.setToolTip("Filter fields by reference point / interface applicability")
        self.interface_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.interface_combo.setVisible(False)
        self.interface_combo.currentIndexChanged.connect(lambda _: self._rebuild_matrix_model())
        matrix_header_layout.addWidget(self.interface_combo)
        matrix_layout.addLayout(matrix_header_layout)

        self.matrix_table = QTableView()
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QTableView.SelectRows)
        self.matrix_table.setSelectionMode(QTableView.SingleSelection)
        self.matrix_table.verticalHeader().setDefaultSectionSize(24)
        self.matrix_table.verticalHeader().setVisible(False)

        h_header = self.matrix_table.horizontalHeader()
        h_header.setHighlightSections(False)
        h_header.setStretchLastSection(False)

        self.matrix_table.clicked.connect(self._on_table_cell_clicked)
        self.matrix_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.matrix_table.customContextMenuRequested.connect(self._on_matrix_context_menu)
        matrix_layout.addWidget(self.matrix_table)
        right_splitter.addWidget(matrix_widget)

        self.inspector = NASInspectorWidget(db=self.db)
        self.inspector.jump_to_message_requested.connect(self._jump_to_message)
        self.inspector.filter_by_ie_requested.connect(lambda term: self.ie_search.setText(term))
        self.inspector.import_spec_requested.connect(self._on_import_cross_ref_spec_requested)
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([380, 320])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([300, 650])
        layout.addWidget(main_splitter)

    # -------------------------------------------------------------------------
    # Configuration Persistence
    # -------------------------------------------------------------------------

    def _load_config(self) -> Dict[str, Any]:
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Could not load Protocol config from {self.config_path}: {e}")
        return {}

    def _save_config(self):
        if self._loading_config or self.version_tree._updating_checks:
            return

        try:
            config_data = {
                "checked_versions": self.version_tree.get_checked_versions_info(),
                "msg_filter": self.msg_search.text(),
                "ie_filter": self.ie_search.text(),
                "search_descriptions": self.deep_search_btn.isChecked(),
                "selected_message": self.current_selected_message_name or "",
                "collapsed_specs": self.version_tree.get_collapsed_spec_numbers(),
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.warning(f"Could not save Protocol config to {self.config_path}: {e}")

    # -------------------------------------------------------------------------
    # Search Handlers
    # -------------------------------------------------------------------------

    def _on_deep_search_toggled(self, checked: bool):
        placeholder = (
            "Filter by Field or Description (e.g. emergency)..."
            if checked
            else "Filter by IE / Field (e.g. RadioBearer, CellGroup)..."
        )
        self.ie_search.setPlaceholderText(placeholder)
        self.deep_search_btn.setText("📖 Desc: ON" if checked else "📖 Desc")
        self._save_config()
        self._populate_messages()

    def _on_ie_search_context_menu(self, pos):
        menu = self.ie_search.createStandardContextMenu()
        menu.addSeparator()

        act_toggle = QAction("📖 Include Descriptions in Search", self)
        act_toggle.setCheckable(True)
        act_toggle.setChecked(self.deep_search_btn.isChecked())
        act_toggle.toggled.connect(self.deep_search_btn.setChecked)
        menu.addAction(act_toggle)
        menu.exec_(self.ie_search.mapToGlobal(pos))

    # -------------------------------------------------------------------------
    # Version Tree Management
    # -------------------------------------------------------------------------

    def refresh_versions(self):
        versions = self.db.get_imported_versions()
        saved_config = self._load_config()

        self.version_tree.populate(
            versions=versions,
            saved_checked=saved_config.get("checked_versions"),
            saved_collapsed=set(saved_config.get("collapsed_specs", [])),
        )

        if not self._initialized_filters:
            self._loading_config = True
            if "msg_filter" in saved_config:
                self.msg_search.setText(saved_config["msg_filter"])
            if "ie_filter" in saved_config:
                self.ie_search.setText(saved_config["ie_filter"])
            if "search_descriptions" in saved_config:
                self.deep_search_btn.setChecked(bool(saved_config["search_descriptions"]))
            if "selected_message" in saved_config and saved_config["selected_message"]:
                self.current_selected_message_name = saved_config["selected_message"]
            self._loading_config = False
            self._initialized_filters = True

        self._populate_messages()

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._save_config()
        self._populate_messages()

        current_msg_item = self.msg_list.currentItem()
        if current_msg_item:
            self._on_message_clicked(current_msg_item)
        else:
            self.matrix_table.setModel(None)

    def _delete_single_version(self, spec_number: str, version: str):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete all stored data for TS {spec_number} v{version}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            if self.db.clear_version(spec_number, version):
                self.current_selected_message_name = None
                self.refresh_versions()
                self._save_config()
                self.log_msg.emit(f"🗑️ Deleted TS {spec_number} v{version} from database.", logging.INFO)
            else:
                QMessageBox.warning(self, "Error", f"Failed to delete TS {spec_number} v{version}.")

    def _delete_spec_group(self, spec_number: str, child_count: int):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete all {child_count} version(s) of TS {spec_number} from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for item in self.version_tree.get_checked_versions_info():
                if item.get("spec_number") == spec_number and item.get("version"):
                    self.db.clear_version(spec_number, item["version"])
            self.current_selected_message_name = None
            self.refresh_versions()
            self._save_config()
            self.log_msg.emit(f"🗑️ Deleted all versions of TS {spec_number} from database.", logging.INFO)

    # -------------------------------------------------------------------------
    # Message & Matrix Handlers
    # -------------------------------------------------------------------------

    def _on_search_timer_timeout(self):
        self._save_config()
        self._populate_messages()

    def _populate_messages(self):
        if not self.selected_version_ids:
            self.msg_list.clear()
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
            self.inspector.clear_display()
            self.current_selected_message_name = None
            return

        ie_query = self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()

        self._msg_query_req_id += 1
        req_id = self._msg_query_req_id

        if self._msg_query_worker and self._msg_query_worker.isRunning():
            self._msg_query_worker.terminate()
            self._msg_query_worker.wait()

        self._msg_query_worker = NASMessageListWorker(
            db=self.db,
            ie_query=ie_query,
            version_ids=self.selected_version_ids,
            search_desc=search_desc,
            request_id=req_id,
        )
        self._msg_query_worker.results_ready.connect(self._on_message_list_ready)
        self._msg_query_worker.start()

    def _on_message_list_ready(self, messages: List[Dict[str, Any]], req_id: int):
        if req_id != self._msg_query_req_id:
            return

        msg_query = self.msg_search.text().strip().lower()
        target_msg_name = self.current_selected_message_name

        self.msg_list.blockSignals(True)
        self.msg_list.clear()

        target_item = None
        for m in messages:
            msg_name = m["message_name"]
            spec_num = m.get("spec_number", "")
            item_text = f"{msg_name} ({m['clause']})"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, msg_name)
            if spec_num:
                item.setToolTip(f"Specification: TS {spec_num}\nClause: {m['clause']}")

            if msg_query and msg_query not in item_text.lower():
                item.setHidden(True)

            self.msg_list.addItem(item)
            if target_msg_name and msg_name == target_msg_name:
                target_item = item

        self.msg_list.blockSignals(False)

        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        elif self.msg_list.count() > 0 and not target_msg_name:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
        else:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
            self.inspector.clear_display()

    def _on_message_clicked(self, item: QListWidgetItem):
        msg_name = item.data(Qt.UserRole)
        self.current_selected_message_name = msg_name
        self._save_config()

        ie_query = self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()
        desc_label = " (incl. Desc)" if search_desc and ie_query else ""
        title_suffix = f" (Filtered by IE{desc_label}: '{ie_query}')" if ie_query else ""

        need_descriptions = search_desc and bool(ie_query)

        df = self.db.get_message_evolution_df(
            message_name=msg_name,
            version_ids=self.selected_version_ids,
            include_descriptions=need_descriptions,
        )

        specs = []
        if not df.empty and "spec_number" in df.columns:
            unique_specs = sorted(df["spec_number"].dropna().unique())
            specs = [f"TS {s}" if not str(s).startswith("TS") else str(s) for s in unique_specs]
            self._current_message_spec = str(unique_specs[0]) if len(unique_specs) == 1 else None
        else:
            self._current_message_spec = None

        spec_prefix = f" ({', '.join(specs)})" if specs else ""
        self.matrix_title.setText(f"Message{spec_prefix}: {msg_name}{title_suffix}")

        has_appl = "applicability" in df.columns and df["applicability"].fillna("").astype(str).str.strip().ne("").any()
        self.interface_combo.blockSignals(True)
        self.interface_combo.setVisible(has_appl)
        if has_appl:
            current_choice = self.interface_combo.currentText()
            unique_ifaces = set()
            for raw_val in df["applicability"].dropna():
                for part in str(raw_val).split(","):
                    token = part.strip()
                    if token and token.upper() not in ("ALL INTERFACES", "ALL"):
                        unique_ifaces.add(token)

            self.interface_combo.clear()
            self.interface_combo.addItem("All Interfaces")
            for iface in sorted(unique_ifaces):
                self.interface_combo.addItem(iface)

            self.interface_combo.setMaxVisibleItems(max(20, len(unique_ifaces) + 2))

            idx = self.interface_combo.findText(current_choice)
            self.interface_combo.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            self.interface_combo.setCurrentIndex(0)
        self.interface_combo.blockSignals(False)

        self._current_matrix_df = df
        self._rebuild_matrix_model()

    def _rebuild_matrix_model(self):
        if not hasattr(self, "_current_matrix_df") or self._current_matrix_df.empty:
            return

        df = self._current_matrix_df
        ie_query = self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()
        active_if = self.interface_combo.currentText() if self.interface_combo.isVisible() else None

        model = NASEvolutionMatrixModel(
            df,
            ie_filter=ie_query,
            search_descriptions=(search_desc and bool(ie_query)),
            interface_filter=active_if,
        )
        self.matrix_table.setModel(model)

        h_header = self.matrix_table.horizontalHeader()
        if model.columnCount() > 0:
            h_header.setSectionResizeMode(0, QHeaderView.Interactive)
            self.matrix_table.setColumnWidth(0, 75)

            h_header.setSectionResizeMode(1, QHeaderView.Stretch)
            h_header.setSectionResizeMode(2, QHeaderView.Interactive)
            self.matrix_table.setColumnWidth(2, 280)

            for col in range(3, model.columnCount()):
                h_header.setSectionResizeMode(col, QHeaderView.Interactive)
                self.matrix_table.setColumnWidth(col, 130)

    def _on_table_cell_clicked(self, index):
        model = self.matrix_table.model()
        if not model:
            return

        row, col = index.row(), index.column()
        ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
        type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

        field_desc = ""
        if hasattr(model, "_pivot_df") and not model._pivot_df.empty and row < len(model._pivot_df):
            row_data = model._pivot_df.iloc[row]
            field_desc = str(row_data.get("field_description", "") or "")

        clean_type = re.sub(r"^SetupRelease\s*\{\s*([A-Za-z0-9\-]+)\s*\}", r"\1", type_ref)
        clean_type = re.sub(r"^SEQUENCE\s*(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\s+([A-Za-z0-9\-]+)", r"\1", clean_type)
        clean_type = re.sub(r"^OCTET STRING\s*\(\s*CONTAINING\s+([A-Za-z0-9\-]+)\s*\)", r"\1", clean_type)
        clean_type = re.sub(r"[\(\{\[].*$", "", clean_type).strip()

        if clean_type.upper() in ("ENUMERATED", "INTEGER", "BOOLEAN", "BIT STRING", "OCTET STRING", "NULL"):
            clean_type = ""

        spec_num = getattr(self, "_current_message_spec", None)
        target_version_hint = None

        if col >= 3:
            col_name = str(model._get_visible_column_name(col))
            match_spec = RE_SPEC_FROM_COL.search(col_name)
            if match_spec:
                spec_num = match_spec.group(0)
            target_version_hint = col_name

        match_clause = RE_CLAUSE_FROM_REF.search(type_ref)
        clause = match_clause.group(1).strip() if match_clause else ie_name

        defs = []
        if clean_type:
            defs = self.db.get_ie_definitions_by_clause(
                clause=clean_type,
                alt_name=ie_name,
                spec_number=spec_num,
                version_ids=self.selected_version_ids,
            )

        if not defs and clause and not any(
                clause.upper().startswith(k) for k in ("ENUMERATED", "INTEGER", "BOOLEAN", "BIT STRING")
        ):
            defs = self.db.get_ie_definitions_by_clause(
                clause=clause,
                alt_name=ie_name,
                spec_number=spec_num,
                version_ids=self.selected_version_ids,
            )

        self.inspector.display_definitions(
            clause=clause,
            ie_name=clean_type or ie_name,
            spec_number=spec_num,
            defs=defs,
            containing_msgs=[],
            target_version_hint=target_version_hint,
            fallback_type_ref=type_ref,
            field_description=field_desc,
        )

        if defs:
            resolved_name = defs[0]["ie_name"]
            search_target_name = clean_type or resolved_name

            if self._reverse_lookup_worker and self._reverse_lookup_worker.isRunning():
                self._reverse_lookup_worker.terminate()
                self._reverse_lookup_worker.wait()

            self._reverse_lookup_request_id += 1
            current_req_id = self._reverse_lookup_request_id

            self._reverse_lookup_worker = ReverseLookupWorker(
                db=self.db,
                clause=clause,
                ie_name=search_target_name,
                spec_number=spec_num,
                version_ids=self.selected_version_ids,
                request_id=current_req_id,
            )
            self._reverse_lookup_worker.results_ready.connect(self._on_reverse_lookup_finished)
            self._reverse_lookup_worker.start()

    def _on_reverse_lookup_finished(self, containing_msgs: List[Dict[str, Any]], req_id: int):
        if req_id == self._reverse_lookup_request_id:
            self.inspector.set_containing_messages(containing_msgs)

    def _on_matrix_context_menu(self, pos):
        model = self.matrix_table.model()
        if not model:
            return

        index = self.matrix_table.indexAt(pos)
        if not index.isValid():
            return

        row = index.row()
        ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
        type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

        match = RE_CLAUSE_FROM_REF.search(type_ref)
        clause = match.group(1).strip() if match else ie_name
        spec_num = getattr(self, "_current_message_spec", None)

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                font-size: 11px;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                padding: 4px;
            }
            QMenu::item {
                padding: 4px 15px;
            }
            QMenu::item:selected {
                background-color: #EBF3FC;
                color: #1E5C99;
            }
        """)

        filter_term = ie_name if ie_name else clause
        act_filter = QAction(f"🔍 Filter message list for '{filter_term}'", self)
        act_filter.triggered.connect(lambda: self.ie_search.setText(filter_term))
        menu.addAction(act_filter)

        act_inspect = QAction(f"📖 Inspect Definition ({filter_term})", self)
        act_inspect.triggered.connect(lambda: self._on_table_cell_clicked(index))
        menu.addAction(act_inspect)

        if clause or ie_name:
            containing_msgs = self.db.get_messages_using_ie(
                clause=clause,
                ie_name=ie_name,
                spec_number=spec_num,
                version_ids=self.selected_version_ids,
            )
            sub_menu = menu.addMenu(f"📂 Open message using this IE ({len(containing_msgs)} found)...")
            for m in containing_msgs:
                mn = m["message_name"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [TS {m_spec}]" if m_spec and not spec_num else ""
                sub_act = QAction(f"{mn} ({m['clause']}){spec_tag}", self)
                sub_act.triggered.connect(lambda checked, target=mn: self._jump_to_message(target))
                sub_menu.addAction(sub_act)

        menu.exec_(self.matrix_table.viewport().mapToGlobal(pos))

    def _jump_to_message(self, message_name: str):
        self.current_selected_message_name = message_name
        self._save_config()

        if self.msg_search.text().strip() or self.ie_search.text().strip():
            visible_names = [
                self.msg_list.item(i).data(Qt.UserRole)
                for i in range(self.msg_list.count())
                if not self.msg_list.item(i).isHidden()
            ]
            if message_name not in visible_names:
                self.msg_search.clear()
                self.ie_search.clear()

        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            if item.data(Qt.UserRole) == message_name:
                item.setHidden(False)
                self.msg_list.setCurrentItem(item)
                self.msg_list.scrollToItem(item)
                self._on_message_clicked(item)
                break

    # -------------------------------------------------------------------------
    # Ingestion Actions
    # -------------------------------------------------------------------------

    def _on_fetch_from_specs_db_clicked(self):
        if not self.specs_db:
            QMessageBox.warning(
                self,
                "Specs DB Unavailable",
                "The 3GPP Specifications database (3gpp_data.db) is not configured or reachable.",
            )
            return

        dialog = NASVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)
        if dialog.exec_() == QDialog.Accepted and dialog.selected_files_info:
            self._start_batch_ingestion(dialog.selected_files_info)

    def _on_import_local_file_clicked(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select 3GPP Specification(s) (.docx)",
            "",
            "Word Files (*.docx)",
        )
        if not file_paths:
            return

        grouped_tasks: Dict[str, List[Path]] = {}
        for fp in file_paths:
            p = Path(fp)
            base_key = re.sub(r"_\d+_.*$", "", p.stem)
            grouped_tasks.setdefault(base_key, []).append(p)

        tasks = []
        for base_key, paths in grouped_tasks.items():
            parser_temp = ProtocolDocxDispatcher(paths)
            spec_num = parser_temp.extract_spec_number()
            version = parser_temp.extract_version_from_filename()
            tasks.append({
                "spec_number": spec_num,
                "version": version,
                "filename": f"{base_key}.docx" if len(paths) == 1 else f"{base_key}.zip",
                "file_url": "",
                "local_docx_paths": paths,
            })

        self._start_batch_ingestion(tasks)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.import_file_btn.setEnabled(False)

        self.thread = NASFetchAndImportThread(
            nas_db_path=self.nas_db_path,
            tasks=tasks,
            cache_dir=self.cache_dir,
        )
        self.thread.progress.connect(self._on_import_progress)
        self.thread.finished_success.connect(self._on_import_success)
        self.thread.error.connect(self._on_import_error)
        self.thread.start()

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        self.log_msg.emit(msg, logging.INFO)

    def _on_import_success(self, spec_count: int, msg_count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        self.log_msg.emit(
            f"✅ Successfully ingested {spec_count} specification(s) ({msg_count} total messages/PDUs).",
            logging.INFO,
        )
        self.refresh_versions()

    def _on_clear_version_clicked(self):
        checked_versions = self.version_tree.get_checked_versions_info()
        if not checked_versions:
            QMessageBox.warning(self, "Select Version", "Please check at least one specific version to clear.")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete {len(checked_versions)} checked specification version(s)?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for item in checked_versions:
                self.db.clear_version(item["spec_number"], item["version"])
            self.current_selected_message_name = None
            self.refresh_versions()
            self._save_config()

    def _on_wipe_db_clicked(self):
        reply = QMessageBox.critical(
            self,
            "Confirm Wipe",
            "This will delete ALL imported specifications and tables. Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.db.wipe_database()
            self.current_selected_message_name = None
            if self.config_path.exists():
                try:
                    self.config_path.unlink()
                except Exception:
                    pass
            self.refresh_versions()
            self.msg_list.clear()
            self.matrix_table.setModel(None)
            self.inspector.clear_display()
            self.log_msg.emit("🧹 Protocol Database wiped.", logging.INFO)

    def _on_import_cross_ref_spec_requested(self, spec_number: str, major_version: int):
        if not self.specs_db:
            QMessageBox.warning(
                self,
                "Specs DB Unavailable",
                "The 3GPP Specifications database (3gpp_data.db) is not configured or reachable.",
            )
            return

        dialog = NASVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)

        for idx in range(dialog.spec_combo.count()):
            if dialog.spec_combo.itemData(idx) == spec_number:
                dialog.spec_combo.setCurrentIndex(idx)
                break

        if major_version > 0:
            dialog.table.clearSelection()
            for row in range(dialog.table.rowCount()):
                item = dialog.table.item(row, 0)
                if item:
                    data = item.data(Qt.UserRole) or {}
                    ver_tuple = parse_version_tuple(data.get("version", ""))
                    if ver_tuple and ver_tuple[0] == major_version:
                        dialog.table.selectRow(row)
                        dialog.table.scrollToItem(item)
                        break

        if dialog.exec_() == QDialog.Accepted and dialog.selected_files_info:
            self._start_batch_ingestion(dialog.selected_files_info)


class ReverseLookupWorker(QThread):
    """Background worker to fetch message references without freezing the UI."""

    results_ready = pyqtSignal(list, int)

    def __init__(
            self,
            db: NASDatabase,
            clause: str,
            ie_name: str,
            spec_number: Optional[str],
            version_ids: List[int],
            request_id: int,
    ):
        super().__init__()
        self.db = db
        self.clause = clause
        self.ie_name = ie_name
        self.spec_number = spec_number
        self.version_ids = version_ids
        self.request_id = request_id

    def run(self):
        try:
            results = self.db.get_messages_using_ie(
                clause=self.clause,
                ie_name=self.ie_name,
                spec_number=self.spec_number,
                version_ids=self.version_ids,
            )
            self.results_ready.emit(results, self.request_id)
        except Exception:
            self.results_ready.emit([], self.request_id)