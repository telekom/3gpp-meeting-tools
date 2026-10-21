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
from core.utils.paths import get_project_root, get_specs_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import NASFetchAndImportThread
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.ui.nas_components import NASInspectorWidget, NASVersionTreeWidget
from modules.nas.ui.nas_dialogs import NASVersionSelectDialog
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase

logger = logging.getLogger(__name__)

RE_SPEC_FROM_COL = re.compile(r"(?:24|25|29|36|37|38)\.[0-9]{3}")
RE_CLAUSE_FROM_REF = re.compile(r"((?:9|8|7|6|5|4|D\.6)(?:\.[0-9A-Za-z]+)+)")

TOGGLE_BUTTON_STYLE = """
    QPushButton {
        font-size: 11px;
        font-weight: bold;
        color: #64748B;
        background-color: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 3px 8px;
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
    QPushButton:checked:hover {
        color: #1E5C99;
        background-color: #DCEBFA;
        border: 1.5px solid #1E5C99;
    }
"""



# Worker threads extracted to nas_workers.py
from modules.nas.ui.nas_workers import NASMessageListWorker, ReverseLookupWorker
from modules.nas.ui.nas_matrix_controller import handle_table_cell_click, show_matrix_context_menu


class NASTab(QWidget):
    """Main UI tab for Protocol Evolution Matrix, Release Ingestion, and Definition Inspection."""

    log_msg = pyqtSignal(str, int)

    def __init__(self, nas_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.nas_db_path = Path(nas_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = get_project_root() / "config" / "nas_config.json"

        self._reverse_lookup_worker: Optional[ReverseLookupWorker] = None
        self._reverse_lookup_request_id: int = 0
        self._msg_query_workers: Dict[int, NASMessageListWorker] = {}
        self._msg_query_req_id: int = 0
        self._reverse_lookup_workers: Dict[int, ReverseLookupWorker] = {}

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = get_specs_root()

        self.db = NASDatabase(self.nas_db_path)
        self.specs_db = (
            SpecsDatabase(self.specs_db_path)
            if self.specs_db_path and self.specs_db_path.exists()
            else None
        )

        self.selected_version_ids: List[int] = []
        self._browse_mode: str = "messages"  # "messages" or "ies"
        self.current_selected_message_name: Optional[str] = None
        self.current_selected_ie_key: Optional[str] = None
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

        # Left Panel (Tree + Message / IE List)
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

        self.msg_group = QGroupBox("Protocol Messages && Information Elements")
        msg_layout = QVBoxLayout(self.msg_group)

        # Mode Toggle Bar: [📨 Messages | 🧩 All IEs / Types]
        mode_bar_layout = QHBoxLayout()
        mode_bar_layout.setContentsMargins(0, 0, 0, 2)
        mode_bar_layout.setSpacing(4)

        self.mode_messages_btn = QPushButton("📨 Messages")
        self.mode_messages_btn.setCheckable(True)
        self.mode_messages_btn.setChecked(True)
        self.mode_messages_btn.setCursor(Qt.PointingHandCursor)
        self.mode_messages_btn.setFixedHeight(26)
        self.mode_messages_btn.setToolTip("Browse Protocol Messages & SIBs (Clause 8 / ASN.1 PDUs)")
        self.mode_messages_btn.setStyleSheet(TOGGLE_BUTTON_STYLE)
        self.mode_messages_btn.clicked.connect(lambda: self._set_browse_mode("messages"))
        mode_bar_layout.addWidget(self.mode_messages_btn)

        self.mode_ies_btn = QPushButton("🧩 All IEs / Types")
        self.mode_ies_btn.setCheckable(True)
        self.mode_ies_btn.setChecked(False)
        self.mode_ies_btn.setCursor(Qt.PointingHandCursor)
        self.mode_ies_btn.setFixedHeight(26)
        self.mode_ies_btn.setToolTip(
            "Browse all parsed Information Elements & Types directly (including standalone & nested Clause 9 containers)"
        )
        self.mode_ies_btn.setStyleSheet(TOGGLE_BUTTON_STYLE)
        self.mode_ies_btn.clicked.connect(lambda: self._set_browse_mode("ies"))
        mode_bar_layout.addWidget(self.mode_ies_btn)

        msg_layout.addLayout(mode_bar_layout)

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
        self.deep_search_btn.setStyleSheet(TOGGLE_BUTTON_STYLE)
        self.deep_search_btn.toggled.connect(self._on_deep_search_toggled)
        ie_search_bar_layout.addWidget(self.deep_search_btn)
        msg_layout.addLayout(ie_search_bar_layout)

        self.msg_list = QListWidget()
        self.msg_list.itemClicked.connect(self._on_message_clicked)
        msg_layout.addWidget(self.msg_list)
        left_layout.addWidget(self.msg_group)
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
        self.inspector.filter_by_ie_requested.connect(self._on_filter_by_ie_requested)
        self.inspector.import_spec_requested.connect(self._on_import_cross_ref_spec_requested)
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([380, 320])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([300, 650])
        layout.addWidget(main_splitter)

    # -------------------------------------------------------------------------
    # Browse Mode & Configuration Persistence
    # -------------------------------------------------------------------------

    def _set_browse_mode(self, mode: str, trigger_populate: bool = True):
        mode = "ies" if mode == "ies" else "messages"
        self._browse_mode = mode

        self.mode_messages_btn.blockSignals(True)
        self.mode_ies_btn.blockSignals(True)
        self.mode_messages_btn.setChecked(mode == "messages")
        self.mode_ies_btn.setChecked(mode == "ies")
        self.mode_messages_btn.blockSignals(False)
        self.mode_ies_btn.blockSignals(False)

        self._update_search_placeholders()

        if trigger_populate:
            self._save_config()
            self._populate_messages()

    def _update_search_placeholders(self):
        checked = self.deep_search_btn.isChecked()
        if self._browse_mode == "ies":
            self.msg_search.setPlaceholderText("Filter IE / Type name or Clause (e.g. UE parameters, 9.11.3)...")
            self.ie_search.setPlaceholderText(
                "Filter IE by keyword or Description..."
                if checked
                else "Additional IE / Clause keyword filter..."
            )
        else:
            self.msg_search.setPlaceholderText("Filter message/SIB name (e.g. SIB1, Reconfig)...")
            self.ie_search.setPlaceholderText(
                "Filter by Field or Description (e.g. emergency)..."
                if checked
                else "Filter by IE / Field (e.g. RadioBearer, CellGroup)..."
            )

    def _on_filter_by_ie_requested(self, term: str):
        if self._browse_mode != "messages":
            self._set_browse_mode("messages", trigger_populate=False)
        self.ie_search.setText(term)

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
                "browse_mode": self._browse_mode,
                "msg_filter": self.msg_search.text(),
                "ie_filter": self.ie_search.text(),
                "search_descriptions": self.deep_search_btn.isChecked(),
                "selected_message": self.current_selected_message_name or "",
                "selected_ie": self.current_selected_ie_key or "",
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
        self._update_search_placeholders()
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
            if "browse_mode" in saved_config:
                self._set_browse_mode(str(saved_config["browse_mode"]), trigger_populate=False)
            if "msg_filter" in saved_config:
                self.msg_search.setText(saved_config["msg_filter"])
            if "ie_filter" in saved_config:
                self.ie_search.setText(saved_config["ie_filter"])
            if "search_descriptions" in saved_config:
                self.deep_search_btn.setChecked(bool(saved_config["search_descriptions"]))
            if "selected_message" in saved_config and saved_config["selected_message"]:
                self.current_selected_message_name = saved_config["selected_message"]
            if "selected_ie" in saved_config and saved_config["selected_ie"]:
                self.current_selected_ie_key = saved_config["selected_ie"]
            self._loading_config = False
            self._initialized_filters = True

        # NASVersionTreeWidget.populate() emits selection_changed once after the
        # tree has been rebuilt. That signal already invokes
        # _on_version_selection_changed() -> _populate_messages(). Do not call
        # _populate_messages() a second time here: doing so used to start a
        # worker and immediately terminate it during NASTab construction.

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
                logger.log(logging.INFO, f"🗑️ Deleted TS {spec_number} v{version} from database.")
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
            for item in self.version_tree.get_all_versions_for_spec(spec_number):
                if item.get("version"):
                    self.db.clear_version(spec_number, item["version"])
            self.current_selected_message_name = None
            self.refresh_versions()
            self._save_config()
            logger.log(logging.INFO, f"🗑️ Deleted all versions of TS {spec_number} from database.")

    # -------------------------------------------------------------------------
    # Message / IE List & Matrix Handlers
    # -------------------------------------------------------------------------

    def _on_search_timer_timeout(self):
        self._save_config()
        self._populate_messages()

    def _populate_messages(self):
        if not self.selected_version_ids:
            self.msg_list.clear()
            self.matrix_table.setModel(None)
            default_prompt = (
                "Select an Information Element to View Evolution & Definition"
                if self._browse_mode == "ies"
                else "Select a Message to View Evolution Matrix"
            )
            self.matrix_title.setText(default_prompt)
            self.inspector.clear_display()
            return

        ie_query = self.ie_search.text().strip()
        if self._browse_mode == "ies":
            # In All IEs mode, combine both search boxes so either input filters IE definitions
            top_query = self.msg_search.text().strip()
            combined_query = f"{top_query} {ie_query}".strip()
        else:
            combined_query = ie_query

        search_desc = self.deep_search_btn.isChecked()

        self._msg_query_req_id += 1
        req_id = self._msg_query_req_id

        # Do not terminate an in-flight QThread. QThread.terminate() can stop a
        # worker while native/SQLite code holds internal locks, which can leave
        # Qt/native state inconsistent. Request IDs already make stale results
        # harmless, so let older workers finish naturally.
        worker = NASMessageListWorker(
            db=self.db,
            ie_query=combined_query,
            version_ids=list(self.selected_version_ids),
            search_desc=search_desc,
            request_id=req_id,
            browse_mode=self._browse_mode,
        )
        self._msg_query_workers[req_id] = worker
        worker.results_ready.connect(self._on_message_list_ready)
        worker.finished.connect(lambda rid=req_id: self._release_msg_query_worker(rid))
        worker.start()

    def _release_msg_query_worker(self, req_id: int):
        """Release a completed message-query worker without touching running workers."""
        worker = self._msg_query_workers.pop(req_id, None)
        if worker is not None:
            worker.deleteLater()

    def _on_message_list_ready(self, items: List[Dict[str, Any]], req_id: int):
        if req_id != self._msg_query_req_id:
            return

        msg_query = self.msg_search.text().strip().lower()
        msg_tokens = [t for t in msg_query.split() if t]

        self.msg_list.blockSignals(True)
        self.msg_list.clear()

        target_item = None

        if self._browse_mode == "ies":
            target_key = self.current_selected_ie_key
            for d in items:
                ie_name = d.get("ie_name", "")
                clause = d.get("clause", "")
                spec_num = d.get("spec_number", "")
                item_text = f"{ie_name} ({clause})" if clause and clause != ie_name else ie_name
                ie_key = f"{ie_name}::{clause}"

                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, ie_key)
                item.setData(Qt.UserRole + 1, d)
                if spec_num:
                    item.setToolTip(f"Specification: TS {spec_num}\nClause / Type: {clause}")

                self.msg_list.addItem(item)

                if target_key and ie_key == target_key:
                    target_item = item
        else:
            target_msg_name = self.current_selected_message_name
            for m in items:
                msg_name = m["message_name"]
                spec_num = m.get("spec_number", "")
                item_text = f"{msg_name} ({m['clause']})"
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, msg_name)
                item.setData(Qt.UserRole + 1, m)
                if spec_num:
                    item.setToolTip(f"Specification: TS {spec_num}\nClause: {m['clause']}")

                self.msg_list.addItem(item)

                if msg_tokens and not all(tok in item_text.lower() for tok in msg_tokens):
                    item.setHidden(True)

                if target_msg_name and msg_name == target_msg_name and not item.isHidden():
                    target_item = item

        self.msg_list.blockSignals(False)

        default_prompt = (
            "Select an Information Element to View Evolution & Definition"
            if self._browse_mode == "ies"
            else "Select a Message to View Evolution Matrix"
        )

        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        elif self.msg_list.count() == 1 and (msg_tokens or self.ie_search.text().strip()):
            # Auto-select when search narrows down to a single exact item
            only_item = self.msg_list.item(0)
            if not only_item.isHidden():
                self.msg_list.setCurrentItem(only_item)
                self._on_message_clicked(only_item)
        elif self.msg_list.count() > 0:
            self.matrix_table.setModel(None)
            self.matrix_title.setText(default_prompt)
        else:
            self.matrix_table.setModel(None)
            self.matrix_title.setText(default_prompt)
            self.inspector.clear_display()

    def _on_message_clicked(self, item: QListWidgetItem):
        if self._browse_mode == "ies":
            ie_data = item.data(Qt.UserRole + 1) or {}
            ie_name = ie_data.get("ie_name", "")
            clause = ie_data.get("clause", "")
            self.current_selected_ie_key = item.data(Qt.UserRole)
            self._save_config()

            df = self.db.get_ie_definition_evolution_df(
                clause=clause,
                ie_name=ie_name,
                version_ids=self.selected_version_ids,
            )

            specs = []
            if not df.empty and "spec_number" in df.columns:
                unique_specs = sorted(df["spec_number"].dropna().unique())
                specs = [f"TS {s}" if not str(s).startswith("TS") else str(s) for s in unique_specs]
                self._current_message_spec = str(unique_specs[0]) if len(unique_specs) == 1 else None
            else:
                self._current_message_spec = None

            spec_prefix = f" ({', '.join(specs)})" if specs else ""
            clause_suffix = f" [Clause {clause}]" if clause and clause != ie_name else ""
            self.matrix_title.setText(f"Information Element{spec_prefix}: {ie_name}{clause_suffix}")
            self.interface_combo.setVisible(False)

            self._current_matrix_df = df
            self._rebuild_matrix_model(ignore_ie_filter=True)

            # Immediately display the full Clause 9 / ASN.1 definition in the Inspector below
            defs = self.db.get_ie_definitions_by_clause(
                clause=clause,
                alt_name=ie_name,
                spec_number=self._current_message_spec,
                version_ids=self.selected_version_ids,
            )
            self.inspector.display_definitions(
                clause=clause,
                ie_name=ie_name,
                spec_number=self._current_message_spec,
                defs=defs,
                containing_msgs=[],
                fallback_type_ref=f"{ie_name} {clause}".strip(),
            )
            if defs:
                self._start_reverse_lookup(
                    clause=clause,
                    ie_name=ie_name,
                    spec_num=self._current_message_spec,
                )
            return

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

        has_appl = bool(
            "applicability" in df.columns
            and df["applicability"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .any()
        )
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

    def _rebuild_matrix_model(self, ignore_ie_filter: bool = False):
        if not hasattr(self, "_current_matrix_df") or self._current_matrix_df.empty:
            return

        df = self._current_matrix_df
        ie_query = "" if (ignore_ie_filter or self._browse_mode == "ies") else self.ie_search.text().strip()
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

    def _start_reverse_lookup(self, clause: str, ie_name: str, spec_num: Optional[str]):
        self._reverse_lookup_request_id += 1
        current_req_id = self._reverse_lookup_request_id

        worker = ReverseLookupWorker(
            db=self.db,
            clause=clause,
            ie_name=ie_name,
            spec_number=spec_num,
            version_ids=list(self.selected_version_ids),
            request_id=current_req_id,
        )
        self._reverse_lookup_workers[current_req_id] = worker
        worker.results_ready.connect(self._on_reverse_lookup_finished)
        worker.finished.connect(
            lambda rid=current_req_id: self._release_reverse_lookup_worker(rid)
        )
        worker.start()

    def _on_table_cell_clicked(self, index):
        handle_table_cell_click(self, index)


    def _release_reverse_lookup_worker(self, req_id: int):
        """Release a completed reverse-lookup worker safely."""
        worker = self._reverse_lookup_workers.pop(req_id, None)
        if worker is not None:
            worker.deleteLater()

    def _on_reverse_lookup_finished(self, containing_msgs: List[Dict[str, Any]], req_id: int):
        if req_id == self._reverse_lookup_request_id:
            self.inspector.set_containing_messages(containing_msgs)

    def _on_matrix_context_menu(self, pos):
        show_matrix_context_menu(self, pos)


    def _jump_to_message(self, message_name: str):
        self.current_selected_message_name = message_name
        if self._browse_mode != "messages":
            self.msg_search.clear()
            self.ie_search.clear()
            self._set_browse_mode("messages", trigger_populate=True)
            return

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
                self._populate_messages()
                return

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

    def _on_import_error(self, err_msg: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        logger.log(logging.ERROR, f"❌ Import failed: {err_msg}")
        QMessageBox.critical(
            self,
            "Import Error",
            f"An error occurred during specification ingestion:\n{err_msg}",
        )

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        logger.log(logging.INFO, msg)

    def _on_import_success(self, spec_count: int, msg_count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        logger.info(
            f"✅ Successfully ingested {spec_count} specification(s) ({msg_count} total messages/PDUs)."
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
            logger.log(logging.INFO, "🧹 Protocol Database wiped.")

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
                    if ver_tuple and (ver_tuple[0][1] if isinstance(ver_tuple[0], tuple) else ver_tuple[0]) == major_version:
                        dialog.table.selectRow(row)
                        dialog.table.scrollToItem(item)
                        break

        if dialog.exec_() == QDialog.Accepted and dialog.selected_files_info:
            self._start_batch_ingestion(dialog.selected_files_info)

    def shutdown(self, wait_ms: int = 1500):
        """Cooperatively wait for short-lived NAS query workers during app shutdown."""
        self._msg_search_timer.stop()
        self._ie_search_timer.stop()

        workers = list(self._msg_query_workers.values()) + list(self._reverse_lookup_workers.values())
        ingest_thread = getattr(self, "thread", None)
        if isinstance(ingest_thread, QThread) and ingest_thread.isRunning():
            workers.append(ingest_thread)

        for worker in workers:
            if worker.isRunning():
                worker.requestInterruption()

        for worker in workers:
            if worker.isRunning():
                worker.wait(wait_ms)


