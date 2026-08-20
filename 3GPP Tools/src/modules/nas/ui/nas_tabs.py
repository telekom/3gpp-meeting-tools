import json
import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
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
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTableView,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_parser import NASDocxParser
from modules.nas.core.nas_threads import (
    NASFetchAndImportThread,
    find_cached_spec_file,
)
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase


def get_spec_title(spec_number: str) -> str:
    """Returns a descriptive title for 3GPP NAS and ASN.1 specifications."""
    spec_titles = {
        "38.331": "TS 38.331 (NR RRC)",
        "36.331": "TS 36.331 (LTE RRC)",
        "38.413": "TS 38.413 (NGAP)",
        "24.501": "TS 24.501 (5GS NAS)",
        "24.301": "TS 24.301 (EPS NAS)",
        "24.008": "TS 24.008 (Core Network)",
        "23.501": "TS 23.501 (5GS Architecture)",
        "23.502": "TS 23.502 (5GS Procedures)",
    }
    return spec_titles.get(spec_number, f"TS {spec_number}")


class NASVersionSelectDialog(QDialog):
    """Dialog allowing selection of NAS (24.501/24.301) and RRC/NGAP (38.331/36.331/38.413) versions."""

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
        self.selected_files_info: List[Dict[str, Any]] = []

        self.setWindowTitle("Select 3GPP Specification Version(s) to Ingest")
        self.resize(740, 480)
        self._setup_ui()
        self._load_available_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("Specification:"))

        self.spec_combo = QComboBox()
        self.spec_combo.addItem("TS 38.331 (NR RRC)", "38.331")
        self.spec_combo.addItem("TS 36.331 (LTE RRC)", "36.331")
        self.spec_combo.addItem("TS 38.413 (NGAP)", "38.413")
        self.spec_combo.addItem("TS 24.501 (5GS NAS)", "24.501")
        self.spec_combo.addItem("TS 24.301 (EPS NAS)", "24.301")
        self.spec_combo.currentIndexChanged.connect(self._load_available_versions)
        top_bar.addWidget(self.spec_combo)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        info_lbl = QLabel(
            "Select one or more versions from the specification archive (use Ctrl+Click or Shift+Click).\n"
            "Files not currently cached locally will be downloaded automatically from the 3GPP FTP server."
        )
        info_lbl.setStyleSheet("color: #555; padding-bottom: 4px;")
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Version", "Filename", "Local Cache", "DB Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("📥 Fetch && Ingest Selected")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _load_available_versions(self):
        spec_num = self.spec_combo.currentData()
        spec_files = self.specs_db.search_files(spec_number=spec_num)

        imported_entries = {
            (v["spec_number"], v["version"]) for v in self.nas_db.get_imported_versions()
        }

        spec_files = sorted(
            spec_files,
            key=lambda row: parse_version_tuple(row[5]),
            reverse=True,
        )

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(spec_files):
            _, s_num, _, _, filename, version, url = row_data
            self.table.insertRow(row_idx)

            v_item = QTableWidgetItem(f"v{version}")
            v_item.setData(
                Qt.UserRole,
                {
                    "spec_number": s_num,
                    "version": version,
                    "filename": filename,
                    "file_url": url,
                },
            )
            self.table.setItem(row_idx, 0, v_item)
            self.table.setItem(row_idx, 1, QTableWidgetItem(filename))

            cached_file = find_cached_spec_file(filename, s_num)
            if cached_file:
                cache_text = f"🟢 Cached ({cached_file.suffix[1:].upper()})"
                cache_item = QTableWidgetItem(cache_text)
                cache_item.setForeground(Qt.darkGreen)
            else:
                cache_item = QTableWidgetItem("🌐 Remote (FTP)")
                cache_item.setForeground(Qt.darkGray)
            self.table.setItem(row_idx, 2, cache_item)

            in_db = (s_num, version) in imported_entries
            db_item = QTableWidgetItem("✅ Ingested" if in_db else "⚪ Ready")
            if in_db:
                db_item.setForeground(Qt.blue)
            self.table.setItem(row_idx, 3, db_item)

    def _on_item_double_clicked(self, item: QTableWidgetItem):
        row = item.row()
        self.selected_files_info = [self.table.item(row, 0).data(Qt.UserRole)]
        self.accept()

    def _on_accept(self):
        selected_rows = sorted(list(set(item.row() for item in self.table.selectedItems())))
        if not selected_rows:
            QMessageBox.warning(
                self, "Selection Required", "Please select at least one specification version to import."
            )
            return

        self.selected_files_info = [
            self.table.item(row, 0).data(Qt.UserRole)
            for row in selected_rows
        ]
        self.accept()


class NASTab(QWidget):
    log_msg = pyqtSignal(str, int)

    def __init__(self, nas_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.nas_db_path = Path(nas_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = self.nas_db_path.parent / "nas_config.json"

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
        self._current_clause_defs: Dict[str, Dict[str, Any]] = {}
        self._current_ie_clause: Optional[str] = None
        self._current_ie_name: Optional[str] = None
        self._current_ie_spec: Optional[str] = None
        self._current_message_spec: Optional[str] = None

        self._updating_checks: bool = False
        self._updating_combo: bool = False
        self._loading_config: bool = False
        self._initialized_filters: bool = False

        self._msg_search_timer = QTimer(self)
        self._msg_search_timer.setSingleShot(True)
        self._msg_search_timer.setInterval(250)
        self._msg_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._ie_search_timer = QTimer(self)
        self._ie_search_timer.setSingleShot(True)
        self._ie_search_timer.setInterval(250)
        self._ie_search_timer.timeout.connect(self._on_search_timer_timeout)

        self._setup_ui()
        self.refresh_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

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

        self.wipe_db_btn = QPushButton("⚠️ Wipe DB")
        self.wipe_db_btn.setStyleSheet("color: #D32F2F; font-weight: bold;")
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        toolbar.addWidget(self.wipe_db_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        main_splitter = QSplitter(Qt.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        ver_group = QGroupBox("Specification Versions && Releases")
        ver_layout = QVBoxLayout(ver_group)

        self.version_tree = QTreeWidget()
        self.version_tree.setHeaderHidden(True)
        self.version_tree.setAlternatingRowColors(True)
        self.version_tree.setIndentation(14)
        self.version_tree.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                padding: 1px;
            }
            QTreeWidget::item {
                padding: 1px 0px;
                min-height: 18px;
            }
            QTreeWidget::item:hover {
                background-color: #F1F5F9;
            }
            QTreeWidget::item:selected {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        self.version_tree.itemChanged.connect(self._on_version_tree_item_changed)
        self.version_tree.itemExpanded.connect(lambda _: self._save_config())
        self.version_tree.itemCollapsed.connect(lambda _: self._save_config())
        self.version_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.version_tree.customContextMenuRequested.connect(self._on_version_tree_context_menu)
        ver_layout.addWidget(self.version_tree)
        left_layout.addWidget(ver_group)

        msg_group = QGroupBox("Protocol Messages && SIBs")
        msg_layout = QVBoxLayout(msg_group)

        self.msg_search = QLineEdit()
        self.msg_search.setPlaceholderText("Filter message/SIB name (e.g. SIB1, Reconfig)...")
        self.msg_search.setClearButtonEnabled(True)
        self.msg_search.textChanged.connect(lambda: self._msg_search_timer.start())
        msg_layout.addWidget(self.msg_search)

        # IE Search Layout with Extended Description Search Toggle Button
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
                color: #0369A1;
                background-color: #E0F2FE;
                border-color: #0284C7;
            }
            QPushButton:checked {
                color: #0369A1;
                background-color: #E0F2FE;
                border: 1.5px solid #0284C7;
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

        right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)
        self.matrix_title = QLabel("Select a Message to View Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        matrix_layout.addWidget(self.matrix_title)

        self.matrix_table = QTableView()
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QTableView.SelectRows)
        self.matrix_table.setSelectionMode(QTableView.SingleSelection)
        self.matrix_table.verticalHeader().setDefaultSectionSize(26)
        self.matrix_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.matrix_table.clicked.connect(self._on_table_cell_clicked)

        self.matrix_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.matrix_table.customContextMenuRequested.connect(self._on_matrix_context_menu)

        matrix_layout.addWidget(self.matrix_table)
        right_splitter.addWidget(matrix_widget)

        inspector_group = QGroupBox("Structure && Field Descriptions Inspector")
        inspector_layout = QVBoxLayout(inspector_group)
        inspector_layout.setContentsMargins(8, 8, 8, 8)

        insp_header = QHBoxLayout()
        insp_header.setContentsMargins(0, 0, 0, 4)

        self.inspector_title_lbl = QLabel("No Information Element selected")
        self.inspector_title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        insp_header.addWidget(self.inspector_title_lbl)

        insp_header.addStretch()

        self.ie_usage_btn = QPushButton("Used in: 0 messages ▾")
        self.ie_usage_btn.setVisible(False)
        self.ie_usage_btn.setCursor(Qt.PointingHandCursor)
        self.ie_usage_btn.setToolTip("View other messages that contain this Information Element")
        self.ie_usage_btn.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: bold;
                color: #0369A1;
                background-color: #E0F2FE;
                border: 1px solid #BAE6FD;
                border-radius: 4px;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #BAE6FD;
                border-color: #0284C7;
            }
        """)
        self.ie_usage_btn.clicked.connect(self._show_usage_menu)
        insp_header.addWidget(self.ie_usage_btn)

        self.inspector_version_combo = QComboBox()
        self.inspector_version_combo.setToolTip("Switch specification release for this definition")
        self.inspector_version_combo.setStyleSheet("""
            QComboBox {
                font-weight: bold;
                padding: 2px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #F8FAFC;
                min-width: 110px;
            }
            QComboBox:hover {
                border-color: #0284C7;
                background-color: #FFFFFF;
            }
        """)
        self.inspector_version_combo.currentIndexChanged.connect(self._on_inspector_version_changed)
        insp_header.addWidget(self.inspector_version_combo)
        inspector_layout.addLayout(insp_header)

        self.inspector_text = QTextEdit()
        self.inspector_text.setReadOnly(True)
        self.inspector_text.setPlaceholderText(
            "Click on an Information Element above to inspect its details and field descriptions..."
        )
        inspector_layout.addWidget(self.inspector_text)
        right_splitter.addWidget(inspector_group)

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
        if self._loading_config or self._updating_checks:
            return

        try:
            checked_versions = []
            collapsed_specs = []

            for s_idx in range(1, self.version_tree.topLevelItemCount()):
                spec_item = self.version_tree.topLevelItem(s_idx)
                s_data = spec_item.data(0, Qt.UserRole) or {}
                spec_num = s_data.get("spec_number", "")
                if not spec_item.isExpanded():
                    collapsed_specs.append(spec_num)

                for c_idx in range(spec_item.childCount()):
                    child = spec_item.child(c_idx)
                    if child.checkState(0) == Qt.Checked:
                        c_data = child.data(0, Qt.UserRole) or {}
                        checked_versions.append({
                            "spec_number": c_data.get("spec_number"),
                            "version": c_data.get("version"),
                        })

            config_data = {
                "checked_versions": checked_versions,
                "msg_filter": self.msg_search.text(),
                "ie_filter": self.ie_search.text(),
                "search_descriptions": self.deep_search_btn.isChecked(),
                "selected_message": self.current_selected_message_name or "",
                "collapsed_specs": collapsed_specs,
            }

            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.warning(f"Could not save Protocol config to {self.config_path}: {e}")

    # -------------------------------------------------------------------------
    # Search Modes & Context Menus
    # -------------------------------------------------------------------------

    def _on_deep_search_toggled(self, checked: bool):
        if checked:
            self.ie_search.setPlaceholderText("Filter by Field or Description (e.g. emergency)...")
            self.deep_search_btn.setText("📖 Desc: ON")
        else:
            self.ie_search.setPlaceholderText("Filter by IE / Field (e.g. RadioBearer, CellGroup)...")
            self.deep_search_btn.setText("📖 Desc")

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
        self._updating_checks = True
        self.version_tree.clear()
        versions = self.db.get_imported_versions()

        if not versions:
            self.selected_version_ids = []
            self._updating_checks = False
            self._populate_messages()
            return

        saved_config = self._load_config()
        saved_checked = saved_config.get("checked_versions")
        saved_collapsed = set(saved_config.get("collapsed_specs", []))

        saved_tuples: Optional[Set[Tuple[str, str]]] = None
        if saved_checked is not None:
            saved_tuples = {(item.get("spec_number"), item.get("version")) for item in saved_checked}

        base_font = self.version_tree.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        text_color_header = QBrush(QColor("#0F172A"))
        text_color_item = QBrush(QColor("#334155"))

        # 1. Master toggle item
        all_item = QTreeWidgetItem(self.version_tree)
        all_item.setText(0, "All Specifications")
        all_item.setData(0, Qt.UserRole, {"type": "all"})
        all_item.setFlags(all_item.flags() | Qt.ItemIsUserCheckable)
        all_item.setFont(0, bold_font)
        all_item.setForeground(0, text_color_header)
        all_item.setCheckState(0, Qt.Checked)

        # 2. Group versions by specification number
        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        sorted_specs = sorted(specs_map.keys())

        for spec_num in sorted_specs:
            spec_versions = specs_map[spec_num]
            spec_item = QTreeWidgetItem(self.version_tree)
            spec_title = get_spec_title(spec_num)
            spec_item.setText(0, spec_title)
            spec_item.setData(0, Qt.UserRole, {"type": "spec", "spec_number": spec_num})
            spec_item.setFlags(spec_item.flags() | Qt.ItemIsUserCheckable)
            spec_item.setFont(0, bold_font)
            spec_item.setForeground(0, text_color_header)
            spec_item.setCheckState(0, Qt.Checked)

            sorted_v_list = sorted(spec_versions, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

            for v in sorted_v_list:
                child = QTreeWidgetItem(spec_item)
                child.setText(0, f"v{v['version']}")
                child.setToolTip(0, f"TS {v['spec_number']} v{v['version']}\n(Right-click to delete)")
                child.setData(
                    0,
                    Qt.UserRole,
                    {
                        "type": "version",
                        "id": v["id"],
                        "spec_number": v["spec_number"],
                        "version": v["version"],
                    },
                )
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                child.setFont(0, base_font)
                child.setForeground(0, text_color_item)

                if saved_tuples is not None:
                    is_checked = (v["spec_number"], v["version"]) in saved_tuples
                    child.setCheckState(0, Qt.Checked if is_checked else Qt.Unchecked)
                else:
                    child.setCheckState(0, Qt.Checked)

            if spec_num in saved_collapsed:
                spec_item.setExpanded(False)
            else:
                spec_item.setExpanded(True)

        self._update_parent_states()
        self._recalculate_selected_versions()
        self._updating_checks = False

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

    def _on_version_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._updating_checks:
            return

        self._updating_checks = True
        state = item.checkState(0)
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        if item_type == "all":
            target_state = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for s_idx in range(1, self.version_tree.topLevelItemCount()):
                spec_item = self.version_tree.topLevelItem(s_idx)
                spec_item.setCheckState(0, target_state)
                for c_idx in range(spec_item.childCount()):
                    child = spec_item.child(c_idx)
                    child.setCheckState(0, target_state)
        elif item_type == "spec":
            target_state = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for c_idx in range(item.childCount()):
                child = item.child(c_idx)
                child.setCheckState(0, target_state)
            self._update_parent_states()
        elif item_type == "version":
            self._update_parent_states()

        self._recalculate_selected_versions()
        self._updating_checks = False

        self._save_config()
        self._populate_messages()

        current_msg_item = self.msg_list.currentItem()
        if current_msg_item:
            self._on_message_clicked(current_msg_item)
        else:
            self.matrix_table.setModel(None)

    def _on_version_tree_context_menu(self, pos):
        item = self.version_tree.itemAt(pos)
        if not item:
            return

        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

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
                background-color: #FEE2E2;
                color: #B91C1C;
            }
        """)

        if item_type == "version":
            spec_num = data.get("spec_number", "")
            version = data.get("version", "")
            act_delete = QAction(f"🗑️ Delete TS {spec_num} v{version}", self)
            act_delete.triggered.connect(lambda: self._delete_single_version(spec_num, version))
            menu.addAction(act_delete)
            menu.exec_(self.version_tree.viewport().mapToGlobal(pos))

        elif item_type == "spec":
            spec_num = data.get("spec_number", "")
            child_count = item.childCount()
            act_delete_spec = QAction(f"🗑️ Delete all versions of TS {spec_num} ({child_count} versions)", self)
            act_delete_spec.triggered.connect(lambda: self._delete_spec_group(spec_num, item))
            menu.addAction(act_delete_spec)
            menu.exec_(self.version_tree.viewport().mapToGlobal(pos))

    def _delete_single_version(self, spec_number: str, version: str):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete all stored data for TS {spec_number} v{version}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            success = self.db.clear_version(spec_number, version)
            if success:
                self.current_selected_message_name = None
                self.refresh_versions()
                self._save_config()
                self.log_msg.emit(f"🗑️ Deleted TS {spec_number} v{version} from database.", logging.INFO)
            else:
                QMessageBox.warning(self, "Error", f"Failed to delete TS {spec_number} v{version}.")

    def _delete_spec_group(self, spec_number: str, spec_item: QTreeWidgetItem):
        child_count = spec_item.childCount()
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete all {child_count} version(s) of TS {spec_number} from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for c_idx in range(child_count):
                child = spec_item.child(c_idx)
                c_data = child.data(0, Qt.UserRole) or {}
                ver = c_data.get("version")
                if ver:
                    self.db.clear_version(spec_number, ver)
            self.current_selected_message_name = None
            self.refresh_versions()
            self._save_config()
            self.log_msg.emit(f"🗑️ Deleted all versions of TS {spec_number} from database.", logging.INFO)

    def _update_parent_states(self):
        total_version_count = 0
        total_checked_count = 0

        for s_idx in range(1, self.version_tree.topLevelItemCount()):
            spec_item = self.version_tree.topLevelItem(s_idx)
            child_count = spec_item.childCount()
            if child_count == 0:
                continue

            checked_children = sum(
                1 for c_idx in range(child_count)
                if spec_item.child(c_idx).checkState(0) == Qt.Checked
            )

            total_version_count += child_count
            total_checked_count += checked_children

            if checked_children == child_count:
                spec_item.setCheckState(0, Qt.Checked)
            elif checked_children == 0:
                spec_item.setCheckState(0, Qt.Unchecked)
            else:
                spec_item.setCheckState(0, Qt.PartiallyChecked)

        all_item = self.version_tree.topLevelItem(0)
        if all_item:
            if total_version_count > 0 and total_checked_count == total_version_count:
                all_item.setCheckState(0, Qt.Checked)
            elif total_checked_count == 0:
                all_item.setCheckState(0, Qt.Unchecked)
            else:
                all_item.setCheckState(0, Qt.PartiallyChecked)

    def _recalculate_selected_versions(self):
        selected_ids = []
        for s_idx in range(1, self.version_tree.topLevelItemCount()):
            spec_item = self.version_tree.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    data = child.data(0, Qt.UserRole)
                    if data and "id" in data:
                        selected_ids.append(data["id"])
        self.selected_version_ids = selected_ids

    # -------------------------------------------------------------------------
    # Message & Matrix Handlers
    # -------------------------------------------------------------------------

    def _on_search_timer_timeout(self):
        self._save_config()
        self._populate_messages()

    def _populate_messages(self):
        target_msg_name = self.current_selected_message_name
        self.msg_list.clear()

        if not self.selected_version_ids:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
            self.inspector_title_lbl.setText("No Information Element selected")
            self.ie_usage_btn.setVisible(False)
            self.inspector_version_combo.clear()
            self.inspector_text.clear()
            self.current_selected_message_name = None
            return

        ie_query = self.ie_search.text().strip()
        msg_query = self.msg_search.text().strip().lower()
        search_desc = self.deep_search_btn.isChecked()

        if ie_query:
            messages = self.db.get_messages_by_ie_search(
                ie_query,
                self.selected_version_ids,
                search_descriptions=search_desc
            )
        else:
            messages = self.db.get_messages_list(self.selected_version_ids)

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

        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        elif self.msg_list.count() > 0 and not target_msg_name:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
        else:
            self.matrix_table.setModel(None)
            self.matrix_title.setText("Select a Message to View Evolution Matrix")
            self.inspector_title_lbl.setText("No Information Element selected")
            self.ie_usage_btn.setVisible(False)
            self.inspector_version_combo.clear()
            self.inspector_text.clear()

    def _on_message_clicked(self, item: QListWidgetItem):
        msg_name = item.data(Qt.UserRole)
        self.current_selected_message_name = msg_name
        self._save_config()

        ie_query = self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()
        desc_label = " (incl. Desc)" if search_desc and ie_query else ""
        title_suffix = f" (Filtered by IE{desc_label}: '{ie_query}')" if ie_query else ""

        df = self.db.get_message_evolution_df(msg_name, self.selected_version_ids)

        specs = []
        if not df.empty and "spec_number" in df.columns:
            unique_specs = sorted(df["spec_number"].dropna().unique())
            specs = [f"TS {s}" if not str(s).startswith("TS") else str(s) for s in unique_specs]
            self._current_message_spec = str(unique_specs[0]) if len(unique_specs) == 1 else None
        else:
            self._current_message_spec = None

        spec_prefix = f" ({', '.join(specs)})" if specs else ""
        self.matrix_title.setText(f"Message{spec_prefix}: {msg_name}{title_suffix}")

        model = NASEvolutionMatrixModel(df, ie_filter=ie_query, search_descriptions=search_desc)
        self.matrix_table.setModel(model)
        self.matrix_table.resizeColumnsToContents()

    def _on_table_cell_clicked(self, index):
        model = self.matrix_table.model()
        if not model:
            return

        row = index.row()
        col = index.column()

        ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
        type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

        spec_num = getattr(self, "_current_message_spec", None)
        if col >= 3:
            col_name = str(model._get_visible_column_name(col))
            match_spec = re.search(r"(?:24|36|38)\.[0-9]{3}", col_name)
            if match_spec:
                spec_num = match_spec.group(0)

        match = re.search(r"((?:9|6|D\.6)(?:\.[0-9A-Za-z]+)+)", type_ref)
        clause = match.group(1).strip() if match else ie_name

        self._current_ie_clause = clause
        self._current_ie_name = ie_name
        self._current_ie_spec = spec_num

        defs = self.db.get_ie_definitions_by_clause(
            clause, spec_number=spec_num, version_ids=self.selected_version_ids
        )
        if not defs:
            defs = self.db.get_ie_definitions_by_clause(clause, version_ids=self.selected_version_ids)

        if not defs:
            self.inspector_title_lbl.setText(f"{ie_name} ({type_ref})")
            self.ie_usage_btn.setVisible(False)
            self.inspector_version_combo.clear()
            self._current_clause_defs.clear()
            self.inspector_text.setPlainText(f"Type / Reference: {type_ref}\n(No structure definition found)")
            return

        resolved_name = defs[0]["ie_name"]
        spec_badge = f" (TS {spec_num})" if spec_num else ""
        self.inspector_title_lbl.setText(f"{resolved_name}{spec_badge}")
        self._current_clause_defs = {f"{d['spec_number']} v{d['version']}": d for d in defs}

        containing_msgs = self.db.get_messages_using_ie(
            clause=clause,
            ie_name=resolved_name,
            spec_number=spec_num,
            version_ids=self.selected_version_ids,
        )
        num_msgs = len(containing_msgs)
        self.ie_usage_btn.setText(f"Used in: {num_msgs} message{'s' if num_msgs != 1 else ''} ▾")
        self.ie_usage_btn.setVisible(True)

        self._updating_combo = True
        self.inspector_version_combo.clear()
        for d in defs:
            key_name = f"TS {d['spec_number']} v{d['version']}"
            self.inspector_version_combo.addItem(key_name, f"{d['spec_number']} v{d['version']}")

        target_version_key: Optional[str] = None
        if col >= 3:
            col_name = str(model._get_visible_column_name(col))
            clean_col = col_name.replace("TS ", "").strip()
            if clean_col in self._current_clause_defs:
                target_version_key = clean_col
            else:
                for k in self._current_clause_defs:
                    if col_name in k or k.endswith(f"v{col_name}"):
                        target_version_key = k
                        break

        if not target_version_key and defs:
            target_version_key = f"{defs[0]['spec_number']} v{defs[0]['version']}"

        for idx in range(self.inspector_version_combo.count()):
            if self.inspector_version_combo.itemData(idx) == target_version_key:
                self.inspector_version_combo.setCurrentIndex(idx)
                break

        self._updating_combo = False

        selected_def = self._current_clause_defs.get(target_version_key) or defs[0]
        self.inspector_text.setHtml(selected_def["raw_description"])

    def _show_usage_menu(self):
        if not self._current_ie_clause:
            return

        clause = self._current_ie_clause
        name = self._current_ie_name or ""
        spec_num = getattr(self, "_current_ie_spec", None)

        containing_msgs = self.db.get_messages_using_ie(
            clause=clause,
            ie_name=name,
            spec_number=spec_num,
            version_ids=self.selected_version_ids,
        )

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                font-size: 11px;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                padding: 4px;
            }
            QMenu::item {
                padding: 4px 20px 4px 10px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #E0F2FE;
                color: #0369A1;
            }
            QMenu::separator {
                height: 1px;
                background-color: #E2E8F0;
                margin: 4px 0;
            }
        """)

        header_title = f"Messages referencing {name or clause}"
        if spec_num:
            header_title += f" [TS {spec_num}]"
        header_action = QAction(f"{header_title}:", self)
        header_action.setEnabled(False)
        menu.addAction(header_action)
        menu.addSeparator()

        if containing_msgs:
            for m in containing_msgs:
                msg_name = m["message_name"]
                clause_ref = m["clause"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [TS {m_spec}]" if m_spec and not spec_num else ""
                is_current = msg_name == self.current_selected_message_name
                label = f"{'👉 ' if is_current else ''}{msg_name} ({clause_ref}){spec_tag}"
                action = QAction(label, self)
                action.triggered.connect(lambda checked, mn=msg_name: self._jump_to_message(mn))
                menu.addAction(action)
        else:
            none_act = QAction("No messages found in active releases", self)
            none_act.setEnabled(False)
            menu.addAction(none_act)

        menu.addSeparator()
        filter_action = QAction(f"🔍 Filter message list by '{name or clause}'", self)
        filter_action.triggered.connect(lambda: self.ie_search.setText(name or clause))
        menu.addAction(filter_action)

        menu.exec_(self.ie_usage_btn.mapToGlobal(self.ie_usage_btn.rect().bottomLeft()))

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

        match = re.search(r"((?:9|6|D\.6)(?:\.[0-9A-Za-z]+)+)", type_ref)
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
                background-color: #E0F2FE;
                color: #0369A1;
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

    def _on_inspector_version_changed(self, index: int):
        if self._updating_combo or index < 0:
            return

        version_key = self.inspector_version_combo.itemData(index)
        if version_key in self._current_clause_defs:
            self.inspector_text.setHtml(self._current_clause_defs[version_key]["raw_description"])

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
            parser_temp = NASDocxParser(paths)
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

    def _on_import_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_file_btn.setEnabled(True)
        QMessageBox.critical(self, "Ingestion Error", err)
        self.log_msg.emit(f"❌ Ingestion failed: {err}", logging.ERROR)

    def _on_clear_version_clicked(self):
        checked_versions_to_delete = []
        for s_idx in range(1, self.version_tree.topLevelItemCount()):
            spec_item = self.version_tree.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    c_data = child.data(0, Qt.UserRole)
                    if c_data and c_data.get("type") == "version":
                        checked_versions_to_delete.append((c_data["spec_number"], c_data["version"]))

        if not checked_versions_to_delete:
            QMessageBox.warning(self, "Select Version", "Please check at least one specific version to clear.")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete {len(checked_versions_to_delete)} checked specification version(s)?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for s_num, ver in checked_versions_to_delete:
                self.db.clear_version(s_num, ver)
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
            self.inspector_text.clear()
            self.log_msg.emit("🧹 Protocol Database wiped.", logging.INFO)