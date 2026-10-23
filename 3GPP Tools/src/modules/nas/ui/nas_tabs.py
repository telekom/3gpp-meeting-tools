import json
import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
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
    QStackedWidget,
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
from modules.nas.ui.nas_components import (
    NASInspectorWidget,
    NASMessageItemDelegate,
    NASVersionTreeWidget,
)
from modules.nas.ui.nas_dialogs import (
    NASVersionSelectDialog,
    get_spec_metadata,
    get_spec_title,
    get_spec_tooltip,
)
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.nas.ui.nas_ingest_controller import (
    clear_version_clicked,
    delete_single_version,
    delete_spec_group,
    fetch_from_specs_db_clicked,
    import_cross_ref_spec_requested,
    import_local_file_clicked,
    on_import_error,
    on_import_progress,
    on_import_success,
    start_batch_ingestion,
    wipe_db_clicked,
)
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

TINY_HEADER_BTN_STYLE = """
    QPushButton {
        font-size: 10px;
        font-weight: 500;
        color: #475569;
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 3px;
        padding: 1px 7px;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        color: #0F172A;
        border-color: #94A3B8;
    }
    QPushButton:pressed {
        background-color: #E2E8F0;
    }
"""

RIBBON_BADGE_STYLE = """
    background-color: #F1F5F9;
    color: #334155;
    font-size: 11px;
    font-weight: 600;
    border: 1px solid #E2E8F0;
    border-radius: 11px;
    padding: 3px 10px;
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
        layout.setSpacing(8)

        # ---------------------------------------------------------------------
        # 1. Top Action Bar: Primary Ingestion, Stats Ribbon, and Manage Menu
        # ---------------------------------------------------------------------
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

        # Subtle separator between actions and metrics
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        sep.setStyleSheet("color: #CBD5E1;")
        toolbar.addWidget(sep)

        # Live Database KPI Stats Ribbon
        self.stats_specs_lbl = QLabel("📦 0 Specs")
        self.stats_specs_lbl.setStyleSheet(RIBBON_BADGE_STYLE)
        toolbar.addWidget(self.stats_specs_lbl)

        self.stats_vers_lbl = QLabel("🏷️ 0 Versions")
        self.stats_vers_lbl.setStyleSheet(RIBBON_BADGE_STYLE)
        toolbar.addWidget(self.stats_vers_lbl)

        self.stats_msgs_lbl = QLabel("📨 0 Items")
        self.stats_msgs_lbl.setStyleSheet(RIBBON_BADGE_STYLE)
        toolbar.addWidget(self.stats_msgs_lbl)

        toolbar.addStretch()

        # Manage / Maintenance Dropdown Menu
        self.manage_btn = QPushButton("⚙️ Manage ▾")
        self.manage_btn.setCursor(Qt.PointingHandCursor)
        self.manage_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)

        manage_menu = QMenu(self)
        manage_menu.setStyleSheet("""
            QMenu {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }
            QMenu::item {
                padding: 6px 18px 6px 10px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #EBF3FC;
                color: #1E5C99;
            }
        """)

        act_clear = QAction("🗑️ Clear Selected Version / Spec...", self)
        act_clear.triggered.connect(self._on_clear_version_clicked)
        manage_menu.addAction(act_clear)

        manage_menu.addSeparator()

        act_wipe = QAction("⚠️ Wipe Protocol Database...", self)
        act_wipe.triggered.connect(self._on_wipe_db_clicked)
        manage_menu.addAction(act_wipe)

        self.manage_btn.setMenu(manage_menu)
        toolbar.addWidget(self.manage_btn)

        layout.addLayout(toolbar)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # ---------------------------------------------------------------------
        # 2. Main Horizontal Splitter
        # ---------------------------------------------------------------------
        main_splitter = QSplitter(Qt.Horizontal)

        # --- Left Panel (Spec Tree Card + Protocol Library Card) ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        # Specification Versions Card Frame
        ver_card = QFrame()
        ver_card.setObjectName("cardFrame")
        ver_card.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        ver_card_layout = QVBoxLayout(ver_card)
        ver_card_layout.setContentsMargins(0, 0, 0, 0)
        ver_card_layout.setSpacing(0)

        ver_header = QFrame()
        ver_header.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border-bottom: 1px solid #E2E8F0;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }
        """)
        ver_header_layout = QHBoxLayout(ver_header)
        ver_header_layout.setContentsMargins(6, 4, 6, 4)
        ver_header_layout.setSpacing(6)

        ver_title = QLabel("🏷️ Specification Releases")
        ver_title.setStyleSheet("font-weight: bold; font-size: 11px; color: #1E293B;")
        ver_header_layout.addWidget(ver_title)
        ver_header_layout.addStretch()

        btn_all = QPushButton("All")
        btn_all.setToolTip("Select all specifications and releases")
        btn_all.setStyleSheet(TINY_HEADER_BTN_STYLE)
        btn_all.clicked.connect(self._on_select_all_specs)
        ver_header_layout.addWidget(btn_all)

        btn_latest = QPushButton("Latest")
        btn_latest.setToolTip("Select only latest version for each specification")
        btn_latest.setStyleSheet(TINY_HEADER_BTN_STYLE)
        btn_latest.clicked.connect(self._on_select_latest_specs)
        ver_header_layout.addWidget(btn_latest)

        btn_none = QPushButton("Clear")
        btn_none.setToolTip("Deselect all specifications")
        btn_none.setStyleSheet(TINY_HEADER_BTN_STYLE)
        btn_none.clicked.connect(self._on_select_none_specs)
        ver_header_layout.addWidget(btn_none)

        ver_card_layout.addWidget(ver_header)

        self.version_tree = NASVersionTreeWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.structure_changed.connect(self._save_config)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        ver_card_layout.addWidget(self.version_tree)
        left_layout.addWidget(ver_card, 1)

        # Protocol Library Card Frame
        msg_card = QFrame()
        msg_card.setObjectName("cardFrame")
        msg_card.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        msg_card_layout = QVBoxLayout(msg_card)
        msg_card_layout.setContentsMargins(0, 0, 0, 0)
        msg_card_layout.setSpacing(0)

        msg_header = QFrame()
        msg_header.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border-bottom: 1px solid #E2E8F0;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }
        """)
        msg_header_layout = QHBoxLayout(msg_header)
        msg_header_layout.setContentsMargins(6, 4, 6, 4)
        msg_header_layout.setSpacing(6)

        msg_card_title = QLabel("📨 Protocol Library")
        msg_card_title.setStyleSheet("font-weight: bold; font-size: 11px; color: #1E293B;")
        msg_header_layout.addWidget(msg_card_title)
        msg_header_layout.addStretch()

        self.msg_count_lbl = QLabel("0 items")
        self.msg_count_lbl.setStyleSheet("""
            color: #64748B;
            font-size: 10px;
            font-weight: bold;
            background-color: #F1F5F9;
            border: 1px solid #E2E8F0;
            border-radius: 9px;
            padding: 1px 7px;
        """)
        msg_header_layout.addWidget(self.msg_count_lbl)
        msg_card_layout.addWidget(msg_header)

        msg_body = QWidget()
        msg_body_layout = QVBoxLayout(msg_body)
        msg_body_layout.setContentsMargins(6, 6, 6, 6)
        msg_body_layout.setSpacing(5)

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

        msg_body_layout.addLayout(mode_bar_layout)

        self.msg_search = QLineEdit()
        self.msg_search.setPlaceholderText("Filter by name, acronym, or spec (e.g. SIB1, PFCP, 24.501)...")
        self.msg_search.setClearButtonEnabled(True)
        self.msg_search.textChanged.connect(lambda: self._msg_search_timer.start())
        msg_body_layout.addWidget(self.msg_search)

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
        msg_body_layout.addLayout(ie_search_bar_layout)

        self.msg_list = QListWidget()
        self.msg_list.setItemDelegate(NASMessageItemDelegate(self.msg_list))
        self.msg_list.itemClicked.connect(self._on_message_clicked)
        self.msg_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                padding: 2px;
                outline: none;
            }
        """)
        msg_body_layout.addWidget(self.msg_list)
        msg_card_layout.addWidget(msg_body)
        left_layout.addWidget(msg_card, 2)

        main_splitter.addWidget(left_widget)

        # --- Right Panel (Matrix Card + Inspector Card) ---
        right_splitter = QSplitter(Qt.Vertical)

        matrix_card = QFrame()
        matrix_card.setObjectName("cardFrame")
        matrix_card.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        matrix_card_layout = QVBoxLayout(matrix_card)
        matrix_card_layout.setContentsMargins(0, 0, 0, 0)
        matrix_card_layout.setSpacing(0)

        matrix_header_frame = QFrame()
        matrix_header_frame.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border-bottom: 1px solid #E2E8F0;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }
        """)
        matrix_header_layout = QHBoxLayout(matrix_header_frame)
        matrix_header_layout.setContentsMargins(6, 4, 6, 4)
        matrix_header_layout.setSpacing(8)

        self.matrix_title = QLabel("Select a Message to View Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B;")
        matrix_header_layout.addWidget(self.matrix_title)

        # Diff Legend Pill Bar
        diff_legend = QHBoxLayout()
        diff_legend.setSpacing(4)

        lbl_added = QLabel("🟢 Added")
        lbl_added.setStyleSheet("background-color: #E8F5E9; color: #166534; font-size: 10px; font-weight: bold; border: 1px solid #BBF7D0; border-radius: 3px; padding: 1px 5px;")
        diff_legend.addWidget(lbl_added)

        lbl_mod = QLabel("🟡 Modified")
        lbl_mod.setStyleSheet("background-color: #FEF9C3; color: #854D0E; font-size: 10px; font-weight: bold; border: 1px solid #FEF08A; border-radius: 3px; padding: 1px 5px;")
        diff_legend.addWidget(lbl_mod)

        lbl_rem = QLabel("🔴 Removed")
        lbl_rem.setStyleSheet("background-color: #FEE2E2; color: #991B1B; font-size: 10px; font-weight: bold; border: 1px solid #FECACA; border-radius: 3px; padding: 1px 5px;")
        diff_legend.addWidget(lbl_rem)

        matrix_header_layout.addLayout(diff_legend)
        matrix_header_layout.addStretch()

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

        self.copy_matrix_btn = QPushButton("📋 Copy TSV")
        self.copy_matrix_btn.setCursor(Qt.PointingHandCursor)
        self.copy_matrix_btn.setToolTip("Copy entire evolution matrix to clipboard as tab-separated values (for Excel / Word)")
        self.copy_matrix_btn.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 500;
                color: #334155;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                color: #0F172A;
                border-color: #94A3B8;
            }
            QPushButton:disabled {
                color: #94A3B8;
                background-color: #F8FAFC;
                border-color: #E2E8F0;
            }
        """)
        self.copy_matrix_btn.setEnabled(False)
        self.copy_matrix_btn.clicked.connect(self._copy_matrix_tsv)
        matrix_header_layout.addWidget(self.copy_matrix_btn)

        matrix_card_layout.addWidget(matrix_header_frame)

        # Matrix Stack: Page 0 = Empty State View, Page 1 = Table View
        self.matrix_stack = QStackedWidget()

        self.matrix_empty_widget = QWidget()
        matrix_empty_layout = QVBoxLayout(self.matrix_empty_widget)
        matrix_empty_layout.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.setSpacing(8)

        m_empty_icon = QLabel("🔀")
        m_empty_icon.setStyleSheet("font-size: 32px;")
        m_empty_icon.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.addWidget(m_empty_icon)

        m_empty_title = QLabel("Select a Protocol Message or Information Element")
        m_empty_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        m_empty_title.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.addWidget(m_empty_title)

        m_empty_sub = QLabel(
            "Choose a message or IE from the library on the left to inspect its evolution across 3GPP releases.\n"
            "Field presence, ASN.1 syntax, and format changes will be diff-highlighted here."
        )
        m_empty_sub.setStyleSheet("font-size: 11px; color: #64748B; line-height: 1.4;")
        m_empty_sub.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.addWidget(m_empty_sub)

        self.matrix_stack.addWidget(self.matrix_empty_widget)

        self.matrix_table = QTableView()
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QTableView.SelectRows)
        self.matrix_table.setSelectionMode(QTableView.SingleSelection)
        self.matrix_table.verticalHeader().setDefaultSectionSize(24)
        self.matrix_table.verticalHeader().setVisible(False)
        self.matrix_table.setStyleSheet("""
            QTableView {
                border: none;
                gridline-color: #F1F5F9;
                background-color: #FFFFFF;
                alternate-background-color: #F8FAFC;
            }
            QTableView::item:selected {
                background-color: #EBF3FC;
                color: #0F172A;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #334155;
                font-weight: bold;
                font-size: 11px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
                border-right: 1px solid #E2E8F0;
                padding: 4px 6px;
            }
        """)

        h_header = self.matrix_table.horizontalHeader()
        h_header.setHighlightSections(False)
        h_header.setStretchLastSection(False)

        self.matrix_table.clicked.connect(self._on_table_cell_clicked)
        self.matrix_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.matrix_table.customContextMenuRequested.connect(self._on_matrix_context_menu)
        self.matrix_stack.addWidget(self.matrix_table)

        matrix_card_layout.addWidget(self.matrix_stack)
        right_splitter.addWidget(matrix_card)

        # Inspector Card (Bottom Right)
        self.inspector = NASInspectorWidget(db=self.db)
        self.inspector.jump_to_message_requested.connect(self._jump_to_message)
        self.inspector.filter_by_ie_requested.connect(self._on_filter_by_ie_requested)
        self.inspector.import_spec_requested.connect(self._on_import_cross_ref_spec_requested)
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([380, 320])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([320, 680])
        layout.addWidget(main_splitter)

    # -------------------------------------------------------------------------
    # Quick Spec Tree Helpers
    # -------------------------------------------------------------------------

    def _on_select_all_specs(self):
        self.version_tree.select_all()

    def _on_select_latest_specs(self):
        self.version_tree.select_latest_only()

    def _on_select_none_specs(self):
        self.version_tree.select_none()

    def _copy_matrix_tsv(self):
        """Exports the active evolution matrix to the system clipboard as tab-separated values."""
        model = self.matrix_table.model()
        if not model or model.rowCount() == 0:
            return

        headers = [
            str(model.headerData(c, Qt.Horizontal, Qt.DisplayRole) or "").replace("\n", " ")
            for c in range(model.columnCount())
        ]
        lines = ["\t".join(headers)]

        for r in range(model.rowCount()):
            row_vals = [
                str(model.data(model.index(r, c), Qt.DisplayRole) or "")
                for c in range(model.columnCount())
            ]
            lines.append("\t".join(row_vals))

        tsv_content = "\n".join(lines)
        QApplication.clipboard().setText(tsv_content)
        prev_text = self.copy_matrix_btn.text()
        self.copy_matrix_btn.setText("✓ Copied!")
        QTimer.singleShot(1500, lambda: self.copy_matrix_btn.setText(prev_text))

    def _update_stats_ribbon(self):
        """Updates live KPI stats badges in the toolbar ribbon."""
        try:
            versions = self.db.get_imported_versions()
            unique_specs = len(set(v.get("spec_number") for v in versions if v.get("spec_number")))
            total_versions = len(versions)

            with self.db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(DISTINCT message_name) FROM nas_messages")
                total_msgs = cursor.fetchone()[0] or 0
                cursor.execute("SELECT COUNT(DISTINCT ie_name) FROM ie_definitions")
                total_ies = cursor.fetchone()[0] or 0

            self.stats_specs_lbl.setText(f"📦 {unique_specs} Spec{'s' if unique_specs != 1 else ''}")
            self.stats_vers_lbl.setText(f"🏷️ {total_versions} Version{'s' if total_versions != 1 else ''}")
            self.stats_msgs_lbl.setText(f"📨 {total_msgs} Messages · {total_ies} IEs")
        except Exception as e:
            logger.debug(f"Could not update stats ribbon: {e}")

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
            self.msg_search.setPlaceholderText("Filter IE / Type name, acronym, or Clause (e.g. UE parameters, 9.11.3)...")
            self.ie_search.setPlaceholderText(
                "Filter IE by keyword or Description..."
                if checked
                else "Additional IE / Clause keyword filter..."
            )
        else:
            self.msg_search.setPlaceholderText("Filter by name, acronym, or spec (e.g. SIB1, PFCP, 24.501)...")
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

        self._update_stats_ribbon()

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

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._save_config()
        self._populate_messages()

        current_msg_item = self.msg_list.currentItem()
        if current_msg_item:
            self._on_message_clicked(current_msg_item)
        else:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)

    def _delete_single_version(self, spec_number: str, version: str):
        delete_single_version(self, spec_number, version)

    def _delete_spec_group(self, spec_number: str, child_count: int):
        delete_spec_group(self, spec_number, child_count)

    def _on_search_timer_timeout(self):
        self._save_config()
        self._populate_messages()

    def _populate_messages(self):
        if not self.selected_version_ids:
            self.msg_list.clear()
            self.msg_count_lbl.setText("0 items")
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            default_prompt = (
                "Select an Information Element to View Evolution && Definition"
                if self._browse_mode == "ies"
                else "Select a Message to View Evolution Matrix"
            )
            self.matrix_title.setText(default_prompt)
            self.inspector.clear_display()
            return

        ie_query = self.ie_search.text().strip()
        if self._browse_mode == "ies":
            top_query = self.msg_search.text().strip()
            combined_query = f"{top_query} {ie_query}".strip()
        else:
            combined_query = ie_query

        search_desc = self.deep_search_btn.isChecked()

        self._msg_query_req_id += 1
        req_id = self._msg_query_req_id

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
        visible_count = 0

        if self._browse_mode == "ies":
            target_key = self.current_selected_ie_key
            for d in items:
                ie_name = d.get("ie_name", "")
                clause = d.get("clause", "")
                spec_num = str(d.get("spec_number", "")).replace("TS", "").strip()
                meta = get_spec_metadata(spec_num) if spec_num else {}

                ie_key = f"{ie_name}::{clause}"

                item = QListWidgetItem(ie_name)
                item.setData(Qt.UserRole, ie_key)
                item.setData(Qt.UserRole + 1, d)
                item.setData(Qt.UserRole + 2, spec_num)
                item.setData(Qt.UserRole + 3, clause)
                item.setData(Qt.UserRole + 4, meta.get("acronym", ""))
                item.setData(Qt.UserRole + 5, meta.get("wg", ""))

                tooltip_lines = [f"Information Element: {ie_name}"]
                if clause:
                    tooltip_lines.append(f"Clause / Type: {clause}")
                if spec_num:
                    tooltip_lines.append(f"Specification: TS {spec_num} ({meta.get('title', '')})")
                    tooltip_lines.append(f"Working Group: {meta.get('wg', '')} • Domain: {meta.get('tech', '')}")
                item.setToolTip("\n".join(tooltip_lines))

                self.msg_list.addItem(item)

                search_target = f"{ie_name} {clause} {spec_num} {meta.get('acronym', '')} {meta.get('title', '')}".lower()
                if msg_tokens and not all(tok in search_target for tok in msg_tokens):
                    item.setHidden(True)
                else:
                    visible_count += 1

                if target_key and ie_key == target_key and not item.isHidden():
                    target_item = item
        else:
            target_msg_name = self.current_selected_message_name
            for m in items:
                msg_name = m["message_name"]
                clause = m.get("clause", "")
                spec_num = str(m.get("spec_number", "")).replace("TS", "").strip()
                meta = get_spec_metadata(spec_num) if spec_num else {}

                item = QListWidgetItem(msg_name)
                item.setData(Qt.UserRole, msg_name)
                item.setData(Qt.UserRole + 1, m)
                item.setData(Qt.UserRole + 2, spec_num)
                item.setData(Qt.UserRole + 3, clause)
                item.setData(Qt.UserRole + 4, meta.get("acronym", ""))
                item.setData(Qt.UserRole + 5, meta.get("wg", ""))

                tooltip_lines = [f"Message: {msg_name}"]
                if clause:
                    tooltip_lines.append(f"Clause: {clause}")
                if spec_num:
                    tooltip_lines.append(f"Specification: TS {spec_num} ({meta.get('title', '')})")
                    tooltip_lines.append(f"Working Group: {meta.get('wg', '')} • Domain: {meta.get('tech', '')}")
                item.setToolTip("\n".join(tooltip_lines))

                self.msg_list.addItem(item)

                search_target = f"{msg_name} {clause} {spec_num} {meta.get('acronym', '')} {meta.get('title', '')}".lower()
                if msg_tokens and not all(tok in search_target for tok in msg_tokens):
                    item.setHidden(True)
                else:
                    visible_count += 1

                if target_msg_name and msg_name == target_msg_name and not item.isHidden():
                    target_item = item

        self.msg_list.blockSignals(False)

        # Update counter label
        total_count = self.msg_list.count()
        if visible_count == total_count:
            self.msg_count_lbl.setText(f"{total_count} items")
        else:
            self.msg_count_lbl.setText(f"{visible_count} / {total_count} items")

        default_prompt = (
            "Select an Information Element to View Evolution && Definition"
            if self._browse_mode == "ies"
            else "Select a Message to View Evolution Matrix"
        )

        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        elif self.msg_list.count() == 1 and (msg_tokens or self.ie_search.text().strip()):
            only_item = self.msg_list.item(0)
            if not only_item.isHidden():
                self.msg_list.setCurrentItem(only_item)
                self._on_message_clicked(only_item)
        elif self.msg_list.count() > 0:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self.matrix_title.setText(default_prompt)
        else:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
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

            meta = get_spec_metadata(self._current_message_spec) if self._current_message_spec else {}
            spec_badge = f" [{meta.get('acronym', f'TS {self._current_message_spec}')}]" if self._current_message_spec else ""
            clause_suffix = f" [Clause {clause}]" if clause and clause != ie_name else ""
            self.matrix_title.setText(f"Information Element{spec_badge}: {ie_name}{clause_suffix}")
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

        meta = get_spec_metadata(self._current_message_spec) if self._current_message_spec else {}
        spec_badge = f" [{meta.get('acronym', f'TS {self._current_message_spec}')}]" if self._current_message_spec else ""
        self.matrix_title.setText(f"Message{spec_badge}: {msg_name}{title_suffix}")

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
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
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
        has_rows = model.rowCount() > 0
        self.matrix_stack.setCurrentIndex(1 if has_rows else 0)
        self.copy_matrix_btn.setEnabled(has_rows)

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
        fetch_from_specs_db_clicked(self)

    def _on_import_local_file_clicked(self):
        import_local_file_clicked(self)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        start_batch_ingestion(self, tasks)

    def _on_import_error(self, err_msg: str):
        on_import_error(self, err_msg)

    def _on_import_progress(self, msg: str, val: int):
        on_import_progress(self, msg, val)

    def _on_import_success(self, spec_count: int, msg_count: int):
        on_import_success(self, spec_count, msg_count)

    def _on_clear_version_clicked(self):
        clear_version_clicked(self)

    def _on_wipe_db_clicked(self):
        wipe_db_clicked(self)

    def _on_import_cross_ref_spec_requested(self, spec_number: str, major_version: int):
        import_cross_ref_spec_requested(self, spec_number, major_version)

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
