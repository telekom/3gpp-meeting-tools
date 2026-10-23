"""
Main Tab for Specification Substring Search, Release Evolution Tracking, and Cutoff Date Analysis.
Visualizes multi-specification search results in dedicated per-specification tabs with automatic
persistence of active filters and selected specification versions.
"""

import html
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QDate, Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_WARNING,
    BUTTON_STYLE_TOOLBAR_DANGER,
)
from core.utils.paths import get_project_root, get_specs_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.spec_search.core.spec_clause_exporter import (
    SpecClauseWordExtractorThread,
    copy_clause_to_clipboard,
    export_clause_to_markdown_file,
    get_default_clause_export_path,
)
from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_search_threads import (
    SpecSearchImportThread,
    SpecSearchQueryWorker,
    SpecSearchWipeWorker,
    is_change_mark_file,
)
from modules.spec_search.ui.spec_search_components import (
    SpecClauseInspector,
    SpecEvolutionStatusDelegate,
    SpecSearchVersionTreeWidget,
)
from modules.spec_search.ui.spec_search_dialogs import (
    SpecClauseDiffDialog,
    SpecSearchVersionSelectDialog,
)
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase

logger = logging.getLogger(__name__)


class SpecSearchTab(QWidget):
    """Main Search Tab: Tree, Query Inputs, Cutoff Date Filter, Tabbed Evolution Matrix, and Inspector."""

    log_msg = pyqtSignal(str, int)
    navigate_to_specs_tab_requested = pyqtSignal(str, str)  # spec_number, version

    def __init__(self, search_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.search_db_path = Path(search_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = get_project_root() / "config" / "spec_search_config.json"

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = get_specs_root()

        self.db = SpecSearchDatabase(self.search_db_path)
        self.specs_db = SpecsDatabase(self.specs_db_path) if self.specs_db_path and self.specs_db_path.exists() else None
        self._spec_titles: Dict[str, str] = {}
        self._load_spec_titles()

        self._query_worker: Optional[SpecSearchQueryWorker] = None
        self._active_query_workers: Dict[int, SpecSearchQueryWorker] = {}
        self._current_req_id: int = 0
        self.selected_version_ids: List[int] = []

        # Debounce timer for search execution
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self._execute_search)

        # Debounce timer for saving filter/selection state
        self._save_config_timer = QTimer(self)
        self._save_config_timer.setSingleShot(True)
        self._save_config_timer.setInterval(500)
        self._save_config_timer.timeout.connect(self._save_config)

        self._setup_ui()
        self._load_config_and_refresh()

    def _load_spec_titles(self):
        """Loads specification numbers and formal titles from SpecsDatabase into memory cache."""
        if not self.specs_db:
            default_path = get_project_root() / "db" / "3gpp_data.db"
            if default_path.exists():
                self.specs_db = SpecsDatabase(default_path)

        if self.specs_db:
            try:
                self._spec_titles = self.specs_db.get_all_spec_titles()
            except Exception as e:
                logging.getLogger(__name__).warning(f"Could not load spec titles: {e}")
                self._spec_titles = {}

    def get_spec_title(self, spec_number: str) -> str:
        """Returns the formal specification title (e.g. 'System architecture for the 5G System (5GS)')."""
        if not spec_number:
            return ""
        clean = re.sub(r'^(TS|TR)\s*', '', spec_number.strip(), flags=re.IGNORECASE)
        if clean in self._spec_titles:
            return self._spec_titles[clean]
        # Check base number without sub-part (e.g. 23.801-01 -> 23.801)
        base = clean.split('-')[0]
        if base in self._spec_titles:
            return self._spec_titles[base]
        # Fallback to PROTOCOL_SPEC_CATALOG if available
        try:
            from modules.nas.ui.nas_dialogs import PROTOCOL_SPEC_CATALOG
            if clean in PROTOCOL_SPEC_CATALOG:
                return PROTOCOL_SPEC_CATALOG[clean].get("full_title", "")
        except Exception:
            pass
        return ""

    def get_spec_short_name(self, spec_number: str) -> str:
        """Returns a concise acronym or short name for a spec (e.g. '5GS NAS', 'NR RRC', 'UAS', '5GS Architecture')."""
        if not spec_number:
            return ""
        clean = re.sub(r'^(TS|TR)\s*', '', spec_number.strip(), flags=re.IGNORECASE)
        # Check PROTOCOL_SPEC_CATALOG first for gold-standard curated telecom acronyms
        try:
            from modules.nas.ui.nas_dialogs import PROTOCOL_SPEC_CATALOG
            if clean in PROTOCOL_SPEC_CATALOG:
                return PROTOCOL_SPEC_CATALOG[clean].get("acronym", "")
        except Exception:
            pass

        title = self.get_spec_title(clean)
        if not title:
            return ""

        # Check for parenthesized acronym e.g. '(UAS)', '(5GS)', '(V2X)', '(ProSe)'
        acronyms = re.findall(r'\(([A-Z0-9\-_]{2,8})\)', title)
        for acr in acronyms:
            if acr.upper() not in ("TS", "TR", "STAGE 2", "STAGE 3", "REL"):
                return acr

        # Fallback to concise title (first phrase before semicolon)
        s = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', title, flags=re.IGNORECASE)
        parts = [p.strip() for p in s.split(';') if p.strip()]
        first_part = parts[0] if parts else title
        if len(first_part) > 22:
            return first_part[:21].rstrip() + "…"
        return first_part

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # Main Splitter Layout (Left: Version Tree, Right: Search, Tabs & Inspector)
        # ---------------------------------------------------------------------
        self.main_splitter = QSplitter(Qt.Horizontal)

        # Left Panel: Versions Tree inside a Card Frame
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        left_card = QFrame()
        left_card.setObjectName("specTreeCard")
        left_card.setStyleSheet("""
            QFrame#specTreeCard {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        card_layout = QVBoxLayout(left_card)
        card_layout.setContentsMargins(6, 8, 6, 6)
        card_layout.setSpacing(6)

        # Card Header: Title + Count Badge + Compact Import & Maintenance Dropdowns
        tree_hdr_layout = QHBoxLayout()
        tree_hdr_layout.setContentsMargins(2, 0, 2, 0)
        tree_hdr_layout.setSpacing(4)
        tree_title = QLabel("📚 Indexed Specs")
        tree_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B;")
        tree_hdr_layout.addWidget(tree_title)

        self.tree_stats_badge = QLabel("0")
        self.tree_stats_badge.setStyleSheet("""
            background-color: #F1F5F9;
            color: #475569;
            font-size: 10px;
            font-weight: bold;
            border: 1px solid #E2E8F0;
            border-radius: 4px;
            padding: 1px 5px;
        """)
        tree_hdr_layout.addWidget(self.tree_stats_badge)
        tree_hdr_layout.addStretch()

        # 1. Compact Import Menu Button
        self.import_menu_btn = QPushButton("📥 Import ▾")
        self.import_menu_btn.setCursor(Qt.PointingHandCursor)
        self.import_menu_btn.setToolTip("Import specifications from DB or local Word .docx")
        self.import_menu_btn.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 7px;
                font-size: 11px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
            QPushButton::menu-indicator {
                image: none;
                width: 0px;
            }
        """)
        import_menu = QMenu(self)
        act_import_db = import_menu.addAction("📥 Import from Specifications DB...")
        act_import_db.triggered.connect(self._on_fetch_clicked)
        act_import_docx = import_menu.addAction("📁 Import Local Word Document (.docx)...")
        act_import_docx.triggered.connect(self._on_import_local_clicked)
        self.import_menu_btn.setMenu(import_menu)
        tree_hdr_layout.addWidget(self.import_menu_btn)

        # 2. Compact Database Management Menu Button
        self.db_menu_btn = QPushButton("⚙️ ▾")
        self.db_menu_btn.setCursor(Qt.PointingHandCursor)
        self.db_menu_btn.setToolTip("Search database maintenance (clear versions, wipe DB)")
        self.db_menu_btn.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 11px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
            QPushButton::menu-indicator {
                image: none;
                width: 0px;
            }
        """)
        db_menu = QMenu(self)
        self.act_clear_ver = db_menu.addAction("🗑️ Clear Specific Version...")
        self.act_clear_ver.triggered.connect(self._on_clear_version_clicked)
        db_menu.addSeparator()
        self.act_wipe_db = db_menu.addAction("⚠️ Wipe Entire Search Database...")
        self.act_wipe_db.triggered.connect(self._on_wipe_db_clicked)
        self.db_menu_btn.setMenu(db_menu)
        tree_hdr_layout.addWidget(self.db_menu_btn)

        # Aliases for background workers and existing callbacks
        self.fetch_btn = self.import_menu_btn
        self.import_local_btn = self.import_menu_btn
        self.clear_ver_btn = self.db_menu_btn
        self.wipe_db_btn = self.db_menu_btn
        self.db_stats_label = self.tree_stats_badge

        card_layout.addLayout(tree_hdr_layout)

        # Progress bar for imports (sleek 4px inline bar inside left card)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #E2E8F0;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
                border-radius: 2px;
            }
        """)
        self.progress_bar.setVisible(False)
        card_layout.addWidget(self.progress_bar)

        # Quick Filter + Select All / None Actions
        tree_actions_layout = QHBoxLayout()
        tree_actions_layout.setSpacing(4)
        self.tree_filter_input = QLineEdit()
        self.tree_filter_input.setPlaceholderText("Filter specs (e.g. 23.501)...")
        self.tree_filter_input.setClearButtonEnabled(True)
        self.tree_filter_input.setStyleSheet("""
            QLineEdit {
                padding: 2px 6px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-size: 11px;
                background-color: #F8FAFC;
            }
            QLineEdit:focus {
                border: 1px solid #0284C7;
                background-color: #FFFFFF;
            }
        """)
        tree_actions_layout.addWidget(self.tree_filter_input, stretch=1)

        btn_all = QPushButton("All")
        btn_all.setToolTip("Select all indexed releases")
        btn_all.setCursor(Qt.PointingHandCursor)
        btn_all.setStyleSheet("""
            QPushButton {
                font-size: 10px;
                font-weight: 600;
                color: #334155;
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 3px;
                padding: 2px 6px;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        tree_actions_layout.addWidget(btn_all)

        btn_none = QPushButton("None")
        btn_none.setToolTip("Deselect all indexed releases")
        btn_none.setCursor(Qt.PointingHandCursor)
        btn_none.setStyleSheet("""
            QPushButton {
                font-size: 10px;
                font-weight: 600;
                color: #334155;
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 3px;
                padding: 2px 6px;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        tree_actions_layout.addWidget(btn_none)
        card_layout.addLayout(tree_actions_layout)

        self.version_tree = SpecSearchVersionTreeWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        self.tree_filter_input.textChanged.connect(self.version_tree.filter_tree)
        btn_all.clicked.connect(self.version_tree.select_all)
        btn_none.clicked.connect(self.version_tree.deselect_all)
        card_layout.addWidget(self.version_tree)

        left_layout.addWidget(left_card)
        self.main_splitter.addWidget(left_widget)

        # Right Panel: Search Controls, Tabbed Evolution Matrix, Inspector
        self.right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)
        matrix_layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # Search & Filter Console Card
        # ---------------------------------------------------------------------
        search_card = QFrame()
        search_card.setObjectName("searchCard")
        search_card.setStyleSheet("""
            QFrame#searchCard {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        search_card_layout = QVBoxLayout(search_card)
        search_card_layout.setContentsMargins(8, 8, 8, 8)
        search_card_layout.setSpacing(6)

        # Row 1: Substring Search & Clause Filter
        search_bar_layout = QHBoxLayout()
        search_bar_layout.setSpacing(6)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search exact substring or phrase (e.g. 'No-Transmit Zone', 'emergency', 'PDU session')...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setStyleSheet("""
            QLineEdit {
                padding: 5px 10px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-size: 13px;
                background-color: #FFFFFF;
                color: #0F172A;
            }
            QLineEdit:focus {
                border: 1.5px solid #0284C7;
            }
        """)
        self.search_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.search_input, stretch=1)

        self.clause_filter_input = QLineEdit()
        self.clause_filter_input.setPlaceholderText("Clause (e.g. 5.2)...")
        self.clause_filter_input.setFixedWidth(150)
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.setStyleSheet("""
            QLineEdit {
                padding: 5px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-size: 12px;
                background-color: #FFFFFF;
            }
            QLineEdit:focus {
                border: 1.5px solid #0284C7;
            }
        """)
        self.clause_filter_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.clause_filter_input)
        search_card_layout.addLayout(search_bar_layout)

        # Row 2: Cutoff Date Filter + Post-Cutoff Checkbox
        cutoff_date_bar = QHBoxLayout()
        cutoff_date_bar.setSpacing(8)

        self.chk_filing_date = QCheckBox("🎯 Date Cutoff:")
        self.chk_filing_date.setStyleSheet("font-weight: 600; font-size: 11px; color: #334155;")
        self.chk_filing_date.toggled.connect(self._on_cutoff_date_filter_toggled)
        cutoff_date_bar.addWidget(self.chk_filing_date)

        self.cutoff_date_edit = QDateEdit()
        self.cutoff_date_edit.setCalendarPopup(True)
        self.cutoff_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.cutoff_date_edit.setDate(QDate.currentDate().addYears(-3))
        self.cutoff_date_edit.setMinimumWidth(125)
        self.cutoff_date_edit.setEnabled(False)
        self.cutoff_date_edit.setStyleSheet("""
            QDateEdit {
                padding: 3px 6px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-size: 12px;
                background-color: #FFFFFF;
                color: #1E293B;
            }
            QDateEdit:disabled {
                background-color: #F8FAFC;
                color: #94A3B8;
            }
        """)
        self.cutoff_date_edit.dateChanged.connect(self._on_filter_changed)
        cutoff_date_bar.addWidget(self.cutoff_date_edit)

        self.chk_post_date_only = QCheckBox("Show Only Post-Cutoff Additions")
        self.chk_post_date_only.setToolTip("Hides clauses where matching text was already present prior to the selected date.")
        self.chk_post_date_only.setStyleSheet("font-size: 11px; color: #475569;")
        self.chk_post_date_only.setEnabled(False)
        self.chk_post_date_only.toggled.connect(self._on_filter_changed)
        cutoff_date_bar.addWidget(self.chk_post_date_only)

        cutoff_date_bar.addStretch()
        search_card_layout.addLayout(cutoff_date_bar)

        matrix_layout.addWidget(search_card)

        # Executive Results Status Pill Bar
        self.matrix_title = QLabel("💡 Type a query above to trace clause evolution across releases.")
        self.matrix_title.setStyleSheet("font-size: 12px; color: #64748B; margin-top: 2px; margin-bottom: 2px;")
        matrix_layout.addWidget(self.matrix_title)

        # Tabbed Container for Per-Specification Matrices
        self.spec_results_tabs = QTabWidget()
        self.spec_results_tabs.setDocumentMode(True)
        self.spec_results_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #F1F5F9;
                color: #475569;
                font-weight: 600;
                font-size: 11px;
                border: 1px solid #CBD5E1;
                border-bottom: none;
                padding: 5px 14px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #0284C7;
                font-weight: bold;
                border-bottom: 2px solid #0284C7;
            }
            QTabBar::tab:hover:!selected {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        matrix_layout.addWidget(self.spec_results_tabs)
        self.right_splitter.addWidget(matrix_widget)

        self.inspector = SpecClauseInspector()
        self.inspector.compare_clause_requested.connect(self._open_clause_diff_dialog)
        self.inspector.open_document_requested.connect(self._handle_open_spec_document)
        self.inspector.show_in_specs_tab_requested.connect(self._handle_show_in_specs_tab)
        self.inspector.quick_diff_requested.connect(self._handle_quick_diff)
        self.inspector.jump_to_version_requested.connect(self._handle_jump_to_version)
        self.inspector.extract_clause_requested.connect(self._handle_extract_clause_word)
        self.inspector.export_markdown_requested.connect(self._handle_export_clause_markdown)
        self.inspector.layout_orientation_toggle_requested.connect(self._toggle_inspector_layout)
        self.inspector.maximize_toggle_requested.connect(self._toggle_inspector_maximize)
        self.right_splitter.addWidget(self.inspector)

        self.right_splitter.setSizes([380, 260])
        self.main_splitter.addWidget(self.right_splitter)
        self.main_splitter.setSizes([300, 750])
        layout.addWidget(self.main_splitter)

    # -------------------------------------------------------------------------
    # Configuration Persistence
    # -------------------------------------------------------------------------

    def _save_config(self):
        """Saves active search text, clause filter, cutoff dates, and checked specification versions."""
        try:
            checked_versions = self.version_tree.get_checked_versions_info()
            config_data = {
                "search_query": self.search_input.text(),
                "clause_filter": self.clause_filter_input.text(),
                "enable_cutoff": self.chk_filing_date.isChecked(),
                "cutoff_date": self.cutoff_date_edit.date().toString("yyyy-MM-dd"),
                "only_post_cutoff": self.chk_post_date_only.isChecked(),
                "checked_versions": checked_versions,
                "inspector_side_by_side": (self.right_splitter.orientation() == Qt.Horizontal),
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to save spec search config: {e}")

    def _load_config(self) -> dict:
        """Loads saved specification search settings from JSON."""
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"Failed to read spec search config: {e}")
        return {}

    def _load_config_and_refresh(self):
        """Restores UI filter inputs, populates the tree with saved checkboxes, and runs search."""
        config = self._load_config()
        versions = self.db.get_imported_versions()

        saved_checked = config.get("checked_versions", None) if config else None
        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()

        if config:
            # Block signals while restoring to avoid redundant triggers
            self.search_input.blockSignals(True)
            self.clause_filter_input.blockSignals(True)
            self.chk_filing_date.blockSignals(True)
            self.cutoff_date_edit.blockSignals(True)
            self.chk_post_date_only.blockSignals(True)

            self.search_input.setText(config.get("search_query", ""))
            self.clause_filter_input.setText(config.get("clause_filter", ""))

            enable_cutoff = config.get("enable_cutoff", False)
            self.chk_filing_date.setChecked(enable_cutoff)
            self.cutoff_date_edit.setEnabled(enable_cutoff)
            self.chk_post_date_only.setEnabled(enable_cutoff)

            saved_date = config.get("cutoff_date", "")
            if saved_date:
                d = QDate.fromString(saved_date, "yyyy-MM-dd")
                if d.isValid():
                    self.cutoff_date_edit.setDate(d)

            self.chk_post_date_only.setChecked(config.get("only_post_cutoff", False))

            self.search_input.blockSignals(False)
            self.clause_filter_input.blockSignals(False)
            self.chk_filing_date.blockSignals(False)
            self.cutoff_date_edit.blockSignals(False)
            self.chk_post_date_only.blockSignals(False)

            if config.get("inspector_side_by_side", False):
                self.inspector.set_side_by_side(True)
                self.right_splitter.setOrientation(Qt.Horizontal)
                self.right_splitter.setSizes([550, 450])

        self._update_db_stats()

        # Trigger search if a saved query exists
        if self.search_input.text().strip():
            self._execute_search()

    def _toggle_inspector_layout(self, side_by_side: bool):
        if side_by_side:
            self.right_splitter.setOrientation(Qt.Horizontal)
            self.right_splitter.setSizes([550, 450])
        else:
            self.right_splitter.setOrientation(Qt.Vertical)
            self.right_splitter.setSizes([380, 260])
        self._save_config()

    def _toggle_inspector_maximize(self, is_maximized: bool):
        if is_maximized:
            self._pre_maximize_sizes = self.right_splitter.sizes()
            self.right_splitter.setSizes([0, 1000])
        else:
            if hasattr(self, "_pre_maximize_sizes") and self._pre_maximize_sizes and any(self._pre_maximize_sizes):
                self.right_splitter.setSizes(self._pre_maximize_sizes)
            else:
                if self.right_splitter.orientation() == Qt.Horizontal:
                    self.right_splitter.setSizes([550, 450])
                else:
                    self.right_splitter.setSizes([380, 260])

    def _update_db_stats(self):
        try:
            versions = self.db.get_imported_versions()
            specs = set(v.get("spec_number") for v in versions if v.get("spec_number"))
            num_specs = len(specs)
            num_vers = len(versions)
            if hasattr(self, "tree_stats_badge"):
                self.tree_stats_badge.setText(f"{num_specs}")
                self.tree_stats_badge.setToolTip(f"{num_specs} Specification(s), {num_vers} Release(s) Indexed")
            if hasattr(self, "db_stats_label") and self.db_stats_label != self.tree_stats_badge:
                self.db_stats_label.setText(f"📦 {num_specs} Specs · {num_vers} Releases Indexed")
        except Exception:
            if hasattr(self, "tree_stats_badge"):
                self.tree_stats_badge.setText("0")

    def _on_filter_changed(self):
        self._search_timer.start()
        self._save_config_timer.start()

    def _on_cutoff_date_filter_toggled(self, checked: bool):
        self.cutoff_date_edit.setEnabled(checked)
        self.chk_post_date_only.setEnabled(checked)
        self._on_filter_changed()

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._search_timer.start()
        self._save_config_timer.start()

    # -------------------------------------------------------------------------
    # Search Execution
    # -------------------------------------------------------------------------

    def _execute_search(self):
        query = self.search_input.text().strip()
        clause_filter = self.clause_filter_input.text().strip()

        if not self.selected_version_ids or not query:
            self.spec_results_tabs.clear()
            self.matrix_title.setText("<span style='color: #64748B; font-size: 12px;'>💡 Type a query above to trace clause evolution across releases.</span>")
            self.inspector.clear_display()
            return

        self._current_req_id += 1
        req_id = self._current_req_id

        for w in list(self._active_query_workers.values()):
            if w.isRunning():
                w.requestInterruption()

        worker = SpecSearchQueryWorker(
            db=self.db,
            query_str=query,
            version_ids=self.selected_version_ids,
            clause_filter=clause_filter,
            request_id=req_id,
        )
        self._query_worker = worker
        self._active_query_workers[req_id] = worker
        worker.results_ready.connect(self._on_search_results_ready)
        worker.finished.connect(lambda r=req_id: self._release_query_worker(r))
        worker.start()

    def _release_query_worker(self, req_id: int):
        worker = self._active_query_workers.pop(req_id, None)
        if worker:
            worker.deleteLater()

    def _on_search_results_ready(self, df, req_id: int):
        if req_id != self._current_req_id:
            return

        query = self.search_input.text().strip()
        self.spec_results_tabs.clear()

        if df.empty:
            self.matrix_title.setText(f"<span style='color: #64748B; font-size: 12px;'>No matching clauses found for '<b>{html.escape(query)}</b>' across selected releases.</span>")
            self.inspector.clear_display()
            return

        cutoff_date = self.cutoff_date_edit.date().toString("yyyy-MM-dd") if self.chk_filing_date.isChecked() else None
        only_post_date = self.chk_post_date_only.isChecked() if self.chk_filing_date.isChecked() else False

        total_clauses_count = 0
        specs_with_hits = 0

        grouped_specs = sorted(df["spec_number"].unique())

        for spec_num in grouped_specs:
            spec_df = df[df["spec_number"] == spec_num].copy()

            model = SpecEvolutionMatrixModel(
                raw_df=spec_df,
                cutoff_date=cutoff_date,
                only_added_after_cutoff=only_post_date,
            )

            if model.rowCount() == 0:
                continue

            specs_with_hits += 1
            total_clauses_count += model.rowCount()

            table = QTableView()
            table.setAlternatingRowColors(False)
            table.setSelectionBehavior(QTableView.SelectRows)
            table.setSelectionMode(QTableView.SingleSelection)
            table.verticalHeader().setDefaultSectionSize(26)
            table.verticalHeader().setVisible(False)
            table.setItemDelegate(SpecEvolutionStatusDelegate(table))
            table.setModel(model)
            table.setStyleSheet("""
                QTableView {
                    background-color: #FFFFFF;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    gridline-color: transparent;
                    font-size: 12px;
                    selection-background-color: #EBF3FC;
                    selection-color: #0F172A;
                }
                QHeaderView::section {
                    background-color: #F8FAFC;
                    color: #475569;
                    font-weight: bold;
                    font-size: 11px;
                    border: none;
                    border-bottom: 1px solid #CBD5E1;
                    border-right: 1px solid #E2E8F0;
                    padding: 5px 8px;
                }
            """)

            h_header = table.horizontalHeader()
            h_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            h_header.setSectionResizeMode(1, QHeaderView.Stretch)
            for c in range(2, model.columnCount()):
                h_header.setSectionResizeMode(c, QHeaderView.Interactive)
                table.setColumnWidth(c, 115)

            table.clicked.connect(lambda idx, s=spec_num, t=table: self._on_tab_table_clicked(idx, s, t))
            table.doubleClicked.connect(lambda idx, s=spec_num, t=table: self._on_table_double_clicked(idx, s, t))
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(lambda pos, t=table, s=spec_num: self._show_matrix_context_menu(pos, t, s))

            short_name = self.get_spec_short_name(spec_num)
            tab_title = f"TS {spec_num} · {short_name} ({model.rowCount()})" if short_name else f"TS {spec_num} ({model.rowCount()})"
            tab_idx = self.spec_results_tabs.addTab(table, tab_title)
            full_title = self.get_spec_title(spec_num)
            if full_title:
                self.spec_results_tabs.setTabToolTip(tab_idx, f"TS {spec_num}: {full_title} ({model.rowCount()} clauses)")
            else:
                self.spec_results_tabs.setTabToolTip(tab_idx, f"TS {spec_num} ({model.rowCount()} clauses)")

        if specs_with_hits == 0:
            self.matrix_title.setText(f"<span style='color: #64748B; font-size: 12px;'>No matching clauses found for '<b>{html.escape(query)}</b>' matching active filters.</span>")
            self.inspector.clear_display()
            return

        date_badge = (
            f" <span style='background-color: #FEF3C7; color: #B45309; font-weight: bold; padding: 2px 7px; border-radius: 4px; border: 1px solid #FDE68A; font-size: 11px;'>Cutoff ≥ {cutoff_date}</span>"
            if cutoff_date and only_post_date
            else ""
        )
        self.matrix_title.setText(
            f"<span style='color: #475569; font-size: 11px; font-weight: 600;'>Results:</span> "
            f"<span style='background-color: #EBF3FC; color: #0284C7; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid #BEE3F8; font-size: 11px;'>🔍 {html.escape(query)}</span> "
            f"<span style='background-color: #DCFCE7; color: #15803D; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid #86EFAC; font-size: 11px;'>{total_clauses_count} Clauses</span> "
            f"<span style='background-color: #F1F5F9; color: #334155; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid #CBD5E1; font-size: 11px;'>{specs_with_hits} Spec(s)</span>"
            f"{date_badge}"
        )

    def _on_tab_table_clicked(self, index, spec_num: str, table: QTableView):
        model = table.model()
        if not model:
            return

        row, col = index.row(), index.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        clause_title = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip()
        query = self.search_input.text().strip()

        # Resolve target version header
        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])

        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")

        # Resolve introductory version & preceding version
        first_added_ver_header = model.get_first_added_version(row)
        first_added_ver = None
        if first_added_ver_header:
            m = re.search(r"v([0-9\.]+)", first_added_ver_header)
            first_added_ver = m.group(1) if m else first_added_ver_header.lstrip("v")

        prev_ver_header = model.get_preceding_version(ver_header)
        prev_ver = None
        if prev_ver_header:
            m = re.search(r"v([0-9\.]+)", prev_ver_header)
            prev_ver = m.group(1) if m else prev_ver_header.lstrip("v")

        # 1. Fetch by direct SQLite primary key (clause_pk)
        clause_data = None
        target_pk = model.get_clause_pk(row, ver_header)
        if target_pk is None:
            target_pk = model.get_row_first_pk(row)

        if target_pk is not None:
            clause_data = self.db.get_clause_content(target_pk)

        # 2. Fallback to longest content query by spec/version/clause
        if not clause_data:
            clause_data = self.db.get_clause_content_by_spec_ver(spec_num, version, clause_num)

        if clause_data:
            spec_title = self.get_spec_short_name(spec_num) or self.get_spec_title(spec_num)
            self.inspector.display_clause(
                clause_number=clause_num,
                clause_title=clause_title,
                spec_number=clause_data.get("spec_number", spec_num),
                version=clause_data.get("version", version),
                release_date=clause_data.get("release_date"),
                content=clause_data.get("content", ""),
                search_query=query,
                first_added_ver=first_added_ver,
                prev_ver=prev_ver,
                spec_title=spec_title,
            )

    # -------------------------------------------------------------------------
    # Ingestion Actions
    # -------------------------------------------------------------------------

    def refresh_versions(self):
        versions = self.db.get_imported_versions()
        saved_checked = self.version_tree.get_checked_versions_info()
        if not saved_checked:
            cfg = self._load_config()
            saved_checked = cfg.get("checked_versions") if cfg else None

        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._update_db_stats()

    def _on_fetch_clicked(self):
        if not self.specs_db:
            QMessageBox.warning(self, "Specs DB Unavailable", "The specifications database (3gpp_data.db) is not configured.")
            return

        dialog = SpecSearchVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)
        if dialog.exec_() == dialog.Accepted and dialog.selected_files_info:
            self._start_batch_ingestion(dialog.selected_files_info)

    def _on_import_local_clicked(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Specification Document(s) (.docx)", "", "Word Files (*.docx)"
        )
        if not paths:
            return

        clean_paths = [p for p in paths if not is_change_mark_file(p)]
        if len(clean_paths) < len(paths):
            skipped = len(paths) - len(clean_paths)
            logger.log(logging.INFO, f"ℹ️ Skipped {skipped} revision mark (-rm) file(s).")

        grouped: Dict[str, List[Path]] = {}
        for fp in clean_paths:
            p = Path(fp)
            base_key = re.sub(r"_\d+_.*$", "", p.stem)
            base_key = re.sub(r"[-_]cl$", "", base_key, flags=re.IGNORECASE)
            grouped.setdefault(base_key, []).append(p)

        tasks = []
        for base_key, p_list in grouped.items():
            dispatcher = ProtocolDocxDispatcher(p_list)
            spec_num = dispatcher.extract_spec_number()
            version = dispatcher.extract_version_from_filename()
            tasks.append({
                "spec_number": spec_num,
                "version": version,
                "filename": f"{base_key}.docx",
                "file_url": "",
                "release_date": "",
                "local_docx_paths": p_list,
            })

        self._start_batch_ingestion(tasks)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        self._current_ingesting_tasks = tasks
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)

        self.thread = SpecSearchImportThread(
            search_db_path=self.search_db_path,
            tasks=tasks,
            cache_dir=self.cache_dir,
        )
        self.thread.progress.connect(self._on_import_progress)
        self.thread.finished_success.connect(self._on_import_success)
        self.thread.error.connect(self._on_import_error)
        self.thread.start()

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        logger.log(logging.INFO, msg)

    def _on_import_success(self, specs_count: int, clauses_count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        logger.log(logging.INFO, f"✅ Indexed {specs_count} specification(s) ({clauses_count} total clauses).")

        # Ensure newly ingested versions are preserved in saved_checked
        tasks = getattr(self, "_current_ingesting_tasks", [])
        saved_checked = self.version_tree.get_checked_versions_info()
        if saved_checked is None:
            saved_checked = []
        for t in tasks:
            s_num = str(t.get("spec_number", "")).strip()
            ver = str(t.get("version", "")).lstrip("vV").strip()
            if s_num and ver:
                if not any(
                    str(item.get("spec_number", "")).strip() == s_num
                    and str(item.get("version", "")).lstrip("vV").strip() == ver
                    for item in saved_checked
                ):
                    saved_checked.append({"spec_number": s_num, "version": ver})

        versions = self.db.get_imported_versions()
        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._save_config()

        # Focus newly indexed version if single spec
        if len(tasks) == 1:
            t = tasks[0]
            self.focus_spec(t.get("spec_number", ""), t.get("version", ""))

    def _on_import_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        QMessageBox.critical(self, "Indexing Error", err)
        logger.log(logging.ERROR, f"❌ Indexing error: {err}")

    def _delete_single_version(self, spec_number: str, version: str):
        if QMessageBox.question(self, "Confirm Delete", f"Delete TS {spec_number} v{version}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.clear_version(spec_number, version)
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _delete_spec_group(self, spec_number: str, count: int):
        if QMessageBox.question(self, "Confirm Delete", f"Delete all {count} version(s) of TS {spec_number}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in self.version_tree.get_all_versions_for_spec(spec_number):
                if item.get("version"):
                    self.db.clear_version(spec_number, item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_clear_version_clicked(self):
        checked = self.version_tree.get_checked_versions_info()
        if not checked:
            QMessageBox.warning(self, "Select Version", "Please check at least one version to clear.")
            return
        if QMessageBox.question(self, "Confirm Delete", f"Delete {len(checked)} checked version(s)?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in checked:
                self.db.clear_version(item["spec_number"], item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_wipe_db_clicked(self):
        if QMessageBox.critical(
            self,
            "Confirm Wipe",
            "Delete all indexed specification clauses and full-text indexes?",
            QMessageBox.Yes | QMessageBox.No,
        ) != QMessageBox.Yes:
            return

        for w in list(self._active_query_workers.values()):
            if w.isRunning():
                w.requestInterruption()
                w.wait(1500)

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.wipe_db_btn.setEnabled(False)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)
        self.clear_ver_btn.setEnabled(False)

        self._wipe_worker = SpecSearchWipeWorker(self.db)
        self._wipe_worker.finished_success.connect(self._on_wipe_db_success)
        self._wipe_worker.error.connect(self._on_wipe_db_error)
        self._wipe_worker.start()

    def _on_wipe_db_success(self):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        self.refresh_versions()
        self.spec_results_tabs.clear()
        self.inspector.clear_display()
        self.matrix_title.setText("Type a query above to see Release Evolution Matrix")
        self._save_config()
        logger.log(logging.INFO, "🧹 Specification Search DB wiped successfully.")

    def _on_wipe_db_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        QMessageBox.critical(self, "Wipe Error", f"Failed to wipe search database: {err}")
        logger.log(logging.ERROR, f"❌ Wipe error: {err}")

    def _open_clause_diff_dialog(
        self,
        spec_num: str,
        version: str,
        clause_num: str,
        clause_title: str,
        base_version: Optional[str] = None,
    ):
        """Opens the unified clause comparison & analysis dialog (Visual Diff, LLM prompt, Word Track Changes)."""
        dialog = SpecClauseDiffDialog(
            db=self.db,
            spec_number=spec_num,
            current_version=version,
            clause_number=clause_num,
            clause_title=clause_title,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            base_version=base_version,
            parent=self,
        )
        dialog.exec_()

    def _handle_open_spec_document(self, spec_number: str, version: str, clause_number: str, doc_format: str = "word"):
        """Dispatches specification document opening via SpecDocResolver."""
        SpecDocResolver.open_spec(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            doc_format=doc_format,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            parent_widget=self,
        )

    def _start_clause_word_compare(self, spec_num: str, base_ver: str, target_ver: str, clause_num: str, clause_title: str):
        """Extracts clause from both releases and executes Word comparison with Track Changes."""
        from modules.spec_search.core.spec_clause_word_comparator import SpecClauseWordCompareThread
        from PyQt5.QtWidgets import QProgressDialog

        prog = QProgressDialog(f"Comparing Clause {clause_num} in Microsoft Word (Track Changes)...", "Cancel", 0, 100, self)
        prog.setWindowTitle("Word Comparison (Track Changes)")
        prog.setWindowModality(1)
        prog.setValue(5)

        worker = SpecClauseWordCompareThread(
            spec_number=spec_num,
            base_version=base_ver,
            target_version=target_ver,
            clause_number=clause_num,
            clause_title=clause_title,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(msg):
            prog.close()
            logger.info(f"✅ {msg}")

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Word Comparison Failed",
                f"Could not generate Microsoft Word comparison with Track Changes:\n{err}\n\nTip: You can inspect the clause differences directly in the Diff & AI Analysis dialog."
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(lambda: self._active_word_workers.remove(worker) if worker in self._active_word_workers else None)

        worker.start()
        prog.exec_()

    def _handle_extract_clause_word(
        self,
        spec_number: str,
        version: str,
        clause_number: str,
        clause_title: str,
        content: str,
        save_as: bool = False,
    ):
        """Extracts clause into a standalone Word (.docx) document with native formatting and tables."""
        from PyQt5.QtWidgets import QFileDialog, QProgressDialog

        dest_path = None
        if save_as:
            default_path = get_default_clause_export_path(spec_number, version, clause_number, ext=".docx")
            chosen, _ = QFileDialog.getSaveFileName(
                self,
                f"Save Clause {clause_number} as Word Document",
                str(default_path),
                "Word Document (*.docx)",
            )
            if not chosen:
                return
            dest_path = Path(chosen)

        prog = QProgressDialog(
            f"Extracting Clause {clause_number} from TS {spec_number} v{version}...",
            "Cancel",
            0,
            100,
            self,
        )
        prog.setWindowTitle("Extracting Clause to Word")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        worker = SpecClauseWordExtractorThread(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            clause_title=clause_title,
            clause_content=content,
            output_path=dest_path,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(out_file):
            prog.close()
            logger.info(f"Clause excerpt created at: {out_file}")
            try:
                os.startfile(str(out_file))
            except Exception as e:
                logger.warning(f"Could not open file directly: {e}")
                QMessageBox.information(
                    self,
                    "Clause Extracted",
                    f"Clause {clause_number} extracted successfully to:\n{out_file}",
                )

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Clause Extraction Failed",
                f"Could not extract Clause {clause_number} in Word:\n{err}",
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(
            lambda: self._active_word_workers.remove(worker)
            if worker in self._active_word_workers
            else None
        )

        worker.start()
        prog.exec_()

    def _handle_export_clause_markdown(
        self,
        spec_number: str,
        version: str,
        clause_number: str,
        clause_title: str,
        content: str,
    ):
        """Exports the clause to a standalone Markdown (.md) document file with pictures and diagrams."""
        from PyQt5.QtWidgets import QFileDialog, QProgressDialog

        default_path = get_default_clause_export_path(spec_number, version, clause_number, ext=".md")
        chosen, _ = QFileDialog.getSaveFileName(
            self,
            f"Export Clause {clause_number} to Markdown",
            str(default_path),
            "Markdown File (*.md);;All Files (*)",
        )
        if not chosen:
            return

        dest_path = Path(chosen)
        prog = QProgressDialog(
            f"Exporting Clause {clause_number} to Markdown with diagrams...",
            "Cancel",
            0,
            100,
            self,
        )
        prog.setWindowTitle("Exporting Clause to Markdown")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        worker = SpecClauseWordExtractorThread(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            clause_title=clause_title,
            clause_content=content,
            output_path=dest_path,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            export_format="markdown",
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(out_file):
            prog.close()
            img_count = getattr(worker, "images_count", 0)
            img_msg = f" with {img_count} diagram(s)/image(s) extracted into '{out_file.stem}_media/'" if img_count > 0 else ""
            logger.info(f"Clause exported to Markdown at: {out_file}{img_msg}")
            reply = QMessageBox.question(
                self,
                "Clause Exported to Markdown",
                f"Clause {clause_number} exported successfully to:\n{out_file}{img_msg}\n\nDo you want to open the file?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                try:
                    os.startfile(str(out_file))
                except Exception as e:
                    logger.warning(f"Could not open file: {e}")

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Markdown Export Failed",
                f"Could not export Clause {clause_number} to Markdown:\n{err}",
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(
            lambda: self._active_word_workers.remove(worker)
            if worker in self._active_word_workers
            else None
        )

        worker.start()
        prog.exec_()

    def _handle_show_in_specs_tab(self, spec_number: str, version: str):
        """Emits signal to switch to Specifications tab with spec and version pre-selected."""
        self.navigate_to_specs_tab_requested.emit(spec_number, version)

    def _handle_quick_diff(self, spec_num: str, base_ver: str, target_ver: str, clause_num: str, clause_title: str):
        """Opens the clause comparison dialog pre-configured with base and target releases."""
        self._open_clause_diff_dialog(
            spec_num=spec_num,
            version=target_ver,
            clause_num=clause_num,
            clause_title=clause_title,
            base_version=base_ver,
        )

    def _handle_jump_to_version(self, spec_number: str, target_version: str, clause_number: str):
        """Jumps the inspector and matrix selection directly to the target version for this clause."""
        clean_target = target_version.lstrip("v").strip()
        for i in range(self.spec_results_tabs.count()):
            tab_title = self.spec_results_tabs.tabText(i)
            if f"TS {spec_number}" in tab_title:
                self.spec_results_tabs.setCurrentIndex(i)
                table = self.spec_results_tabs.widget(i)
                if isinstance(table, QTableView):
                    model = table.model()
                    if isinstance(model, SpecEvolutionMatrixModel):
                        for r in range(model.rowCount()):
                            if str(model.data(model.index(r, 0), Qt.DisplayRole)).strip() == clause_number:
                                target_col = 0
                                for c_idx, v_col in enumerate(model._version_cols):
                                    if clean_target in v_col:
                                        target_col = c_idx + 2
                                        break
                                target_idx = model.index(r, target_col)
                                table.setCurrentIndex(target_idx)
                                self._on_tab_table_clicked(target_idx, spec_number, table)
                                break
                break

    def _on_table_double_clicked(self, index, spec_num: str, table: QTableView):
        """Double-clicking a matrix cell directly opens the specification in Word."""
        model = table.model()
        if not model:
            return
        row, col = index.row(), index.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])
        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")
        self._handle_open_spec_document(spec_num, version, clause_num, "word")

    def _show_matrix_context_menu(self, pos, table: QTableView, spec_num: str):
        """Right-click context menu on Evolution Matrix cells."""
        idx = table.indexAt(pos)
        if not idx.isValid():
            return
        model = table.model()
        if not model:
            return

        row, col = idx.row(), idx.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        clause_title = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip()

        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])

        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")

        prev_ver_header = model.get_preceding_version(ver_header)
        prev_ver = None
        if prev_ver_header:
            m = re.search(r"v([0-9\.]+)", prev_ver_header)
            prev_ver = m.group(1) if m else prev_ver_header.lstrip("v")

        menu = QMenu(self)

        act_word = menu.addAction(f"📝 Open TS {spec_num} v{version} (Word)")
        act_word.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "word"))

        act_pdf = menu.addAction(f"📕 Open TS {spec_num} v{version} (PDF)")
        act_pdf.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "pdf"))

        act_folder = menu.addAction("📂 Open Local Folder in Explorer")
        act_folder.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "folder"))

        menu.addSeparator()

        act_specs = menu.addAction(f"🔍 Show TS {spec_num} v{version} in Specifications Tab")
        act_specs.triggered.connect(lambda: self._handle_show_in_specs_tab(spec_num, version))

        menu.addSeparator()

        if prev_ver:
            act_compare = menu.addAction(f"🔍 Compare Clause {clause_num} (Diff && AI Analysis vs v{prev_ver})...")
            act_compare.triggered.connect(lambda: self._open_clause_diff_dialog(spec_num, version, clause_num, clause_title, base_version=prev_ver))
        else:
            act_compare = menu.addAction(f"🔍 Compare Clause {clause_num} (Diff && AI Analysis)...")
            act_compare.triggered.connect(lambda: self._open_clause_diff_dialog(spec_num, version, clause_num, clause_title))

        menu.addSeparator()
        clause_rec = self.db.get_clause_content_by_spec_ver(spec_num, version, clause_num)
        clause_text = clause_rec.get("content", "") if clause_rec else ""
        rel_date = clause_rec.get("release_date", "") if clause_rec else ""
        actual_title = clause_rec.get("clause_title", clause_title) if clause_rec else clause_title

        menu.addSeparator()
        act_ext_word = menu.addAction(f"✂️ Extract Clause {clause_num} in Word (.docx)")
        act_ext_word.triggered.connect(
            lambda: self._handle_extract_clause_word(
                spec_num, version, clause_num, actual_title, clause_text, save_as=False
            )
        )
        act_save_word = menu.addAction(f"💾 Save Clause {clause_num} as Word (.docx)...")
        act_save_word.triggered.connect(
            lambda: self._handle_extract_clause_word(
                spec_num, version, clause_num, actual_title, clause_text, save_as=True
            )
        )
        act_export_md = menu.addAction(f"📘 Export Clause {clause_num} to Markdown (.md)...")
        act_export_md.triggered.connect(
            lambda: self._handle_export_clause_markdown(
                spec_num, version, clause_num, actual_title, clause_text
            )
        )

        menu.addSeparator()
        copy_menu = menu.addMenu("📋 Copy Clause")
        act_cp_email = copy_menu.addAction("📋 Copy for Email / Outlook (Rich Text)")
        act_cp_email.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="email"
            )
        )
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown (Quote Block)")
        act_cp_md.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="markdown"
            )
        )
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="text"
            )
        )
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="citation"
            )
        )
        act_copy = menu.addAction("📋 Copy Clause Citation")
        act_copy.triggered.connect(lambda: QApplication.clipboard().setText(f"3GPP TS {spec_num} v{version}, Clause {clause_num} ({clause_title})"))

        menu.exec_(table.viewport().mapToGlobal(pos))

    def focus_spec(self, spec_number: str, version: Optional[str] = None):
        """
        External API to focus on a specific specification and version from other tabs.
        Ensures the spec is selected in the version tree and search bar is focused.
        """
        clean_spec = spec_number.strip()
        clean_ver = str(version).lstrip("vV").strip() if version else None

        # Check if spec is in version tree
        all_specs = [
            self.version_tree.topLevelItem(i).data(0, Qt.UserRole).get("spec_number")
            for i in range(1, self.version_tree.topLevelItemCount())
            if self.version_tree.topLevelItem(i).data(0, Qt.UserRole)
        ]

        if clean_spec in all_specs:
            target_spec_item = None
            target_child_item = None
            for i in range(1, self.version_tree.topLevelItemCount()):
                item = self.version_tree.topLevelItem(i)
                data = item.data(0, Qt.UserRole) or {}
                if data.get("spec_number") == clean_spec:
                    target_spec_item = item
                    item.setExpanded(True)
                    if clean_ver:
                        for c_idx in range(item.childCount()):
                            child = item.child(c_idx)
                            c_data = child.data(0, Qt.UserRole) or {}
                            if str(c_data.get("version", "")).lstrip("vV").strip() == clean_ver:
                                target_child_item = child
                                child.setCheckState(0, Qt.Checked)
                                break
                    if not target_child_item and not clean_ver:
                        item.setCheckState(0, Qt.Checked)
                    break

            if target_child_item:
                self.version_tree.setCurrentItem(target_child_item)
                self.version_tree.scrollToItem(target_child_item)
            elif target_spec_item:
                self.version_tree.setCurrentItem(target_spec_item)
                self.version_tree.scrollToItem(target_spec_item)
                if clean_ver:
                    reply = QMessageBox.question(
                        self,
                        "Version Not Indexed",
                        f"TS {clean_spec} is indexed in Spec Search, but version v{clean_ver} is not yet added.\n\n"
                        f"Would you like to index TS {clean_spec} v{clean_ver} now?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes,
                    )
                    if reply == QMessageBox.Yes and self.specs_db:
                        files = self.specs_db.get_spec_files(clean_spec)
                        matching_file = next(
                            (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                            None
                        )
                        if matching_file:
                            task = {
                                "spec_number": clean_spec,
                                "version": clean_ver,
                                "filename": matching_file.get("filename", ""),
                                "file_url": matching_file.get("url", ""),
                                "release_date": matching_file.get("upload_date", ""),
                            }
                            self._start_batch_ingestion([task])
                            return

            self.selected_version_ids = self.version_tree.get_selected_version_ids()
        else:
            prompt_msg = (
                f"TS {clean_spec} v{clean_ver} is not yet indexed in the Spec Search database.\n\nWould you like to index it now?"
                if clean_ver
                else f"TS {clean_spec} is not yet indexed in the Spec Search database.\n\nWould you like to import it from the Specifications database now?"
            )
            reply = QMessageBox.question(
                self,
                "Spec Not Indexed",
                prompt_msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                if clean_ver and self.specs_db:
                    files = self.specs_db.get_spec_files(clean_spec)
                    matching_file = next(
                        (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                        None
                    )
                    if matching_file:
                        task = {
                            "spec_number": clean_spec,
                            "version": clean_ver,
                            "filename": matching_file.get("filename", ""),
                            "file_url": matching_file.get("url", ""),
                            "release_date": matching_file.get("upload_date", ""),
                        }
                        self._start_batch_ingestion([task])
                        return
                self._on_fetch_clicked()
                return

        self.search_input.setFocus()
        self.search_input.selectAll()
        if self.search_input.text().strip():
            self._execute_search()

    def shutdown(self, wait_ms: int = 1500):
        """Cooperatively stop timers and wait for active background workers during app shutdown."""
        self._search_timer.stop()
        self._save_config_timer.stop()

        workers = list(self._active_query_workers.values())
        if hasattr(self, "thread") and self.thread and self.thread.isRunning():
            workers.append(self.thread)
        if hasattr(self, "_wipe_worker") and self._wipe_worker and self._wipe_worker.isRunning():
            workers.append(self._wipe_worker)

        for worker in workers:
            if worker.isRunning():
                worker.requestInterruption()

        for worker in workers:
            if worker.isRunning():
                worker.wait(wait_ms)