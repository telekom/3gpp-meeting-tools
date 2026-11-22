"""UI construction mixin for SpecSearchTab."""

from __future__ import annotations

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QCheckBox,
    QDateEdit,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QProgressBar,
    QPushButton,
    QShortcut,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.spec_search.ui.spec_search_components import (
    ActiveSpecsScopeBar,
    SpecClauseInspector,
    SpecSearchVersionCardsWidget,
)


class SpecSearchTabUiSetupMixin:
    """Builds the main splitter, indexed specs card, search ribbon, matrix tabs, and inspector."""

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
        self.left_specs_panel = left_widget
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        left_card = QFrame()
        left_card.setObjectName("specTreeCard")
        left_card.setStyleSheet(f"""
            QFrame#specTreeCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)
        card_layout = QVBoxLayout(left_card)
        card_layout.setContentsMargins(6, 8, 6, 6)
        card_layout.setSpacing(6)

        # Card Header: Title + Count Badge + Compact Import & Maintenance Dropdowns
        tree_hdr_layout = QHBoxLayout()
        tree_hdr_layout.setContentsMargins(2, 0, 2, 0)
        tree_hdr_layout.setSpacing(4)
        tree_title = QLabel("📚 Indexed Specs")
        tree_title.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {C.TEXT_STRONG};")
        tree_hdr_layout.addWidget(tree_title)

        self.tree_stats_badge = QLabel("0")
        self.tree_stats_badge.setStyleSheet(f"""
            background-color: {C.BG_CANVAS};
            color: {C.TEXT_SECONDARY};
            font-size: 10px;
            font-weight: bold;
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: 4px;
            padding: 1px 5px;
        """)
        tree_hdr_layout.addWidget(self.tree_stats_badge)
        tree_hdr_layout.addStretch()

        # 1. Compact Import Menu Button
        self.import_menu_btn = QPushButton("📥 Import ▾")
        self.import_menu_btn.setCursor(Qt.PointingHandCursor)
        self.import_menu_btn.setToolTip("Import specifications from DB or local Word .docx")
        self.import_menu_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 7px;
                font-size: 11px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton::menu-indicator {{
                image: none;
                width: 0px;
            }}
        """)
        import_menu = QMenu(self)
        act_import_smart = import_menu.addAction("⭐ Smart Import Specification (Latest per Release)...")
        act_import_smart.triggered.connect(lambda: self._open_smart_import_dialog())
        import_menu.addSeparator()
        act_import_db = import_menu.addAction("📥 Browse All Specs && Releases...")
        act_import_db.triggered.connect(self._on_fetch_clicked)
        act_import_docx = import_menu.addAction("📁 Import Local Word Document (.docx)...")
        act_import_docx.triggered.connect(self._on_import_local_clicked)
        self.import_menu_btn.setMenu(import_menu)
        tree_hdr_layout.addWidget(self.import_menu_btn)

        # 2. Compact Database Management Menu Button
        self.db_menu_btn = QPushButton("⚙️ ▾")
        self.db_menu_btn.setCursor(Qt.PointingHandCursor)
        self.db_menu_btn.setToolTip("Search database maintenance (clear versions, wipe DB)")
        self.db_menu_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 11px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton::menu-indicator {{
                image: none;
                width: 0px;
            }}
        """)
        db_menu = QMenu(self)
        self.act_auto_replace = db_menu.addAction("Auto-supersede Older Versions on Import")
        self.act_auto_replace.setCheckable(True)
        self.act_auto_replace.setChecked(getattr(self, "auto_replace_older_versions", True))
        self.act_auto_replace.toggled.connect(self._on_toggle_auto_replace)
        db_menu.addSeparator()
        self.act_prune = db_menu.addAction("🧹 Prune Superseded Versions...")
        self.act_prune.triggered.connect(self._on_prune_superseded_clicked)
        self.act_clear_ver = db_menu.addAction("🗑️ Clear Specific Version...")
        self.act_clear_ver.triggered.connect(self._on_clear_version_clicked)
        db_menu.addSeparator()
        self.act_wipe_db = db_menu.addAction("⚠️ Wipe Entire Search Database...")
        self.act_wipe_db.triggered.connect(self._on_wipe_db_clicked)
        self.db_menu_btn.setMenu(db_menu)
        tree_hdr_layout.addWidget(self.db_menu_btn)

        # 3. Collapse Sidebar to Minimal Mode Button
        self.btn_collapse_specs = QPushButton("◀")
        self.btn_collapse_specs.setCursor(Qt.PointingHandCursor)
        self.btn_collapse_specs.setToolTip("Hide specs sidebar (Minimal Mode — Ctrl+B)")
        self.btn_collapse_specs.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                color: {C.ACCENT_SKY_DARK};
                border-color: {C.ACCENT_SKY};
            }}
        """)
        self.btn_collapse_specs.clicked.connect(lambda: self.set_specs_minimal_mode(True))
        tree_hdr_layout.addWidget(self.btn_collapse_specs)

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
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                background-color: {C.BORDER_SUBTLE};
                border-radius: 2px;
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_SKY};
                border-radius: 2px;
            }}
        """)
        self.progress_bar.setVisible(False)
        card_layout.addWidget(self.progress_bar)

        # Quick Filter Input
        self.tree_filter_input = QLineEdit()
        self.tree_filter_input.setPlaceholderText("Filter specs (e.g. 23.501, slicing)...")
        self.tree_filter_input.setClearButtonEnabled(True)
        self.tree_filter_input.setStyleSheet(f"""
            QLineEdit {{
                padding: 3px 6px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                font-size: 11px;
                background-color: {C.BG_SURFACE};
            }}
            QLineEdit:focus {{
                border: 1px solid {C.ACCENT_SKY};
                background-color: {C.BG_CARD};
            }}
        """)
        card_layout.addWidget(self.tree_filter_input)

        # Quick Filter Action Chips
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(4)

        btn_all = QPushButton("All")
        btn_all.setToolTip("Select all indexed releases across all specs")
        btn_all.setCursor(Qt.PointingHandCursor)
        btn_all.setStyleSheet(f"""
            QPushButton {{
                font-size: 10px;
                font-weight: 600;
                color: {C.TEXT_BODY};
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 3px;
                padding: 2px 6px;
            }}
            QPushButton:hover {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_PRIMARY};
            }}
        """)
        chips_layout.addWidget(btn_all)

        btn_none = QPushButton("None")
        btn_none.setToolTip("Deselect all indexed releases")
        btn_none.setCursor(Qt.PointingHandCursor)
        btn_none.setStyleSheet(f"""
            QPushButton {{
                font-size: 10px;
                font-weight: 600;
                color: {C.TEXT_BODY};
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 3px;
                padding: 2px 6px;
            }}
            QPushButton:hover {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_PRIMARY};
            }}
        """)
        chips_layout.addWidget(btn_none)

        btn_latest = QPushButton("⭐ Latest/Rel")
        btn_latest.setToolTip("Check only the newest version of each Release per specification")
        btn_latest.setCursor(Qt.PointingHandCursor)
        btn_latest.setStyleSheet(f"""
            QPushButton {{
                font-size: 10px;
                font-weight: 600;
                color: {C.ACCENT_SKY_DARK};
                background-color: {C.SKY_SURFACE};
                border: 1px solid {C.SKY_BORDER};
                border-radius: 3px;
                padding: 2px 6px;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_SOFT};
                color: {C.ACCENT_SKY};
            }}
        """)
        chips_layout.addWidget(btn_latest)

        btn_r18 = QPushButton("Rel-18+")
        btn_r18.setToolTip("Select only Release 18, 19, and 20 (5G-Advanced & 6G)")
        btn_r18.setCursor(Qt.PointingHandCursor)
        btn_r18.setStyleSheet(f"""
            QPushButton {{
                font-size: 10px;
                font-weight: 600;
                color: {C.INDIGO_DARK};
                background-color: {C.WG_SA_BG};
                border: 1px solid {C.WG_SA_BORDER};
                border-radius: 3px;
                padding: 2px 6px;
            }}
            QPushButton:hover {{
                background-color: {C.INDIGO_SOFT};
                color: {C.WG_SA_TEXT};
            }}
        """)
        chips_layout.addWidget(btn_r18)
        chips_layout.addStretch()
        card_layout.addLayout(chips_layout)

        self.version_tree = SpecSearchVersionCardsWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        self.version_tree.prune_spec_requested.connect(self._prune_single_spec)
        self.version_tree.smart_import_requested.connect(self._open_smart_import_dialog)
        self.tree_filter_input.textChanged.connect(self.version_tree.filter_tree)
        btn_all.clicked.connect(self.version_tree.select_all)
        btn_none.clicked.connect(self.version_tree.deselect_all)
        btn_latest.clicked.connect(self.version_tree.select_latest_per_release)
        btn_r18.clicked.connect(lambda: self.version_tree.select_releases_min(18))
        card_layout.addWidget(self.version_tree)

        left_layout.addWidget(left_card)
        self.main_splitter.addWidget(left_widget)

        # Right Panel: Search Controls, Tabbed Evolution Matrix, Inspector
        self.right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)
        matrix_layout.setSpacing(6)

        # Minimal-Mode Top Scope Badge Bar (visible when left sidebar is collapsed)
        self.scope_bar = ActiveSpecsScopeBar(self)
        self.scope_bar.setVisible(False)
        self.scope_bar.expand_full_mode_requested.connect(lambda: self.set_specs_minimal_mode(False))
        self.scope_bar.solo_spec_toggled.connect(self._on_solo_spec_toggled)
        self.scope_bar.preset_requested.connect(self._on_scope_preset_requested)
        self.scope_bar.import_smart_requested.connect(lambda: self._open_smart_import_dialog())
        self.scope_bar.import_browse_requested.connect(self._on_fetch_clicked)
        self.scope_bar.import_local_requested.connect(self._on_import_local_clicked)
        matrix_layout.addWidget(self.scope_bar)

        # ---------------------------------------------------------------------
        # 2-Tier Carded Search & Filter Ribbon
        # ---------------------------------------------------------------------
        self.search_ribbon = PageRibbonCard("Spec Search", parent=self)

        search_bar_widget = QWidget(self.search_ribbon)
        search_bar_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_bar_layout = QHBoxLayout(search_bar_widget)
        search_bar_layout.setContentsMargins(0, 0, 0, 0)
        search_bar_layout.setSpacing(6)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search exact substring or phrase (e.g. 'No-Transmit Zone', 'emergency', 'PDU session')...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(scale_px(200))
        self.search_input.setStyleSheet(f"""
            QLineEdit {{
                padding: 5px 10px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                font-size: {T.SIZE_BODY}px;
                background-color: {C.BG_CARD};
                color: {C.TEXT_PRIMARY};
            }}
            QLineEdit:focus {{
                border: 1.5px solid {C.ACCENT_SKY};
            }}
        """)
        self.search_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.search_input, stretch=1)

        self.clause_filter_input = QLineEdit()
        self.clause_filter_input.setPlaceholderText("Clause (e.g. 5.2)...")
        self.clause_filter_input.setToolTip(
            "Filter by clause number or title (also works with no keyword to browse a clause hierarchy, e.g. 5.15)"
        )
        self.clause_filter_input.setMinimumWidth(scale_px(150))
        self.clause_filter_input.setMaximumWidth(scale_px(240))
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.setStyleSheet(f"""
            QLineEdit {{
                padding: 5px 8px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                font-size: {T.SIZE_TABLE}px;
                background-color: {C.BG_CARD};
            }}
            QLineEdit:focus {{
                border: 1.5px solid {C.ACCENT_SKY};
            }}
        """)
        self.clause_filter_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.clause_filter_input)
        self.search_ribbon.add_expanding_primary_widget(search_bar_widget, stretch=1)

        self.btn_find_origin_cr = QPushButton("Find Origin CR")
        self.btn_find_origin_cr.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_find_origin_cr.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.btn_find_origin_cr, "search", size=14)
        self.btn_find_origin_cr.setToolTip("Ranks candidate Change Requests (CRs) across all releases to identify the origin CR that introduced the searched feature.")
        self.btn_find_origin_cr.clicked.connect(self._on_find_origin_cr_clicked)
        self.search_ribbon.add_utility_widget(self.btn_find_origin_cr)

        # Tier 2: Cutoff Date Filter + Post-Cutoff Checkbox (auto-wraps cleanly)
        cutoff_group = QWidget(self.search_ribbon)
        cutoff_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        cutoff_layout = QHBoxLayout(cutoff_group)
        cutoff_layout.setContentsMargins(0, 0, 0, 0)
        cutoff_layout.setSpacing(S.SM)

        self.chk_filing_date = QCheckBox("Date Cutoff:")
        self.chk_filing_date.setStyleSheet(f"font-weight: 600; font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_BODY};")
        self.chk_filing_date.toggled.connect(self._on_cutoff_date_filter_toggled)
        cutoff_layout.addWidget(self.chk_filing_date)

        self.cutoff_date_edit = QDateEdit()
        self.cutoff_date_edit.setCalendarPopup(True)
        self.cutoff_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.cutoff_date_edit.setDate(QDate.currentDate().addYears(-3))
        self.cutoff_date_edit.setMinimumWidth(scale_px(125))
        self.cutoff_date_edit.setEnabled(False)
        self.cutoff_date_edit.setStyleSheet(f"""
            QDateEdit {{
                padding: 3px 6px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                font-size: {T.SIZE_TABLE}px;
                background-color: {C.BG_CARD};
                color: {C.TEXT_STRONG};
            }}
            QDateEdit:disabled {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_FAINT};
            }}
        """)
        self.cutoff_date_edit.dateChanged.connect(self._on_filter_changed)
        cutoff_layout.addWidget(self.cutoff_date_edit)

        self.chk_post_date_only = QCheckBox("Show Only Post-Cutoff Additions")
        self.chk_post_date_only.setToolTip("Hides clauses where matching text was already present prior to the selected date.")
        self.chk_post_date_only.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        self.chk_post_date_only.setEnabled(False)
        self.chk_post_date_only.toggled.connect(self._on_filter_changed)
        cutoff_layout.addWidget(self.chk_post_date_only)

        self.search_ribbon.add_filter_widget(cutoff_group)

        matrix_layout.addWidget(self.search_ribbon)

        # Executive Results Status Pill Bar
        self.matrix_title = QLabel("💡 Type a query above to trace clause evolution across releases.")
        self.matrix_title.setStyleSheet(f"font-size: 12px; color: {C.TEXT_MUTED}; margin-top: 2px; margin-bottom: 2px;")
        matrix_layout.addWidget(self.matrix_title)

        # Tabbed Container for Per-Specification Matrices
        self.spec_results_tabs = QTabWidget()
        self.spec_results_tabs.setDocumentMode(True)
        self.spec_results_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                top: -1px;
            }}
            QTabBar::tab {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                font-weight: 600;
                font-size: 11px;
                border: 1px solid {C.BORDER_STRONG};
                border-bottom: none;
                padding: 5px 14px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
            QTabBar::tab:selected {{
                background-color: {C.BG_CARD};
                color: {C.ACCENT_SKY};
                font-weight: bold;
                border-bottom: 2px solid {C.ACCENT_SKY};
            }}
            QTabBar::tab:hover:!selected {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_PRIMARY};
            }}
        """)
        matrix_layout.addWidget(self.spec_results_tabs, 1)
        self.right_splitter.addWidget(matrix_widget)

        self.inspector = SpecClauseInspector()
        self.inspector.compare_clause_requested.connect(self._open_clause_diff_dialog)
        self.inspector.open_document_requested.connect(self._handle_open_spec_document)
        self.inspector.show_in_specs_tab_requested.connect(self._handle_show_in_specs_tab)
        self.inspector.quick_diff_requested.connect(self._handle_quick_diff)
        self.inspector.version_changes_requested.connect(self._open_version_history_dialog)
        self.inspector.identify_origin_tdoc_requested.connect(self._handle_identify_origin_tdoc)
        self.inspector.jump_to_version_requested.connect(self._handle_jump_to_version)
        self.inspector.extract_clause_requested.connect(self._handle_extract_clause_word)
        self.inspector.export_markdown_requested.connect(self._handle_export_clause_markdown)
        self.inspector.layout_orientation_toggle_requested.connect(self._toggle_inspector_layout)
        self.inspector.maximize_toggle_requested.connect(self._toggle_inspector_maximize)
        self.inspector.clause_reference_clicked.connect(self._handle_clause_reference_clicked)
        if hasattr(self.inspector, "doc_menu"):
            self.inspector.doc_menu.aboutToShow.connect(self._refresh_inspector_doc_menu)
        self.right_splitter.addWidget(self.inspector)

        self.right_splitter.setSizes([380, 260])
        self.main_splitter.addWidget(self.right_splitter)
        self.main_splitter.setSizes([300, 750])
        layout.addWidget(self.main_splitter, 1)

        self._toggle_specs_shortcut = QShortcut(QKeySequence("Ctrl+B"), self)
        self._toggle_specs_shortcut.activated.connect(self.toggle_specs_panel_mode)
