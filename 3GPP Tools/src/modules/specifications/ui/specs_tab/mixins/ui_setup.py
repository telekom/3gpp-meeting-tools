"""UI construction and responsive font-scaling mixin for SpecificationsTab."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QKeySequence, QPalette
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QShortcut,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.icons import lucide_svg_path
from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon, make_filter_label
from core.ui.theme import C, R, S, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    CLEAR_ALL_BTN_STYLE,
    FAV_PILL_STYLE,
    SEGMENT_BTN_STYLE,
    SPEC_FILTER_COMBO_STYLE,
    SPEC_FILTER_INPUT_STYLE,
    TOGGLE_PILL_STYLE,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox


class SpecsTabUiSetupMixin:
    """Builds the PageRibbonCard header, filter controls, table, and shortcuts."""

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(S.MD, S.SM + 2, S.MD, S.SM + 2)
        main_layout.setSpacing(S.SM + 2)

        # ==========================================
        # --- 2-TIER CARDED PAGE RIBBON ---
        # ==========================================
        self.ribbon_card = PageRibbonCard("Specifications", parent=self)

        # Tier 1 Primary Actions
        self.sync_menu_btn = QPushButton("Sync && Fetch ▾")
        self.sync_menu_btn.setObjectName("primaryBtn")
        self.sync_menu_btn.setMinimumHeight(scale_px(30))
        self.sync_menu_btn.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.sync_menu_btn, "globe", color=C.TEXT_ON_ACCENT, active_color=C.TEXT_ON_ACCENT, size=14)

        sync_menu = QMenu(self)

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
        self.ribbon_card.add_primary_widget(self.sync_menu_btn)

        # Add Individual Specification Button
        self.add_spec_btn = QPushButton("Add Spec")
        self.add_spec_btn.setMinimumHeight(scale_px(30))
        self.add_spec_btn.setToolTip("Add and inspect an individual specification without running a full sync.")
        self.add_spec_btn.setCursor(Qt.PointingHandCursor)
        self.add_spec_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.add_spec_btn, "plus", size=14)
        self.add_spec_btn.clicked.connect(self._open_add_spec_dialog)
        self.ribbon_card.add_primary_widget(self.add_spec_btn)

        self.bg_sync_label = QLabel("⏳ Fetching deep metadata in background...")
        self.bg_sync_label.setStyleSheet(f"color: {C.WARNING}; font-weight: bold; font-style: italic;")
        self.bg_sync_label.setVisible(False)
        self.ribbon_card.add_primary_widget(self.bg_sync_label)

        # Tier 1 Utility Actions (Right-Aligned)
        self.open_dir_btn = QPushButton("Open Folder")
        self.open_dir_btn.setMinimumHeight(scale_px(30))
        self.open_dir_btn.setToolTip("Open the local specifications storage directory in Explorer.")
        self.open_dir_btn.setCursor(Qt.PointingHandCursor)
        self.open_dir_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.open_dir_btn, "folder-open", size=14)
        self.open_dir_btn.clicked.connect(self._open_download_dir)
        self.ribbon_card.add_utility_widget(self.open_dir_btn)

        self.settings_btn = QPushButton("Settings")
        self.settings_btn.setMinimumHeight(scale_px(30))
        self.settings_btn.setToolTip("Configure the local specifications download folder and paths.")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.settings_btn, "settings", size=14)
        self.settings_btn.clicked.connect(self._open_settings_dialog)
        self.ribbon_card.add_utility_widget(self.settings_btn)

        self.wipe_db_btn = QPushButton("Wipe DB")
        self.wipe_db_btn.setMinimumHeight(scale_px(30))
        self.wipe_db_btn.setToolTip("Clear all specifications, file mappings, and reset the local database.")
        self.wipe_db_btn.setCursor(Qt.PointingHandCursor)
        self.wipe_db_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        apply_lucide_icon(self.wipe_db_btn, "trash-2", color=C.DANGER_ALT, active_color=C.DANGER_DARK, size=14)
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        self.ribbon_card.add_utility_widget(self.wipe_db_btn)

        # --- TIER 2: Search, Quick Scope & Faceted Filters ---
        search_group = QWidget(self.ribbon_card)
        search_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)

        search_layout.addWidget(make_filter_label("Search:", emphasized=True))

        self.spec_search_input = QLineEdit()
        self.spec_search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.spec_search_input.setPlaceholderText("Spec # (e.g. 23.501) or title...")
        self.spec_search_input.setText(self.saved_search)
        self.spec_search_input.setClearButtonEnabled(True)
        self.spec_search_input.setMinimumWidth(scale_px(210))
        self.spec_search_input.setToolTip("Filter instantly by specification number or title keywords.")
        self.spec_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.spec_search_input)

        search_layout.addWidget(make_filter_label("Ver:", emphasized=True))

        self.version_search_input = QLineEdit()
        self.version_search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.version_search_input.setPlaceholderText("17 or 16.2")
        self.version_search_input.setText(self.saved_version)
        self.version_search_input.setClearButtonEnabled(True)
        self.version_search_input.setToolTip("Filter table by 3GPP release version (e.g. '15', '16.2', '17').")
        self.version_search_input.setMinimumWidth(scale_px(105))
        self.version_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.version_search_input)
        self.ribbon_card.add_filter_widget(search_group)

        # Quick Toggle Pills
        scope_group = QWidget(self.ribbon_card)
        scope_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        scope_layout = QHBoxLayout(scope_group)
        scope_layout.setContentsMargins(0, 0, 0, 0)
        scope_layout.setSpacing(6)

        self.downloaded_only_btn = QPushButton("💾 Downloaded Only")
        self.downloaded_only_btn.setCheckable(True)
        self.downloaded_only_btn.setChecked(self.saved_downloaded_only)
        self.downloaded_only_btn.setCursor(Qt.PointingHandCursor)
        self.downloaded_only_btn.setStyleSheet(TOGGLE_PILL_STYLE)
        self.downloaded_only_btn.setToolTip("Show only specifications that currently have local cached files.")
        self.downloaded_only_btn.toggled.connect(self._toggle_downloaded_filter)
        self.downloaded_only_cb = self.downloaded_only_btn  # Backward compatibility alias
        scope_layout.addWidget(self.downloaded_only_btn)

        self.fav_chip = QPushButton(f"⭐ Favorites ({len(self.favorites)})")
        self.fav_chip.setCheckable(True)
        self.fav_chip.setChecked(self.favorites_only)
        self.fav_chip.setCursor(Qt.PointingHandCursor)
        self.fav_chip.setStyleSheet(FAV_PILL_STYLE)
        self.fav_chip.setToolTip("Toggle view to show only your starred/favorite specifications.")
        self.fav_chip.toggled.connect(self._toggle_favorites_filter)
        scope_layout.addWidget(self.fav_chip)
        self.ribbon_card.add_filter_widget(scope_group)

        # TSG Segmented Chips (All, RAN, SA, CT)
        tsg_group = QWidget(self.ribbon_card)
        tsg_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tsg_layout = QHBoxLayout(tsg_group)
        tsg_layout.setContentsMargins(0, 0, 0, 0)
        tsg_layout.setSpacing(4)
        tsg_layout.addWidget(make_filter_label("TSG:"))

        self.tsg_buttons = {}
        preset_tsgs = [
            ("All", ""),
            ("RAN", "RAN"),
            ("SA", "SA"),
            ("CT", "CT"),
        ]
        for label, tsg_val in preset_tsgs:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(
                f"Filter specifications under TSG {label} (Plenary and Working Groups)"
                if tsg_val
                else "Show specifications across all Technical Specification Groups"
            )
            btn.clicked.connect(lambda _, t=tsg_val: self._on_tsg_button_clicked(t))
            self.tsg_buttons[tsg_val] = btn
            tsg_layout.addWidget(btn)
        self.chip_buttons = self.tsg_buttons  # Backward compatibility alias
        self.ribbon_card.add_filter_widget(tsg_group)

        # Working Group + Series + Tech group
        wg_group = QWidget(self.ribbon_card)
        wg_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        wg_layout = QHBoxLayout(wg_group)
        wg_layout.setContentsMargins(0, 0, 0, 0)
        wg_layout.setSpacing(6)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_combo.setMinimumWidth(scale_px(125))
        self.wg_combo.selectionChanged.connect(self._on_wg_filter_changed)
        self.wg_combo.setContextMenuPolicy(Qt.CustomContextMenu)
        self.wg_combo.customContextMenuRequested.connect(self._show_wg_context_menu)
        wg_layout.addWidget(self.wg_combo)

        self.wg_scope_btn = QPushButton("Active WGs")
        self.wg_scope_btn.setCheckable(True)
        self.wg_scope_btn.setMinimumHeight(scale_px(26))
        self.wg_scope_btn.setCursor(Qt.PointingHandCursor)
        self.wg_scope_btn.setStyleSheet(SEGMENT_BTN_STYLE)
        self.wg_scope_btn.clicked.connect(self._toggle_wg_scope)
        self._set_wg_scope(getattr(self, "_show_all_wgs", False))
        wg_layout.addWidget(self.wg_scope_btn)

        self.series_combo = CheckableComboBox("Series")
        self.series_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.series_combo.setMinimumWidth(scale_px(125))
        self.series_combo.addItems(self.filter_options.get("series", []))
        self.series_combo.selectionChanged.connect(self._on_series_filter_changed)
        wg_layout.addWidget(self.series_combo)

        self.tech_combo = CheckableComboBox("Tech")
        self.tech_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.tech_combo.setMinimumWidth(scale_px(105))
        self.tech_combo.addItems(self.filter_options.get("techs", []))
        self.tech_combo.selectionChanged.connect(self._on_tech_filter_changed)
        wg_layout.addWidget(self.tech_combo)
        self.ribbon_card.add_filter_widget(wg_group)

        # Spec Type + Tag Filter + Reset All
        type_tag_group = QWidget(self.ribbon_card)
        type_tag_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tt_layout = QHBoxLayout(type_tag_group)
        tt_layout.setContentsMargins(0, 0, 0, 0)
        tt_layout.setSpacing(4)
        tt_layout.addWidget(make_filter_label("Type:"))

        self.type_buttons = {}
        for t_label, t_val in [("All", "Any"), ("TS", "TS"), ("TR", "TR")]:
            btn = QPushButton(t_label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(f"Filter by specification document type: {t_label}")
            btn.clicked.connect(lambda _, tv=t_val: self._on_type_button_clicked(tv))
            self.type_buttons[t_val] = btn
            tt_layout.addWidget(btn)

        tt_layout.addWidget(make_filter_label("Tag:"))
        self.tag_filter_combo = QComboBox()
        self.tag_filter_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.tag_filter_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.tag_filter_combo.setMinimumWidth(scale_px(115))
        self.tag_filter_combo.currentIndexChanged.connect(self._on_tag_filter_changed)
        tt_layout.addWidget(self.tag_filter_combo)

        self.clear_all_btn = QPushButton("Reset All")
        self.clear_all_btn.setCursor(Qt.PointingHandCursor)
        self.clear_all_btn.setStyleSheet(CLEAR_ALL_BTN_STYLE)
        apply_lucide_icon(self.clear_all_btn, "rotate-ccw", color=C.DANGER_ALT, size=13)
        self.clear_all_btn.setToolTip("Reset all search fields, active filters, and scopes.")
        self.clear_all_btn.clicked.connect(self._clear_all_filters)
        self.clear_all_btn.setVisible(False)
        tt_layout.addWidget(self.clear_all_btn)
        self.ribbon_card.add_filter_widget(type_tag_group)

        main_layout.addWidget(self.ribbon_card)

        # ==========================================
        # --- CARD: Active Filter Strip & Data Table ---
        # ==========================================
        self.table_card = QFrame(self)
        self.table_card.setObjectName("cardFrame")
        self.table_card.setStyleSheet(CARD_FRAME_STYLE)
        table_card_layout = QVBoxLayout(self.table_card)
        table_card_layout.setContentsMargins(0, 0, 0, 0)
        table_card_layout.setSpacing(0)

        self.active_tags_widget = QFrame(self.table_card)
        self.active_tags_widget.setObjectName("specsHeaderStrip")
        self.active_tags_widget.setStyleSheet(f"""
            QFrame#specsHeaderStrip {{
                background-color: {C.BG_SURFACE};
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: {R.LG}px;
                border-top-right-radius: {R.LG}px;
            }}
        """)
        tags_bar_layout = QHBoxLayout(self.active_tags_widget)
        tags_bar_layout.setContentsMargins(S.MD, S.XS + 2, S.MD, S.XS + 2)
        tags_bar_layout.setSpacing(6)

        self.active_tags_label = make_filter_label("Active:", emphasized=True)
        self.active_tags_label.setVisible(False)
        tags_bar_layout.addWidget(self.active_tags_label)

        self.active_tags_container = QWidget(self.active_tags_widget)
        self.active_tags_container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        self.active_tags_container_layout = QHBoxLayout(self.active_tags_container)
        self.active_tags_container_layout.setContentsMargins(0, 0, 0, 0)
        self.active_tags_container_layout.setSpacing(4)
        tags_bar_layout.addWidget(self.active_tags_container)

        tags_bar_layout.addStretch()

        self.count_badge = QLabel("0 specifications")
        self.count_badge.setStyleSheet(f"""
            QLabel {{
                padding: 2px 10px;
                font-size: {T.SIZE_CAPTION}px;
                font-weight: bold;
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 10px;
            }}
        """)
        tags_bar_layout.addWidget(self.count_badge)

        table_card_layout.addWidget(self.active_tags_widget)

        self.table = QTableWidget(self.table_card)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Specification", "Title", "Version / Documents"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        self.table.setColumnWidth(0, max(210, scale_px(210)))
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(2, max(445, scale_px(445)))

        self.table.setShowGrid(False)
        register_scaled_row_height(self.table, 44)
        self.table.verticalHeader().setStyleSheet(f"""
            QHeaderView {{
                background-color: {C.BG_CARD};
                border: none;
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_FAINT};
                font-size: {T.SIZE_CAPTION}px;
                border: none;
                border-bottom: 1px solid {C.BORDER_ROW};
                padding: 0px 6px;
            }}
        """)
        table_pal = self.table.viewport().palette()
        table_pal.setColor(QPalette.Base, QColor(C.BG_CARD))
        table_pal.setColor(QPalette.AlternateBase, QColor(C.BG_SURFACE))
        self.table.viewport().setPalette(table_pal)
        self.table.viewport().setAutoFillBackground(True)

        chevron_url = lucide_svg_path("chevron-down", C.TEXT_MUTED)
        self._version_combo_style = f"""
            QComboBox {{
                background-color: {C.BG_CARD};
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                padding: 2px 24px 2px 8px;
                font-size: {T.SIZE_TABLE}px;
                font-weight: 600;
            }}
            QComboBox:hover {{
                border-color: {C.BORDER_HOVER};
                background-color: {C.BG_SURFACE};
            }}
            QComboBox:focus, QComboBox:on {{
                border-color: {C.ACCENT_PRIMARY};
                background-color: {C.BG_CARD};
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 22px;
                border: none;
            }}
            QComboBox::down-arrow {{
                image: url("{chevron_url}");
                width: 12px;
                height: 12px;
            }}
        """

        self.table.setStyleSheet(f"""
            QTableWidget {{
                border: none;
                border-bottom-left-radius: {R.LG}px;
                border-bottom-right-radius: {R.LG}px;
                background-color: {C.BG_CARD};
                alternate-background-color: {C.BG_SURFACE};
                selection-background-color: {C.ACCENT_SOFT};
                selection-color: {C.TEXT_STRONG};
                font-size: {T.SIZE_TABLE}px;
            }}
            QTableWidget > QWidget {{
                background-color: {C.BG_CARD};
            }}
            QTableWidget::item {{
                padding: 4px 8px;
                border-bottom: 1px solid {C.BORDER_ROW};
            }}
            QTableWidget::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.TEXT_STRONG};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                font-weight: bold;
                font-size: {T.SIZE_TABLE}px;
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                padding: 6px 8px;
            }}
            QTableCornerButton::section {{
                background-color: {C.BG_SURFACE};
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
            }}
            QWidget#specRowCell {{
                background-color: transparent;
                border: none;
            }}
        """)

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.cellDoubleClicked.connect(self._on_table_cell_double_clicked)

        # Quick Copy keyboard shortcut (Ctrl+C)
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

        table_card_layout.addWidget(self.table, 1)
        main_layout.addWidget(self.table_card, 1)
        self.setLayout(main_layout)

        self._sync_ui_to_filters()

    def on_font_scale_changed(self, _scale: float = 0.0):
        """Adapt filter controls, column widths, and row widgets for the updated font scale."""
        self.spec_search_input.setMinimumWidth(scale_px(210))
        self.version_search_input.setMinimumWidth(scale_px(105))
        self.wg_combo.setMinimumWidth(scale_px(125))
        self.series_combo.setMinimumWidth(scale_px(125))
        self.tech_combo.setMinimumWidth(scale_px(105))
        self.tag_filter_combo.setMinimumWidth(scale_px(115))
        self.wg_scope_btn.setMinimumHeight(scale_px(26))
        for btn in (
            self.sync_menu_btn,
            self.add_spec_btn,
            self.open_dir_btn,
            self.settings_btn,
            self.wipe_db_btn,
        ):
            btn.setMinimumHeight(scale_px(30))
        self.table.setColumnWidth(0, max(210, scale_px(210)))
        self.table.setColumnWidth(2, max(445, scale_px(445)))
        if getattr(self, "_is_loaded", False) or self.table.rowCount() > 0:
            self.refresh_table()

    def set_bg_sync_active(self, is_active: bool):
        self.bg_sync_label.setVisible(is_active)
        if not is_active and self._is_loaded:
            self.refresh_table()
