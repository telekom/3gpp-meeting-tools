"""2-tier PageRibbonCard, Meetings table view, footer cart, and DPI font-scale mixin."""

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QSizePolicy,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon, make_filter_label
from core.ui.theme import C, R, S, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_WARNING,
    CARD_FRAME_STYLE,
    SEGMENT_BTN_STYLE,
    SPEC_FILTER_COMBO_STYLE,
    SPEC_FILTER_INPUT_STYLE,
)
from modules.meetings.ui.meeting_details_dialogs import MeetingsConfigDialog
from modules.meetings.ui.meetings_delegates import (
    DynamicHoverMenuDelegate,
    TDocsButtonDelegate,
    WGTagDelegate,
)
from modules.meetings.ui.models import MeetingsTableModel
from modules.meetings.ui.tdocs_components import CheckableComboBox, ComparisonCartWidget


class MeetingsTabUiSetupMixin:
    """Builds the MeetingsTab layout, ribbons, filters bar, table, and cart footer."""

    def _apply_scaled_table_columns(self):
        """Scale fixed table section widths so pills, dates, and TDoc ranges fit cleanly."""
        header = self.table.horizontalHeader()
        header.resizeSection(0, max(38, scale_px(38)))
        header.resizeSection(1, max(96, scale_px(96)))
        header.resizeSection(2, max(78, scale_px(78)))
        header.resizeSection(3, max(100, scale_px(100)))
        # col 4 Stretch (Location)
        header.resizeSection(5, max(155, scale_px(155)))
        header.resizeSection(6, max(215, scale_px(215)))

    def on_font_scale_changed(self, _scale: float = 0.0):
        """Adapt filter inputs, dropdowns, and fixed table columns for the new font scale."""
        self.global_tdoc_input.setMinimumHeight(scale_px(30))
        self.global_tdoc_input.setMinimumWidth(scale_px(250))
        self.search_input.setMinimumHeight(scale_px(28))
        self.search_input.setMinimumWidth(scale_px(210))

        for btn in (self.btn_open_last, self.sync_menu_btn, self.reports_menu_btn, self.btn_config):
            btn.setMinimumHeight(scale_px(30))
        for btn in (
            self.btn_open_tdoc,
            self.btn_open_meeting,
            self.btn_compare_active_word,
            self.btn_copy_tdoc,
            self.btn_reset_filters,
        ):
            btn.setMinimumHeight(scale_px(28))
        for btn in self.tsg_buttons.values():
            btn.setMinimumHeight(scale_px(26))
        self.btn_update_tdoc_metadata.setMinimumHeight(scale_px(26))
        self.tdoc_metadata_warning.setStyleSheet(
            f"color: {C.WARNING_DARK}; font-size: {scale_px(T.SIZE_CAPTION)}px; "
            f"font-weight: 600; background: transparent; border: none;"
        )
        if hasattr(self, "sync_status_label"):
            self.sync_status_label.setStyleSheet(
                f"color: {C.INFO_DARK}; font-size: {scale_px(T.SIZE_CAPTION)}px; "
                f"font-weight: 600; background: transparent; border: none;"
            )
        if hasattr(self, "btn_stop_sync"):
            self.btn_stop_sync.setMinimumHeight(scale_px(26))

        for widget in (
            self.wg_filter,
            self.adhoc_filter,
            self.type_filter,
            self.date_preset_combo,
            self.date_from,
            self.date_to,
        ):
            widget.setMinimumHeight(scale_px(28))
            widget.updateGeometry()

        self._apply_scaled_table_columns()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(S.MD, S.SM + 2, S.MD, S.SM + 2)
        main_layout.setSpacing(S.SM + 2)

        # --- Settings Dialog & Persistent Reference Bridge ---
        self.config_dialog = MeetingsConfigDialog(self.settings, self)
        self.dl_dir_input = self.config_dialog.dl_dir_input
        self.chk_wg = self.config_dialog.chk_wg
        self.chk_dyna = self.config_dialog.chk_dyna
        self.chk_docs = self.config_dialog.chk_docs
        self.delete_all_btn = getattr(self.config_dialog, "btn_clear", None)

        # ==========================================
        # --- 2-TIER CARDED PAGE RIBBON ---
        # ==========================================
        self.ribbon_card = PageRibbonCard("Meetings", parent=self)

        # 1. Quick Launch Last Meeting
        self.btn_open_last = QPushButton("Open Last Meeting")
        self.btn_open_last.setObjectName("primaryBtn")
        self.btn_open_last.setMinimumHeight(scale_px(30))
        self.btn_open_last.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.btn_open_last, "play", color=C.TEXT_ON_ACCENT, active_color=C.TEXT_ON_ACCENT, size=13)
        self.btn_open_last.clicked.connect(self._open_last_meeting)
        self.ribbon_card.add_primary_widget(self.btn_open_last)

        # 2. Quick TDoc Jump group (travels cleanly as one wrapping unit)
        tdoc_group = QWidget(self.ribbon_card)
        tdoc_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tdoc_layout = QHBoxLayout(tdoc_group)
        tdoc_layout.setContentsMargins(0, 0, 0, 0)
        tdoc_layout.setSpacing(6)

        sep_v1 = QFrame(tdoc_group)
        sep_v1.setFixedSize(1, 20)
        sep_v1.setStyleSheet(f"background-color: {C.BORDER_SUBTLE}; border: none;")
        tdoc_layout.addWidget(sep_v1)

        tdoc_lbl = make_filter_label("TDoc:", emphasized=True)
        tdoc_layout.addWidget(tdoc_lbl)

        self.global_tdoc_input = QLineEdit()
        self.global_tdoc_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.global_tdoc_input.setPlaceholderText("Jump to TDoc (e.g., S2-2605740)...")
        self.global_tdoc_input.setMinimumHeight(scale_px(30))
        self.global_tdoc_input.setMinimumWidth(scale_px(250))
        self.global_tdoc_input.setClearButtonEnabled(True)
        self.global_tdoc_input.setContextMenuPolicy(Qt.CustomContextMenu)
        self.global_tdoc_input.customContextMenuRequested.connect(self._on_tdoc_input_context_menu)
        self.global_tdoc_input.setToolTip("Type a TDoc number. Press Enter to open or use the quick buttons.")
        tdoc_layout.addWidget(self.global_tdoc_input)

        self.ribbon_card.add_primary_widget(tdoc_group)

        # Quick TDoc Jump action buttons wrap independently in primary_flow so they never clip off-screen
        self.btn_open_tdoc = QPushButton("Doc")
        self.btn_open_tdoc.setObjectName("primaryBtn")
        self.btn_open_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc.setMinimumHeight(scale_px(28))
        self.btn_open_tdoc.setVisible(False)
        apply_lucide_icon(self.btn_open_tdoc, "file-text", color=C.TEXT_ON_ACCENT, active_color=C.TEXT_ON_ACCENT, size=13)
        self.btn_open_tdoc.setToolTip("Download and open the specified TDoc directly in Microsoft Word.")
        self.ribbon_card.add_primary_widget(self.btn_open_tdoc)

        self.btn_open_meeting = QPushButton("Mtg")
        self.btn_open_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_meeting.setCursor(Qt.PointingHandCursor)
        self.btn_open_meeting.setMinimumHeight(scale_px(28))
        self.btn_open_meeting.setVisible(False)
        apply_lucide_icon(self.btn_open_meeting, "calendar", size=13)
        self.btn_open_meeting.setToolTip("Open the full meeting table containing this TDoc.")
        self.ribbon_card.add_primary_widget(self.btn_open_meeting)

        self.btn_compare_active_word = QPushButton("vs Word")
        self.btn_compare_active_word.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_compare_active_word.setCursor(Qt.PointingHandCursor)
        self.btn_compare_active_word.setMinimumHeight(scale_px(28))
        self.btn_compare_active_word.setVisible(False)
        apply_lucide_icon(self.btn_compare_active_word, "git-compare", size=13)
        self.btn_compare_active_word.setToolTip(
            "Download this TDoc and immediately compare it with your currently active Microsoft Word document."
        )
        self.ribbon_card.add_primary_widget(self.btn_compare_active_word)

        self.btn_copy_tdoc = QPushButton("Copy ▾")
        self.btn_copy_tdoc.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_copy_tdoc.setMinimumHeight(scale_px(28))
        self.btn_copy_tdoc.setVisible(False)
        apply_lucide_icon(self.btn_copy_tdoc, "copy", size=13)
        self.btn_copy_tdoc.setToolTip("Copy TDoc number, title, citation, or direct 3GPP download URLs to clipboard.")
        self.ribbon_card.add_primary_widget(self.btn_copy_tdoc)

        # Dedicated full-width wrapping alert banner for missing/incomplete TDoc metadata
        self.tdoc_metadata_banner = QFrame(self.ribbon_card)
        self.tdoc_metadata_banner.setObjectName("tdocMetadataBanner")
        self.tdoc_metadata_banner.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        self.tdoc_metadata_banner.setStyleSheet(
            f"""
            QFrame#tdocMetadataBanner {{
                background-color: {C.WARNING_SOFT};
                border: 1px solid {C.WARNING_BORDER};
                border-radius: {R.MD}px;
            }}
            """
        )
        self.tdoc_metadata_banner.setVisible(False)
        banner_layout = QHBoxLayout(self.tdoc_metadata_banner)
        banner_layout.setContentsMargins(S.SM + 2, S.XS + 1, S.SM + 2, S.XS + 1)
        banner_layout.setSpacing(S.SM)

        self.tdoc_metadata_warning = QLabel(self.tdoc_metadata_banner)
        self.tdoc_metadata_warning.setWordWrap(True)
        self.tdoc_metadata_warning.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.tdoc_metadata_warning.setStyleSheet(
            f"color: {C.WARNING_DARK}; font-size: {scale_px(T.SIZE_CAPTION)}px; "
            f"font-weight: 600; background: transparent; border: none;"
        )
        self.tdoc_metadata_warning.setVisible(False)
        banner_layout.addWidget(self.tdoc_metadata_warning, 1, Qt.AlignVCenter)

        self.btn_update_tdoc_metadata = QPushButton("Update metadata", self.tdoc_metadata_banner)
        self.btn_update_tdoc_metadata.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.btn_update_tdoc_metadata.setMinimumHeight(scale_px(26))
        self.btn_update_tdoc_metadata.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_update_tdoc_metadata.setCursor(Qt.PointingHandCursor)
        self.btn_update_tdoc_metadata.setVisible(False)
        apply_lucide_icon(self.btn_update_tdoc_metadata, "refresh-cw", color=C.WARNING_DARK, size=13)
        self.btn_update_tdoc_metadata.setToolTip("Deep-scrape Docs/ folders to rebuild missing first/last TDoc ranges.")
        self.btn_update_tdoc_metadata.clicked.connect(self._update_incomplete_tdoc_metadata)
        banner_layout.addWidget(self.btn_update_tdoc_metadata, 0, Qt.AlignVCenter)

        self.ribbon_card.add_banner_widget(self.tdoc_metadata_banner)

        # 3. Network Sync Menu & Inline Progress Banner
        self._setup_sync_controls()

        # 4. Reports & Analytics Menu
        self.reports_menu_btn = QPushButton("Reports ▾")
        self.reports_menu_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.reports_menu_btn.setMinimumHeight(scale_px(30))
        self.reports_menu_btn.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.reports_menu_btn, "bar-chart-2", size=14)
        self.reports_menu_btn.setToolTip("Reports, analytics, and consolidated exports...")

        reports_menu = QMenu(self)
        rep_act_topic = reports_menu.addAction(
            "🌐 Cross-WG Topic Dossier && Briefing Studio...", self._open_topic_dossier_studio
        )
        rep_act_topic.setToolTip(
            "Define multi-WG topics (Meta-Topics) and generate incremental cross-meeting dossiers with unified decisions and LLM feeds."
        )
        rep_act_agreements = reports_menu.addAction(
            "⚖️ Study Item && Meeting Agreements Tracker...", self._open_agreement_tracker
        )
        rep_act_agreements.setToolTip(
            "Track, search, and export agreements, working assumptions, FFS, and agreed TPs/pCRs across RAN2, RAN3, RAN1, and SA2 meetings."
        )

        rep_act_contrib = reports_menu.addAction("📊 Contributions Report...", self._open_contribution_report)
        rep_act_contrib.setToolTip("Filter and audit company contributions across multiple meetings.")

        rep_act_merge = reports_menu.addAction("📥 Export Merged TDocs (Excel)...", self._export_merged_tdocs)
        rep_act_merge.setToolTip("Download and consolidate TDoc lists from filtered meetings into a single Excel workbook.")

        self.reports_menu_btn.setMenu(reports_menu)
        self.ribbon_card.add_utility_widget(self.reports_menu_btn)
        self.btn_contribution_report = self.reports_menu_btn  # Compatibility alias
        self.btn_export_merged = self.reports_menu_btn  # Compatibility alias

        # 5. Settings Button
        self.btn_config = QPushButton("Settings")
        self.btn_config.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_config.setMinimumHeight(scale_px(30))
        self.btn_config.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.btn_config, "settings", size=14)
        self.btn_config.setToolTip("Open Meetings & Cache Configuration...")
        self.btn_config.clicked.connect(self.config_dialog.exec_)
        self.ribbon_card.add_utility_widget(self.btn_config)

        # --- TIER 2: Search & Filter Flow ---
        search_group = QWidget(self.ribbon_card)
        search_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)

        search_lbl = make_filter_label("Search:", emphasized=True)
        search_layout.addWidget(search_lbl)

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.search_input.setPlaceholderText("Filter number, city, or title...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumHeight(scale_px(28))
        self.search_input.setMinimumWidth(scale_px(210))
        self.search_input.setToolTip("Filter visible meetings by meeting number, host city, or title.")
        self.search_input.textChanged.connect(self.refresh_table)
        search_layout.addWidget(self.search_input)
        self.ribbon_card.add_filter_widget(search_group)

        # TSG Segmented Chips (All, SA, RAN, CT)
        tsg_group = QWidget(self.ribbon_card)
        tsg_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tsg_layout = QHBoxLayout(tsg_group)
        tsg_layout.setContentsMargins(0, 0, 0, 0)
        tsg_layout.setSpacing(4)

        tsg_lbl = make_filter_label("TSG:")
        tsg_layout.addWidget(tsg_lbl)

        self.tsg_buttons = {}
        for t_label, t_val in [("All", ""), ("SA", "SA"), ("RAN", "RAN"), ("CT", "CT")]:
            btn = QPushButton(t_label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setMinimumHeight(scale_px(26))
            btn.setToolTip(f"Filter meetings under TSG {t_label}" if t_val else "Show meetings across all TSGs")
            btn.clicked.connect(lambda _, tv=t_val: self._on_tsg_button_clicked(tv))
            self.tsg_buttons[t_val] = btn
            tsg_layout.addWidget(btn)
        self._update_tsg_button_styles()
        self.ribbon_card.add_filter_widget(tsg_group)

        # Working Group Selector
        wg_group = QWidget(self.ribbon_card)
        wg_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        wg_layout = QHBoxLayout(wg_group)
        wg_layout.setContentsMargins(0, 0, 0, 0)
        wg_layout.setSpacing(6)

        wg_lbl = make_filter_label("WG:")
        wg_layout.addWidget(wg_lbl)

        self.wg_filter = CheckableComboBox("Working Group")
        self.wg_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_filter.setMinimumHeight(scale_px(28))
        self.wg_filter.setToolTip("Filter meetings by Working Group (e.g., SA2, RAN1).")
        self.wg_filter.selectionChanged.connect(self.refresh_table)
        wg_layout.addWidget(self.wg_filter)
        self.ribbon_card.add_filter_widget(wg_group)

        # Ad-Hoc / In-Person Format Dropdowns
        nature_group = QWidget(self.ribbon_card)
        nature_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        nature_layout = QHBoxLayout(nature_group)
        nature_layout.setContentsMargins(0, 0, 0, 0)
        nature_layout.setSpacing(6)

        self.adhoc_filter = QComboBox()
        self.adhoc_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.adhoc_filter.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.adhoc_filter.addItems(["All Meetings", "Regular", "Ad-Hoc / BIS"])
        self.adhoc_filter.setMinimumHeight(scale_px(28))
        self.adhoc_filter.setToolTip("Filter by regular vs. ad-hoc meetings.")
        self.adhoc_filter.currentTextChanged.connect(self.refresh_table)
        nature_layout.addWidget(self.adhoc_filter)

        self.type_filter = QComboBox()
        self.type_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.type_filter.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.type_filter.addItems(["All Formats", "In-Person", "Electronic"])
        self.type_filter.setMinimumHeight(scale_px(28))
        self.type_filter.setToolTip("Filter by in-person vs. electronic meetings.")
        self.type_filter.currentTextChanged.connect(self.refresh_table)
        nature_layout.addWidget(self.type_filter)
        self.ribbon_card.add_filter_widget(nature_group)

        # Dates Filter: Preset Combo + Date Range Pickers
        dates_group = QWidget(self.ribbon_card)
        dates_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        dates_layout = QHBoxLayout(dates_group)
        dates_layout.setContentsMargins(0, 0, 0, 0)
        dates_layout.setSpacing(6)

        dates_lbl = make_filter_label("Dates:")
        dates_layout.addWidget(dates_lbl)

        self.date_preset_combo = QComboBox()
        self.date_preset_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.date_preset_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.date_preset_combo.addItems(
            ["All Dates", "Upcoming", "This Year", "Past 12 Months", "Custom Range...", "Unknown Dates"]
        )
        self.date_preset_combo.setMinimumHeight(scale_px(28))
        self.date_preset_combo.setToolTip(
            "Scope meetings by date presets, custom range, or view undated folders under Unknown Dates."
        )
        self.date_preset_combo.currentIndexChanged.connect(self._on_date_preset_changed)
        dates_layout.addWidget(self.date_preset_combo)

        # Hidden checkbox used internally for backward compatibility
        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.setVisible(False)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setMinimumHeight(scale_px(28))
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.setToolTip("Filter meetings starting on or after this date.")
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)
        self.date_from.setVisible(False)
        dates_layout.addWidget(self.date_from)

        self.arrow_lbl = QLabel("➔")
        self.arrow_lbl.setAlignment(Qt.AlignCenter)
        self.arrow_lbl.setStyleSheet(f"color: {C.TEXT_FAINT}; font-weight: bold;")
        self.arrow_lbl.setVisible(False)
        dates_layout.addWidget(self.arrow_lbl)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setMinimumHeight(scale_px(28))
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.setToolTip("Filter meetings ending on or before this date.")
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)
        self.date_to.setVisible(False)
        dates_layout.addWidget(self.date_to)
        self.ribbon_card.add_filter_widget(dates_group)

        # Reset Filters Button
        self.btn_reset_filters = QPushButton("Clear")
        self.btn_reset_filters.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_reset_filters.setMinimumHeight(scale_px(28))
        self.btn_reset_filters.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.btn_reset_filters, "rotate-ccw", size=13)
        self.btn_reset_filters.setToolTip("Reset all filters to defaults.")
        self.btn_reset_filters.clicked.connect(self._reset_filters)
        self.ribbon_card.add_filter_widget(self.btn_reset_filters)

        main_layout.addWidget(self.ribbon_card)

        # ==========================================
        # --- CARD: Full-Width Meetings Table & Footer ---
        # ==========================================
        self.table_card = QFrame(self)
        self.table_card.setObjectName("cardFrame")
        self.table_card.setStyleSheet(CARD_FRAME_STYLE)
        table_card_layout = QVBoxLayout(self.table_card)
        table_card_layout.setContentsMargins(0, 0, 0, 0)
        table_card_layout.setSpacing(0)

        self.table = QTableView(self.table_card)
        self.table.setMouseTracking(True)
        self.table_model = MeetingsTableModel()
        self.table.setModel(self.table_model)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        register_scaled_row_height(self.table, 36)
        self.table.setStyleSheet(f"""
            QTableView {{
                border: none;
                border-top-left-radius: {R.LG}px;
                border-top-right-radius: {R.LG}px;
                gridline-color: {C.BG_CANVAS};
                background-color: {C.BG_CARD};
                alternate-background-color: {C.BG_SURFACE};
                selection-background-color: {C.ACCENT_SOFT};
                selection-color: {C.TEXT_STRONG};
                font-size: {T.SIZE_TABLE}px;
            }}
            QTableView::item {{
                padding: 4px 6px;
                border-bottom: 1px solid {C.BORDER_ROW};
            }}
            QTableView::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.TEXT_STRONG};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                font-weight: bold;
                font-size: {T.SIZE_CAPTION}px;
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                padding: 6px 8px;
            }}
        """)

        # Column Delegates
        self.hover_delegate = DynamicHoverMenuDelegate(self)
        self.table.setItemDelegateForColumn(0, self.hover_delegate)

        self.tdocs_delegate = TDocsButtonDelegate(self)
        self.table.setItemDelegateForColumn(1, self.tdocs_delegate)

        self.wg_delegate = WGTagDelegate(self)
        self.table.setItemDelegateForColumn(2, self.wg_delegate)

        # Column Layout (7 clean columns)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.setSectionResizeMode(6, QHeaderView.Fixed)

        self._apply_scaled_table_columns()

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_right_click_menu)
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        table_card_layout.addWidget(self.table, 1)

        # Footer Status Strip & Floating Comparison Cart
        footer_bar = QFrame(self.table_card)
        footer_bar.setObjectName("meetingsTableFooter")
        footer_bar.setStyleSheet(f"""
            QFrame#meetingsTableFooter {{
                background-color: {C.BG_SURFACE};
                border: none;
                border-top: 1px solid {C.BORDER_SUBTLE};
                border-bottom-left-radius: {R.LG}px;
                border-bottom-right-radius: {R.LG}px;
            }}
        """)
        bottom_layout = QHBoxLayout(footer_bar)
        bottom_layout.setContentsMargins(S.MD, S.XS + 2, S.MD, S.XS + 2)
        bottom_layout.setSpacing(10)

        self.status_lbl = QLabel("Showing 0 meetings")
        self.status_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: {T.SIZE_CAPTION}px; background: transparent; border: none;")
        bottom_layout.addWidget(self.status_lbl)

        bottom_layout.addStretch()

        # Floating Comparison Cart Pill (hidden when empty) + Quick Stage Active Word button
        self.cart_frame = ComparisonCartWidget(self)
        self.btn_cart_launcher = self.cart_frame.create_launcher_button(self)
        self.lbl_slot_a = self.cart_frame.btn_slot_a
        self.lbl_slot_b = self.cart_frame.btn_slot_b
        self.btn_compare = self.cart_frame.btn_compare
        self.btn_clear_cart = self.cart_frame.btn_clear_cart

        bottom_layout.addWidget(self.btn_cart_launcher)
        bottom_layout.addWidget(self.cart_frame)

        table_card_layout.addWidget(footer_bar)
        main_layout.addWidget(self.table_card, 1)

        self._populate_filters()

