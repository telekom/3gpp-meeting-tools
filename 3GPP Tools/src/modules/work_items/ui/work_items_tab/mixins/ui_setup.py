"""UI setup and font-scale responsiveness mixin for WorkItemsTab."""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QShortcut,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon, make_filter_label
from core.ui.theme import C, R, S, T
from core.ui.ui_components import (
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
from modules.work_items.ui.work_items_delegates import (
    ReleaseBadgeDelegate,
    RemarksDelegate,
    SpecsDelegate,
    StatusBadgeDelegate,
    TimelineDelegate,
    TrackingDelegate,
    WgBadgeDelegate,
    WidDelegate,
)
from modules.work_items.ui.work_items_model import WorkItemsTableModel


class WorkItemsUiSetupMixin:
    """Ribbon, filter strip, table construction, and responsive column sizing."""

    def _apply_scaled_table_columns(self):
        """Scale Work Items table column widths to match the current font scale."""
        widths = (
            (0, 50),
            (1, 102),
            (2, 124),
            (4, 54),
            (5, 130),
            (6, 64),
            (7, 106),
            (8, 50),
            (9, 174),
            (10, 104),
        )
        for col_idx, base_w in widths:
            self.table.setColumnWidth(col_idx, max(base_w, scale_px(base_w)))

    def on_font_scale_changed(self, _scale: float = 0.0):
        """Adapt filter boxes, dropdowns, buttons, and fixed table columns for the new font scale."""
        self.search_input.setMinimumWidth(scale_px(210))
        self.release_combo.setMinimumWidth(scale_px(130))
        self.wg_combo.setMinimumWidth(scale_px(120))
        self.status_combo.setMinimumWidth(scale_px(140))
        for btn in (self.add_wi_btn, self.sync_tracked_btn, self.sync_btn):
            btn.setMinimumHeight(scale_px(30))
        for w in (self.search_input, self.release_combo, self.wg_combo, self.status_combo):
            w.updateGeometry()
        self._apply_scaled_table_columns()
        self.table.viewport().update()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(S.MD, S.SM + 2, S.MD, S.SM + 2)
        main_layout.setSpacing(S.SM + 2)

        # ==========================================
        # --- 2-TIER CARDED PAGE RIBBON ---
        # ==========================================
        self.ribbon_card = PageRibbonCard("3GPP Work Items", parent=self)

        # Tier 1 Status / Progress indicator + Utility Actions
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-style: italic; font-size: {T.SIZE_CAPTION}px;")
        self.ribbon_card.add_primary_widget(self.status_lbl)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(180)
        self.ribbon_card.add_primary_widget(self.progress_bar)

        self.add_wi_btn = QPushButton("Add WI")
        self.add_wi_btn.setMinimumHeight(scale_px(30))
        self.add_wi_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.add_wi_btn, "plus", size=14)
        self.add_wi_btn.setToolTip("Add and scrape an individual Work Item by ID or Portal URL without running a full sync.")
        self.add_wi_btn.setCursor(Qt.PointingHandCursor)
        self.add_wi_btn.clicked.connect(lambda: self._open_add_wi_dialog())
        self.ribbon_card.add_utility_widget(self.add_wi_btn)

        self.sync_tracked_btn = QPushButton("Update Tracked (0)")
        self.sync_tracked_btn.setMinimumHeight(scale_px(30))
        self.sync_tracked_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.sync_tracked_btn, "refresh-cw", size=14)
        self.sync_tracked_btn.setToolTip("Quickly update details and remarks for only your Starred and Following Work Items.")
        self.sync_tracked_btn.setCursor(Qt.PointingHandCursor)
        self.sync_tracked_btn.clicked.connect(self._sync_tracked_wis)
        self.ribbon_card.add_utility_widget(self.sync_tracked_btn)

        self.sync_btn = QPushButton("Sync 3GPP WIs")
        self.sync_btn.setObjectName("primaryBtn")
        self.sync_btn.setMinimumHeight(scale_px(30))
        apply_lucide_icon(self.sync_btn, "globe", color=C.TEXT_ON_ACCENT, active_color=C.TEXT_ON_ACCENT, size=14)
        self.sync_btn.setToolTip("Click to download and synchronize 3GPP Work Items in parallel from the server.")
        self.sync_btn.clicked.connect(self._start_sync)
        self.ribbon_card.add_utility_widget(self.sync_btn)

        # Tier 2: Inline Search, TSG Quick Scope & Faceted Filters
        search_group = QWidget(self.ribbon_card)
        search_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(6)
        search_layout.addWidget(make_filter_label("Search:", emphasized=True))

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.search_input.setPlaceholderText("Code, Acronym, or Name...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(scale_px(210))
        self.search_input.setToolTip("Filter the table instantly by typing keywords or WI code.")
        self.search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.search_input)
        self.ribbon_card.add_filter_widget(search_group)

        tsg_group = QWidget(self.ribbon_card)
        tsg_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tsg_layout = QHBoxLayout(tsg_group)
        tsg_layout.setContentsMargins(0, 0, 0, 0)
        tsg_layout.setSpacing(4)
        tsg_layout.addWidget(make_filter_label("TSG:"))

        self.tsg_buttons = {}
        for tsg in ["All", "SA", "RAN", "CT"]:
            btn = QPushButton(tsg)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(f"Filter Work Items to {tsg if tsg != 'All' else 'all'} Technical Specification Groups")
            btn.clicked.connect(lambda _, t=tsg: self._on_tsg_clicked(t))
            self.tsg_buttons[tsg] = btn
            tsg_layout.addWidget(btn)
        self.tsg_buttons["All"].setChecked(True)
        self.ribbon_card.add_filter_widget(tsg_group)

        combos_group = QWidget(self.ribbon_card)
        combos_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        combos_layout = QHBoxLayout(combos_group)
        combos_layout.setContentsMargins(0, 0, 0, 0)
        combos_layout.setSpacing(6)

        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(scale_px(130))
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        combos_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(scale_px(120))
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        combos_layout.addWidget(self.wg_combo)

        self.status_combo = QComboBox()
        self.status_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.status_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.status_combo.setToolTip("Filter by Work Item completion status (End Date)")
        self.status_combo.setMinimumWidth(scale_px(140))
        self.status_combo.addItem("🟢 Active / Ongoing", "active")
        self.status_combo.addItem("All Statuses", "all")
        self.status_combo.addItem("🏁 Finished", "finished")
        self.status_combo.currentIndexChanged.connect(lambda _: self.search_timer.start())
        combos_layout.addWidget(self.status_combo)

        self.clear_filters_btn = QPushButton("Clear")
        self.clear_filters_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.clear_filters_btn, "rotate-ccw", size=13)
        self.clear_filters_btn.setToolTip("Reset all search fields and active filters.")
        self.clear_filters_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filters_btn.clicked.connect(self._clear_all_filters)
        combos_layout.addWidget(self.clear_filters_btn)
        self.ribbon_card.add_filter_widget(combos_group)

        main_layout.addWidget(self.ribbon_card)

        # ==========================================
        # --- CARD: Tracking Chips Strip & Data Table ---
        # ==========================================
        self.table_card = QFrame(self)
        self.table_card.setObjectName("cardFrame")
        self.table_card.setStyleSheet(CARD_FRAME_STYLE)
        table_card_layout = QVBoxLayout(self.table_card)
        table_card_layout.setContentsMargins(0, 0, 0, 0)
        table_card_layout.setSpacing(0)

        chips_strip = QFrame(self.table_card)
        chips_strip.setObjectName("wiHeaderStrip")
        chips_strip.setStyleSheet(f"""
            QFrame#wiHeaderStrip {{
                background-color: {C.BG_SURFACE};
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: {R.LG}px;
                border-top-right-radius: {R.LG}px;
            }}
        """)
        chips_layout = QHBoxLayout(chips_strip)
        chips_layout.setContentsMargins(S.MD, S.XS + 2, S.MD, S.XS + 2)
        chips_layout.setSpacing(6)

        chips_layout.addWidget(make_filter_label("Filter:"))

        # Following Filter Chip (Assigned Duty)
        self.following_chip = QPushButton("🎯 Following (0)")
        self.following_chip.setCheckable(True)
        self.following_chip.setCursor(Qt.PointingHandCursor)
        self.following_chip.setStyleSheet(TOGGLE_PILL_STYLE)
        self.following_chip.setToolTip("Toggle view to show only Work Items you are assigned to follow.")
        self.following_chip.clicked.connect(self._toggle_following_filter)
        chips_layout.addWidget(self.following_chip)

        # Starred Filter Chip (Personal Interest)
        self.starred_chip = QPushButton("⭐ Starred (0)")
        self.starred_chip.setCheckable(True)
        self.starred_chip.setCursor(Qt.PointingHandCursor)
        self.starred_chip.setStyleSheet(FAV_PILL_STYLE)
        self.starred_chip.setToolTip("Toggle view to show only your starred/favorite Work Items.")
        self.starred_chip.clicked.connect(self._toggle_starred_filter)
        chips_layout.addWidget(self.starred_chip)

        div_lbl = QLabel("|")
        div_lbl.setStyleSheet(f"color: {C.BORDER_STRONG}; margin: 0 4px; background: transparent; border: none;")
        chips_layout.addWidget(div_lbl)

        # Active Dismissible Tags Container
        self.active_tags_widget = QWidget(chips_strip)
        self.active_tags_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        self.active_tags_container_layout = QHBoxLayout(self.active_tags_widget)
        self.active_tags_container_layout.setContentsMargins(0, 0, 0, 0)
        self.active_tags_container_layout.setSpacing(4)
        chips_layout.addWidget(self.active_tags_widget)

        # Clear all active tags button
        self.clear_all_btn = QPushButton("Reset All")
        self.clear_all_btn.setCursor(Qt.PointingHandCursor)
        self.clear_all_btn.setStyleSheet(CLEAR_ALL_BTN_STYLE)
        apply_lucide_icon(self.clear_all_btn, "rotate-ccw", color=C.DANGER_ALT, size=12)
        self.clear_all_btn.setToolTip("Clear all active search and filter constraints.")
        self.clear_all_btn.setVisible(False)
        self.clear_all_btn.clicked.connect(self._clear_all_filters)
        chips_layout.addWidget(self.clear_all_btn)

        chips_layout.addStretch()

        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet(f"font-weight: bold; color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; background: transparent; border: none;")
        chips_layout.addWidget(self.count_label)

        table_card_layout.addWidget(chips_strip)

        # --- TABLE VIEW ---
        self.table = QTableView(self.table_card)
        self.table_model = WorkItemsTableModel()
        self.table_model.set_tracking_sets(self.starred, self.following)
        self.table.setModel(self.table_model)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        register_scaled_row_height(self.table, 34)
        self.table.setShowGrid(True)
        self.table.setStyleSheet(f"""
            QTableView {{
                background-color: {C.BG_CARD};
                alternate-background-color: {C.BG_SURFACE};
                border: none;
                border-bottom-left-radius: {R.LG}px;
                border-bottom-right-radius: {R.LG}px;
                gridline-color: {C.BORDER_ROW};
                selection-background-color: {C.ACCENT_SOFT};
                selection-color: {C.TEXT_STRONG};
                font-size: {T.SIZE_TABLE}px;
            }}
            QTableView::item {{
                padding: 4px 6px;
                border-bottom: 1px solid {C.BORDER_ROW};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                font-weight: bold;
                font-size: {T.SIZE_TABLE}px;
                border: none;
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-right: 1px solid {C.BORDER_SUBTLE};
                padding: 6px 6px;
            }}
        """)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.doubleClicked.connect(self._on_row_double_clicked)
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)

        # Column 0: Track (Fixed)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.tracking_delegate = TrackingDelegate(self.table)
        self.tracking_delegate.star_toggled.connect(self._toggle_star)
        self.tracking_delegate.follow_toggled.connect(self._toggle_follow)
        self.table.setItemDelegateForColumn(0, self.tracking_delegate)

        # Column 3: Name (Stretch — primary WI title gets remaining width)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        self.wg_delegate = WgBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(4, self.wg_delegate)

        self.wid_delegate = WidDelegate(self.table)
        self.wid_delegate.link_clicked.connect(self._log_and_forward_action)
        self.table.setItemDelegateForColumn(5, self.wid_delegate)

        self.release_delegate = ReleaseBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(6, self.release_delegate)

        self.status_delegate = StatusBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(7, self.status_delegate)

        self.specs_delegate = SpecsDelegate(self.table)
        self.specs_delegate.specs_clicked.connect(self._on_specs_clicked)
        self.table.setItemDelegateForColumn(8, self.specs_delegate)

        self.timeline_delegate = TimelineDelegate(self.table)
        self.table.setItemDelegateForColumn(9, self.timeline_delegate)

        self.remarks_delegate = RemarksDelegate(self.table)
        self.table.setItemDelegateForColumn(10, self.remarks_delegate)

        self._apply_scaled_table_columns()

        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll)
        table_card_layout.addWidget(self.table, 1)
        main_layout.addWidget(self.table_card, 1)
