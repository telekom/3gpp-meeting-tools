"""UI layout construction mixin for TopicInputsViewerDialog."""

from typing import Dict

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.meetings.ui.dialogs.topic_inputs.styles import (
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
    STYLE_TABLE,
)
from modules.meetings.ui.topic_pills import FilterPillButton, get_wg_colors


class TopicInputsUiSetupMixin:
    """Builds the header card, KPI strip, filter pills, tabbed tables, and reader drawer."""

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # 1. Header Card with topic overview and timeframe
        header_card = QFrame()
        header_card.setObjectName("TopicInputsHeaderCard")
        header_card.setStyleSheet(
            f"QFrame#TopicInputsHeaderCard {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; }}"
        )
        h_layout = QVBoxLayout(header_card)
        h_layout.setContentsMargins(12, 10, 12, 10)
        h_layout.setSpacing(6)

        top_row = QHBoxLayout()
        lbl_title = QLabel(f"🌐 <b>Topic:</b> {self.topic.name}")
        lbl_title.setStyleSheet(f"font-size: 15px; font-weight: bold; color: {C.TEXT_PRIMARY}; background: transparent; border: none;")
        top_row.addWidget(lbl_title)

        top_row.addStretch()

        lbl_timeframe = QLabel(f"🗓️ <b>Scope:</b> {self.timeframe_label}")
        lbl_timeframe.setStyleSheet(
            f"color: {C.TEXT_SECONDARY}; font-size: 11px; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 3px 8px;"
        )
        top_row.addWidget(lbl_timeframe)
        h_layout.addLayout(top_row)

        desc_text = self.topic.description or "Cross-WG topic correlation"
        wgs_str = ", ".join(sorted(self.topic.target_wgs))
        lbl_desc = QLabel(f"<b>Target WGs:</b> <span style='color: {C.ACCENT_SKY_DARK};'>{wgs_str}</span> | {desc_text}")
        lbl_desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; background: transparent; border: none;")
        lbl_desc.setWordWrap(True)
        h_layout.addWidget(lbl_desc)

        layout.addWidget(header_card)

        # 2. Metric KPI Cards Strip
        kpi_frame = QFrame()
        kpi_frame.setObjectName("TopicInputsKpiFrame")
        kpi_frame.setStyleSheet(
            f"QFrame#TopicInputsKpiFrame {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
        )
        kpi_layout = QHBoxLayout(kpi_frame)
        kpi_layout.setContentsMargins(12, 8, 12, 8)
        kpi_layout.setSpacing(18)

        self.kpi_total = QLabel("📊 Total Inputs: <b>0</b>")
        self.kpi_total.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 12px; background: transparent; border: none;")
        kpi_layout.addWidget(self.kpi_total)

        self.kpi_decisions = QLabel("🟢 Decisions: <b>0</b>")
        self.kpi_decisions.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 12px; background: transparent; border: none;")
        kpi_layout.addWidget(self.kpi_decisions)

        self.kpi_ls = QLabel("📬 Liaisons: <b>0</b>")
        self.kpi_ls.setStyleSheet(f"color: {C.PURPLE_DARK}; font-size: 12px; background: transparent; border: none;")
        kpi_layout.addWidget(self.kpi_ls)

        self.kpi_disc = QLabel("📄 Discussions: <b>0</b>")
        self.kpi_disc.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; background: transparent; border: none;")
        kpi_layout.addWidget(self.kpi_disc)

        self.kpi_wgs = QLabel("👥 WGs: <b>0</b>")
        self.kpi_wgs.setStyleSheet(f"color: {C.ACCENT_SKY}; font-size: 12px; background: transparent; border: none;")
        kpi_layout.addWidget(self.kpi_wgs)

        kpi_layout.addStretch()

        self.btn_refresh = QPushButton("🔄 Re-Scan")
        self.btn_refresh.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_refresh.clicked.connect(self._start_collection)
        kpi_layout.addWidget(self.btn_refresh)

        layout.addWidget(kpi_frame)

        # Progress bar strip (indeterminate/step during collection)
        self.prog_frame = QFrame()
        self.prog_frame.setObjectName("TopicInputsProgFrame")
        self.prog_frame.setStyleSheet(
            f"QFrame#TopicInputsProgFrame {{ background-color: {C.INFO_SOFT}; border: 1px solid {C.INFO_BORDER}; border-radius: 6px; }}"
        )
        p_layout = QHBoxLayout(self.prog_frame)
        p_layout.setContentsMargins(10, 6, 10, 6)
        self.lbl_prog_status = QLabel("Scanning meeting records...")
        self.lbl_prog_status.setStyleSheet(
            f"color: {C.INFO_DARK}; font-size: 11px; font-weight: bold; background: transparent; border: none;"
        )
        p_layout.addWidget(self.lbl_prog_status, 1)
        self.prog_bar = QProgressBar()
        self.prog_bar.setFixedHeight(12)
        self.prog_bar.setRange(0, 100)
        self.prog_bar.setTextVisible(False)
        self.prog_bar.setStyleSheet(f"""
            QProgressBar {{ border: 1px solid {C.INFO_BORDER_MEDIUM}; border-radius: 3px; background-color: {C.INFO_SOFT_STRONG}; }}
            QProgressBar::chunk {{ background-color: {C.INFO_ACCENT}; border-radius: 2px; }}
        """)
        p_layout.addWidget(self.prog_bar, 1)
        layout.addWidget(self.prog_frame)

        # 3. Interactive Filter Strip & Quick-Filter Pill Bars
        filter_frame = QFrame()
        filter_frame.setObjectName("TopicInputsFilterFrame")
        filter_frame.setStyleSheet(
            f"QFrame#TopicInputsFilterFrame {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; }}"
        )
        filter_layout = QVBoxLayout(filter_frame)
        filter_layout.setContentsMargins(10, 8, 10, 8)
        filter_layout.setSpacing(8)

        # Row 1: Search + Meeting Round Selector + Reset
        search_row = QHBoxLayout()
        search_row.setSpacing(10)

        lbl_search = QLabel("<b>🔍 Search:</b>")
        lbl_search.setStyleSheet("background: transparent; border: none;")
        search_row.addWidget(lbl_search)
        self.edit_filter = QLineEdit()
        self.edit_filter.setPlaceholderText("Filter by TDoc ID, Title, Company, Agenda Item, or Decision text...")
        self.edit_filter.textChanged.connect(self._apply_filters)
        search_row.addWidget(self.edit_filter, 3)

        lbl_round = QLabel("<b>Round:</b>")
        lbl_round.setStyleSheet("background: transparent; border: none;")
        search_row.addWidget(lbl_round)
        self.combo_round_filter = QComboBox()
        self.combo_round_filter.addItem("All Meeting Rounds")
        for r in self.rounds:
            self.combo_round_filter.addItem(r.label)
        self.combo_round_filter.currentIndexChanged.connect(self._apply_filters)
        search_row.addWidget(self.combo_round_filter, 2)

        self.btn_reset_filters = QPushButton("↺ Reset")
        self.btn_reset_filters.setToolTip("Reset all search and filter pills")
        self.btn_reset_filters.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_reset_filters.clicked.connect(self._reset_filters)
        search_row.addWidget(self.btn_reset_filters)
        filter_layout.addLayout(search_row)

        # Dropdown combos maintained for programmatic filtering & backward compatibility
        self.combo_wg_filter = QComboBox()
        self.combo_wg_filter.addItem("All WGs")
        for wg in sorted(self.topic.target_wgs):
            self.combo_wg_filter.addItem(wg)
        self.combo_wg_filter.setVisible(False)

        self.combo_type_filter = QComboBox()
        self.combo_type_filter.addItems([
            "All Inputs",
            "🟢 Decisions Only (Agreements, pCRs, CRs)",
            "Text Agreements Only",
            "Approved pCRs Only",
            "Approved CRs Only",
            "📬 Liaison Statements Only",
            "📄 Discussion Papers Only",
        ])
        self.combo_type_filter.setVisible(False)

        # Row 2: Working Group Quick-Filter Pills (wrapping FlowLayout)
        wg_row = QHBoxLayout()
        wg_row.setSpacing(6)
        lbl_wg_title = QLabel("<b>Working Group:</b>")
        lbl_wg_title.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY}; background: transparent; border: none; padding-top: 4px;")
        lbl_wg_title.setMinimumWidth(96)
        wg_row.addWidget(lbl_wg_title, 0, Qt.AlignTop)

        wg_pills_container = QWidget()
        wg_sp = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        wg_sp.setHeightForWidth(True)
        wg_pills_container.setSizePolicy(wg_sp)
        wg_flow = FlowLayout(wg_pills_container, margin=0, h_spacing=6, v_spacing=4)

        self.wg_pills: Dict[str, FilterPillButton] = {}
        self.wg_pill_group = QButtonGroup(self)
        self.wg_pill_group.setExclusive(True)

        pill_all_wg = FilterPillButton("All WGs", "All WGs", count=None)
        pill_all_wg.setChecked(True)
        pill_all_wg.toggled.connect(lambda chk: self._on_wg_pill_toggled("All WGs", chk))
        self.wg_pills["All WGs"] = pill_all_wg
        self.wg_pill_group.addButton(pill_all_wg)
        wg_flow.addWidget(pill_all_wg)

        for wg in sorted(self.topic.target_wgs):
            colors = get_wg_colors(wg)
            pill = FilterPillButton(wg, wg, count=0, colors=colors)
            pill.toggled.connect(lambda chk, w=wg: self._on_wg_pill_toggled(w, chk))
            self.wg_pills[wg] = pill
            self.wg_pill_group.addButton(pill)
            wg_flow.addWidget(pill)
        wg_row.addWidget(wg_pills_container, 1)
        filter_layout.addLayout(wg_row)

        # Row 3: Decision & Input Category Quick-Filter Pills (wrapping FlowLayout)
        cat_row = QHBoxLayout()
        cat_row.setSpacing(6)
        lbl_cat_title = QLabel("<b>Category:</b>")
        lbl_cat_title.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY}; background: transparent; border: none; padding-top: 4px;")
        lbl_cat_title.setMinimumWidth(96)
        cat_row.addWidget(lbl_cat_title, 0, Qt.AlignTop)

        cat_pills_container = QWidget()
        cat_sp = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        cat_sp.setHeightForWidth(True)
        cat_pills_container.setSizePolicy(cat_sp)
        cat_flow = FlowLayout(cat_pills_container, margin=0, h_spacing=6, v_spacing=4)

        self.cat_pills: Dict[str, FilterPillButton] = {}
        self.cat_pill_group = QButtonGroup(self)
        self.cat_pill_group.setExclusive(True)

        cat_defs = [
            ("All Inputs", "All Inputs", None, None),
            ("🟢 Decisions Only (Agreements, pCRs, CRs)", "🟢 Decisions", 0, {"bg_active": C.EMERALD_ACCENT, "fg_active": C.BG_CARD, "border_active": C.EMERALD_PRIMARY, "bg_inactive": C.BG_CARD, "fg_inactive": C.EMERALD_DARK, "border_inactive": C.WG_RAN_BORDER, "hover": C.WG_RAN_BG}),
            ("Text Agreements Only", "📝 Agreements", 0, {"bg_active": C.EMERALD_ACCENT, "fg_active": C.BG_CARD, "border_active": C.EMERALD_PRIMARY, "bg_inactive": C.BG_CARD, "fg_inactive": C.WG_RAN_TEXT, "border_inactive": C.WG_RAN_BORDER, "hover": C.WG_RAN_BG}),
            ("Approved pCRs Only", "📑 Approved pCRs", 0, {"bg_active": C.EMERALD_PRIMARY, "fg_active": C.BG_CARD, "border_active": C.EMERALD_DARK, "bg_inactive": C.BG_CARD, "fg_inactive": C.WG_RAN_TEXT, "border_inactive": C.EMERALD_BORDER, "hover": C.WG_RAN_BG}),
            ("Approved CRs Only", "✅ Approved CRs", 0, {"bg_active": C.EMERALD_DARK, "fg_active": C.BG_CARD, "border_active": C.WG_RAN_TEXT, "bg_inactive": C.BG_CARD, "fg_inactive": C.EMERALD_DARK, "border_inactive": C.EMERALD_BORDER, "hover": C.WG_RAN_BG}),
            ("📬 Liaison Statements Only", "📬 Liaisons", 0, {"bg_active": C.VIOLET_VIBRANT, "fg_active": C.BG_CARD, "border_active": C.VIOLET_PRIMARY, "bg_inactive": C.BG_CARD, "fg_inactive": C.PURPLE_DARK, "border_inactive": C.VIOLET_BORDER, "hover": C.VIOLET_SURFACE}),
            ("📄 Discussion Papers Only", "📄 Discussions", 0, {"bg_active": C.TEXT_MUTED, "fg_active": C.BG_CARD, "border_active": C.TEXT_SECONDARY, "bg_inactive": C.BG_CARD, "fg_inactive": C.TEXT_BODY, "border_inactive": C.BORDER_STRONG, "hover": C.BG_SURFACE}),
        ]

        for cat_key, cat_label, init_cnt, custom_col in cat_defs:
            pill = FilterPillButton(cat_key, cat_label, count=init_cnt, colors=custom_col)
            if cat_key == "All Inputs":
                pill.setChecked(True)
            pill.toggled.connect(lambda chk, k=cat_key: self._on_cat_pill_toggled(k, chk))
            self.cat_pills[cat_key] = pill
            self.cat_pill_group.addButton(pill)
            cat_flow.addWidget(pill)
        cat_row.addWidget(cat_pills_container, 1)
        filter_layout.addLayout(cat_row)

        layout.addWidget(filter_frame)

        # 4. Tabbed Visualizations
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; background-color: {C.BG_CARD}; }}
            QTabBar::tab {{ padding: 6px 14px; font-weight: bold; font-size: 11px; }}
            QTabBar::tab:selected {{ background-color: {C.BG_CARD}; color: {C.EMERALD_ACCENT}; border-bottom: 2px solid {C.EMERALD_ACCENT}; }}
        """)

        # Tab 1: Decisions Table
        self.table_decisions = QTableWidget()
        self.table_decisions.setStyleSheet(STYLE_TABLE)
        self.table_decisions.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_decisions.verticalHeader().setVisible(False)
        self.table_decisions.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_decisions, "⚖️ Technical Decisions && Agreements (0)")

        # Tab 2: All Contributions Table
        self.table_contributions = QTableWidget()
        self.table_contributions.setStyleSheet(STYLE_TABLE)
        self.table_contributions.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_contributions.verticalHeader().setVisible(False)
        self.table_contributions.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_contributions, "📄 All Considered Contributions (0)")

        # Tab 3: Liaison Statements Table
        self.table_ls = QTableWidget()
        self.table_ls.setStyleSheet(STYLE_TABLE)
        self.table_ls.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_ls.verticalHeader().setVisible(False)
        self.table_ls.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_ls, "📬 Inter-WG Liaison Statements (0)")

        # Tab 4: WG Distribution Matrix
        self.table_dist = QTableWidget()
        self.table_dist.setStyleSheet(STYLE_TABLE)
        self.table_dist.verticalHeader().setVisible(False)
        self.tabs.addTab(self.table_dist, "📊 WG && Round Distribution")

        # 5. Collapsible Agreement & Decision Detail Reader Drawer
        self.drawer_widget = QFrame()
        self.drawer_widget.setObjectName("TopicInputsDrawerCard")
        self.drawer_widget.setStyleSheet(
            f"QFrame#TopicInputsDrawerCard {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; border-radius: 8px; }}"
        )
        drawer_layout = QVBoxLayout(self.drawer_widget)
        drawer_layout.setContentsMargins(10, 8, 10, 8)
        drawer_layout.setSpacing(6)

        hdr_layout = QHBoxLayout()
        self.lbl_drawer_title = QLabel("📖 Decision & Agreement Detail Reader")
        self.lbl_drawer_title.setStyleSheet(
            f"font-weight: bold; font-size: 12px; color: {C.TEXT_STRONG}; background: transparent; border: none;"
        )

        self.lbl_drawer_provenance = QLabel("")
        self.lbl_drawer_provenance.setStyleSheet(f"background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.SUCCESS_BORDER};")
        self.lbl_drawer_provenance.setVisible(False)

        self.btn_copy_agreement_text = QPushButton("📋 Copy Verbatim Text")
        self.btn_copy_agreement_text.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_copy_agreement_text.setToolTip("Copy the full verbatim agreement text or contribution summary to clipboard")
        self.btn_copy_agreement_text.clicked.connect(self._on_copy_detail_text)

        self.btn_toggle_drawer = QPushButton("▲ Collapse")
        self.btn_toggle_drawer.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_toggle_drawer.clicked.connect(self._toggle_drawer)

        hdr_layout.addWidget(self.lbl_drawer_title)
        hdr_layout.addWidget(self.lbl_drawer_provenance)
        hdr_layout.addStretch()
        hdr_layout.addWidget(self.btn_copy_agreement_text)
        hdr_layout.addWidget(self.btn_toggle_drawer)
        drawer_layout.addLayout(hdr_layout)

        self.detail_text_edit = QTextEdit()
        self.detail_text_edit.setReadOnly(True)
        self.detail_text_edit.setMinimumHeight(90)
        self.detail_text_edit.setMaximumHeight(220)
        self.detail_text_edit.setStyleSheet(f"""
            QTextEdit {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 8px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 11px;
                color: {C.TEXT_STRONG};
            }}
        """)
        self.detail_text_edit.setPlaceholderText("Select a decision, contribution, or liaison statement from the table above to read the full text, verbatim agreement wording, and technical provenance.")
        drawer_layout.addWidget(self.detail_text_edit)

        # Connect table selection change events
        self.table_decisions.itemSelectionChanged.connect(self._on_decision_selection_changed)
        self.table_contributions.itemSelectionChanged.connect(self._on_contribution_selection_changed)
        self.table_ls.itemSelectionChanged.connect(self._on_ls_selection_changed)
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # Vertical Splitter
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.tabs)
        self.splitter.addWidget(self.drawer_widget)
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 1)

        layout.addWidget(self.splitter, 1)

        # 6. Bottom Action Ribbon
        action_layout = QHBoxLayout()

        self.btn_open_doc = QPushButton("📄 Open Document")
        self.btn_open_doc.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_doc.clicked.connect(self._on_open_selected_document)
        action_layout.addWidget(self.btn_open_doc)

        self.btn_open_portal = QPushButton("🌐 Open on 3GPP Portal")
        self.btn_open_portal.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_portal.clicked.connect(self._on_open_selected_portal)
        action_layout.addWidget(self.btn_open_portal)

        self.btn_copy_md = QPushButton("📋 Copy Markdown Table")
        self.btn_copy_md.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_copy_md.clicked.connect(self._on_copy_markdown_table)
        action_layout.addWidget(self.btn_copy_md)

        self.btn_export_csv = QPushButton("💾 Export to CSV...")
        self.btn_export_csv.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_export_csv.clicked.connect(self._on_export_csv)
        action_layout.addWidget(self.btn_export_csv)

        action_layout.addStretch()

        self.btn_export_dossier = QPushButton("🚀 Proceed to Full Dossier Export...")
        self.btn_export_dossier.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_export_dossier.clicked.connect(self._on_proceed_to_export)
        action_layout.addWidget(self.btn_export_dossier)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_close.clicked.connect(self.close)
        action_layout.addWidget(btn_close)

        layout.addLayout(action_layout)
