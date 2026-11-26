"""Left panel builder mixin (Indexed Protocols card and Protocol Library list) for NASTab."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAction,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMenu,
    QProgressBar,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.page_ribbon import apply_lucide_icon
from core.ui.theme import C
from modules.nas.ui.nas_components import (
    NASMessageItemDelegate,
    NASVersionCardsWidget,
)
from ..styles import SEGMENTED_MODE_STYLE, TOGGLE_BUTTON_STYLE


class NASTabLeftPaneMixin:
    """Constructs the left-hand protocol selection and message/IE catalogue splitter pane."""

    def _build_left_pane(self) -> QWidget:
        # --- Left Panel (Vertical Splitter: Protocol Pill Cards + Protocol Library) ---
        left_widget = QWidget()
        self.left_pane_widget = left_widget
        left_widget.setMinimumWidth(scale_px(280))
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        self.left_splitter = QSplitter(Qt.Vertical)
        self.left_splitter.setChildrenCollapsible(False)

        # Indexed Protocols Card Frame (Spec-Search-aligned pill cards & inline header)
        ver_card = QFrame()
        self.ver_card = ver_card
        ver_card.setObjectName("nasSpecTreeCard")
        ver_card.setStyleSheet(f"""
            QFrame#nasSpecTreeCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)
        ver_card_layout = QVBoxLayout(ver_card)
        ver_card_layout.setContentsMargins(6, 8, 6, 6)
        ver_card_layout.setSpacing(6)

        # Card Header: Title + Spec Count Badge + Compact Import & Manage Dropdowns
        ver_hdr_layout = QHBoxLayout()
        ver_hdr_layout.setContentsMargins(2, 0, 2, 0)
        ver_hdr_layout.setSpacing(4)

        ver_title = QLabel("📚 Indexed Protocols")
        ver_title.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {C.TEXT_STRONG};")
        ver_hdr_layout.addWidget(ver_title)

        self.btn_toggle_specs_card = QPushButton("▾")
        self.btn_toggle_specs_card.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_specs_card.setToolTip("Collapse / expand Indexed Protocols list to free space for Protocol Library")
        self.btn_toggle_specs_card.setStyleSheet(
            f"QPushButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 1px 6px; font-size: 10px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; color: {C.TEXT_PRIMARY}; }}"
        )
        ver_hdr_layout.addWidget(self.btn_toggle_specs_card)

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
        ver_hdr_layout.addWidget(self.tree_stats_badge)
        ver_hdr_layout.addStretch()

        self.import_menu_btn = QPushButton("Import ▾")
        self.import_menu_btn.setCursor(Qt.PointingHandCursor)
        self.import_menu_btn.setToolTip("Import protocol specifications from Specs DB or local Word .docx")
        apply_lucide_icon(self.import_menu_btn, "download", color=C.ACCENT_PRIMARY, size=13)
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
        act_import_db = import_menu.addAction("📥 Import from Specifications DB...")
        act_import_db.triggered.connect(self._on_fetch_from_specs_db_clicked)
        act_import_docx = import_menu.addAction("📁 Import Local Word Document (.docx)...")
        act_import_docx.triggered.connect(self._on_import_local_file_clicked)
        self.import_menu_btn.setMenu(import_menu)
        ver_hdr_layout.addWidget(self.import_menu_btn)

        self.db_menu_btn = QPushButton("▾")
        self.db_menu_btn.setCursor(Qt.PointingHandCursor)
        self.db_menu_btn.setToolTip("Protocol database maintenance (reload versions, clear, wipe DB)")
        apply_lucide_icon(self.db_menu_btn, "sliders", color=C.TEXT_SECONDARY, size=13)
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
        manage_menu = QMenu(self)
        manage_menu.setStyleSheet(f"""
            QMenu {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }}
            QMenu::item {{
                padding: 6px 18px 6px 10px;
                border-radius: 4px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)
        act_reload_checked = QAction("🔄 Reload Checked Version(s)...", self)
        act_reload_checked.triggered.connect(self._on_reload_checked_clicked)
        manage_menu.addAction(act_reload_checked)

        act_clear = QAction("🗑️ Clear Selected Version / Spec...", self)
        act_clear.triggered.connect(self._on_clear_version_clicked)
        manage_menu.addAction(act_clear)

        manage_menu.addSeparator()

        act_wipe = QAction("⚠️ Wipe Protocol Database...", self)
        act_wipe.triggered.connect(self._on_wipe_db_clicked)
        manage_menu.addAction(act_wipe)

        self.db_menu_btn.setMenu(manage_menu)
        ver_hdr_layout.addWidget(self.db_menu_btn)

        # Compatibility aliases for ingestion controller and unit tests
        self.fetch_btn = self.import_menu_btn
        self.import_file_btn = self.import_menu_btn
        self.manage_btn = self.db_menu_btn

        ver_card_layout.addLayout(ver_hdr_layout)

        # Compact 1-line summary bar shown when Indexed Protocols card is collapsed
        self.ver_summary_bar = QFrame(ver_card)
        self.ver_summary_bar.setObjectName("verSummaryBar")
        self.ver_summary_bar.setStyleSheet(
            f"#verSummaryBar {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; }}"
        )
        sum_lay = QHBoxLayout(self.ver_summary_bar)
        sum_lay.setContentsMargins(6, 3, 6, 3)
        sum_lay.setSpacing(6)
        self.ver_summary_lbl = QLabel("No specs checked", self.ver_summary_bar)
        self.ver_summary_lbl.setStyleSheet(f"font-size: 10.5px; color: {C.TEXT_SECONDARY};")
        sum_lay.addWidget(self.ver_summary_lbl, 1)
        self.btn_change_specs = QPushButton("Change ▾", self.ver_summary_bar)
        self.btn_change_specs.setCursor(Qt.PointingHandCursor)
        self.btn_change_specs.setStyleSheet(
            f"QPushButton {{ font-size: 10px; font-weight: 600; color: {C.ACCENT_PRIMARY}; "
            f"background-color: {C.SKY_SURFACE}; border: 1px solid {C.SKY_BORDER}; border-radius: 3px; padding: 1px 6px; }} "
            f"QPushButton:hover {{ background-color: {C.SKY_SOFT}; }}"
        )
        self.btn_change_specs.clicked.connect(lambda: self.btn_toggle_specs_card.click())
        sum_lay.addWidget(self.btn_change_specs)
        self.ver_summary_bar.setVisible(False)
        ver_card_layout.addWidget(self.ver_summary_bar)

        self.ver_body_widget = QWidget(ver_card)
        ver_body_layout = QVBoxLayout(self.ver_body_widget)
        ver_body_layout.setContentsMargins(0, 0, 0, 0)
        ver_body_layout.setSpacing(6)

        # Sleek 4px inline progress bar for background ingestion
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
        ver_body_layout.addWidget(self.progress_bar)

        # Quick Spec Filter Input
        self.tree_filter_input = QLineEdit()
        self.tree_filter_input.setPlaceholderText("Filter specs (e.g. 24.501, NGAP, RRC)...")
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
        ver_body_layout.addWidget(self.tree_filter_input)

        # Quick Filter Action Chips (matching Spec Search)
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(4)

        chip_neutral_style = f"""
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
        """

        btn_all = QPushButton("All")
        btn_all.setToolTip("Select all indexed releases across all protocol specifications")
        btn_all.setCursor(Qt.PointingHandCursor)
        btn_all.setStyleSheet(chip_neutral_style)
        chips_layout.addWidget(btn_all)

        btn_none = QPushButton("None")
        btn_none.setToolTip("Deselect all protocol releases")
        btn_none.setCursor(Qt.PointingHandCursor)
        btn_none.setStyleSheet(chip_neutral_style)
        chips_layout.addWidget(btn_none)

        btn_latest_rel = QPushButton("⭐ Latest/Rel")
        btn_latest_rel.setToolTip("Check only the newest version of each Release per specification")
        btn_latest_rel.setCursor(Qt.PointingHandCursor)
        btn_latest_rel.setStyleSheet(f"""
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
        chips_layout.addWidget(btn_latest_rel)

        btn_latest = QPushButton("Latest")
        btn_latest.setToolTip("Select only the single newest version per specification")
        btn_latest.setCursor(Qt.PointingHandCursor)
        btn_latest.setStyleSheet(chip_neutral_style)
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
        ver_body_layout.addLayout(chips_layout)

        self.version_tree = NASVersionCardsWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.structure_changed.connect(self._save_config)
        self.version_tree.reload_version_requested.connect(self._reload_single_version)
        self.version_tree.reload_spec_requested.connect(self._reload_spec_group)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        self.version_tree.open_spec_requested.connect(
            lambda s, v, c, fmt: self.open_specification(s, v, c, fmt)
        )
        self.version_tree.show_in_specs_requested.connect(self.navigate_to_specs_tab_requested)
        self.version_tree.search_in_spec_search_requested.connect(self.search_in_spec_search_requested)

        self.tree_filter_input.textChanged.connect(self.version_tree.filter_tree)
        btn_all.clicked.connect(self.version_tree.select_all)
        btn_none.clicked.connect(self.version_tree.deselect_all)
        btn_latest_rel.clicked.connect(self.version_tree.select_latest_per_release)
        btn_latest.clicked.connect(self.version_tree.select_latest_only)
        btn_r18.clicked.connect(lambda: self.version_tree.select_releases_min(18))

        ver_body_layout.addWidget(self.version_tree)
        ver_card_layout.addWidget(self.ver_body_widget)
        self.left_splitter.addWidget(ver_card)

        # Protocol Library Card Frame
        msg_card = QFrame()
        msg_card.setObjectName("cardFrame")
        msg_card.setStyleSheet(f"""
            QFrame#cardFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)
        msg_card_layout = QVBoxLayout(msg_card)
        msg_card_layout.setContentsMargins(0, 0, 0, 0)
        msg_card_layout.setSpacing(0)

        msg_header = QFrame()
        msg_header.setObjectName("msgLibHeader")
        msg_header.setStyleSheet(f"""
            QFrame#msgLibHeader {{
                background-color: {C.BG_SURFACE};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }}
        """)
        msg_header_layout = QHBoxLayout(msg_header)
        msg_header_layout.setContentsMargins(6, 4, 6, 4)
        msg_header_layout.setSpacing(6)

        msg_card_title = QLabel("📨 Protocol Library")
        msg_card_title.setStyleSheet(f"font-weight: 600; font-size: 11px; color: {C.TEXT_STRONG};")
        msg_header_layout.addWidget(msg_card_title)
        msg_header_layout.addStretch()

        self.msg_count_lbl = QLabel("0 items")
        self.msg_count_lbl.setStyleSheet(f"""
            color: {C.TEXT_MUTED};
            font-size: 10px;
            font-weight: 500;
            background-color: {C.BG_CANVAS};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: 9px;
            padding: 1px 7px;
        """)
        msg_header_layout.addWidget(self.msg_count_lbl)
        msg_card_layout.addWidget(msg_header)

        msg_body = QWidget()
        msg_body_layout = QVBoxLayout(msg_body)
        msg_body_layout.setContentsMargins(6, 6, 6, 6)
        msg_body_layout.setSpacing(5)

        # Sleek Segmented Mode Bar: [📨 Messages | 🧩 All IEs / Types]
        mode_frame = QFrame()
        mode_frame.setObjectName("modeSegmentedFrame")
        mode_frame.setStyleSheet(f"""
            QFrame#modeSegmentedFrame {{
                background-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 5px;
            }}
        """)
        mode_bar_layout = QHBoxLayout(mode_frame)
        mode_bar_layout.setContentsMargins(2, 2, 2, 2)
        mode_bar_layout.setSpacing(2)

        self.mode_messages_btn = QPushButton("📨 Messages")
        self.mode_messages_btn.setCheckable(True)
        self.mode_messages_btn.setChecked(True)
        self.mode_messages_btn.setCursor(Qt.PointingHandCursor)
        self.mode_messages_btn.setMinimumHeight(scale_px(24))
        self.mode_messages_btn.setToolTip("Browse Protocol Messages & SIBs (Clause 8 / ASN.1 PDUs)")
        self.mode_messages_btn.setStyleSheet(SEGMENTED_MODE_STYLE)
        self.mode_messages_btn.clicked.connect(lambda: self._set_browse_mode("messages"))
        mode_bar_layout.addWidget(self.mode_messages_btn)

        self.mode_ies_btn = QPushButton("🧩 All IEs / Types")
        self.mode_ies_btn.setCheckable(True)
        self.mode_ies_btn.setChecked(False)
        self.mode_ies_btn.setCursor(Qt.PointingHandCursor)
        self.mode_ies_btn.setMinimumHeight(scale_px(24))
        self.mode_ies_btn.setToolTip(
            "Browse all parsed Information Elements & Types directly (including standalone & nested Clause 9 containers)"
        )
        self.mode_ies_btn.setStyleSheet(SEGMENTED_MODE_STYLE)
        self.mode_ies_btn.clicked.connect(lambda: self._set_browse_mode("ies"))
        mode_bar_layout.addWidget(self.mode_ies_btn)

        msg_body_layout.addWidget(mode_frame)

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
        self.deep_search_btn.setMinimumHeight(scale_px(26))
        self.deep_search_btn.setStyleSheet(TOGGLE_BUTTON_STYLE)
        self.deep_search_btn.toggled.connect(self._on_deep_search_toggled)
        ie_search_bar_layout.addWidget(self.deep_search_btn)
        msg_body_layout.addLayout(ie_search_bar_layout)

        self.msg_list = QListWidget()
        self.msg_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.msg_list.setItemDelegate(NASMessageItemDelegate(self.msg_list))
        self.msg_list.itemClicked.connect(self._on_message_clicked)
        self.msg_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.msg_list.customContextMenuRequested.connect(self._on_msg_list_context_menu)
        self.msg_list.setStyleSheet(f"""
            QListWidget {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                padding: 2px;
                outline: none;
            }}
        """)
        msg_body_layout.addWidget(self.msg_list)
        msg_card_layout.addWidget(msg_body)
        self.left_splitter.addWidget(msg_card)
        self.left_splitter.setSizes([270, 430])
        left_layout.addWidget(self.left_splitter)
        return left_widget
