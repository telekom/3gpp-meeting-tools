"""Right panel builder and top-level UI setup mixin (Matrix + Inspector) for NASTab."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListView,
    QMenu,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import register_scaled_row_height
from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon
from core.ui.theme import C
from core.ui.ui_components import COMBOBOX_STYLE_TOOLBAR
from modules.nas.ui.nas_components import NASInspectorWidget, NASMatrixTreeDelegate
from ..layout_controller import ProtocolWorkspaceLayoutController


class NASTabRightPaneMixin:
    """Builds the main horizontal splitter and right-hand Matrix/Inspector panes."""

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Hidden backward-compatibility ribbon & KPI labels for legacy callers/tests;
        # primary Import / Manage controls live compactly inside the Indexed Protocols card header.
        self.ribbon_card = PageRibbonCard("Protocols", parent=self)
        self.ribbon_card.setVisible(False)
        self.stats_specs_lbl = QLabel("📦 0 Specs", self.ribbon_card)
        self.stats_specs_lbl.setVisible(False)
        self.stats_vers_lbl = QLabel("🏷️ 0 Versions", self.ribbon_card)
        self.stats_vers_lbl.setVisible(False)
        self.stats_msgs_lbl = QLabel("📨 0 Items", self.ribbon_card)
        self.stats_msgs_lbl.setVisible(False)

        # ---------------------------------------------------------------------
        # 1. Main Horizontal Splitter (Left: Indexed Protocols + Library, Right: Matrix + Inspector)
        # ---------------------------------------------------------------------
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.addWidget(self._build_left_pane())
        self.main_splitter.addWidget(self._build_right_pane())
        self.main_splitter.setSizes([380, 700])
        layout.addWidget(self.main_splitter, 1)

        self.layout_controller = ProtocolWorkspaceLayoutController(
            owner=self,
            main_splitter=self.main_splitter,
            left_pane=self.left_pane_widget,
            left_splitter=self.left_splitter,
            ver_card=self.ver_card,
            ver_body_widget=self.ver_body_widget,
            ver_summary_bar=self.ver_summary_bar,
            ver_summary_lbl=self.ver_summary_lbl,
            btn_toggle_specs_card=self.btn_toggle_specs_card,
            right_splitter=self.right_splitter,
            matrix_card=self.matrix_card,
            matrix_header_layout=self.matrix_header_layout,
            matrix_stack=self.matrix_stack,
            msg_context_banner=self.msg_context_banner,
            inspector=self.inspector,
        )

    def _build_right_pane(self) -> QSplitter:
        # --- Right Panel (Matrix Card + Inspector Card) ---
        right_splitter = QSplitter(Qt.Vertical)
        self.right_splitter = right_splitter

        matrix_card = QFrame()
        self.matrix_card = matrix_card
        matrix_card.setObjectName("cardFrame")
        matrix_card.setStyleSheet(f"""
            QFrame#cardFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)
        matrix_card_layout = QVBoxLayout(matrix_card)
        matrix_card_layout.setContentsMargins(0, 0, 0, 0)
        matrix_card_layout.setSpacing(0)

        matrix_header_frame = QFrame()
        matrix_header_frame.setObjectName("matrixHeaderFrame")
        matrix_header_frame.setStyleSheet(f"""
            QFrame#matrixHeaderFrame {{
                background-color: {C.BG_SURFACE};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }}
        """)
        matrix_header_layout = QHBoxLayout(matrix_header_frame)
        matrix_header_layout.setContentsMargins(6, 4, 6, 4)
        matrix_header_layout.setSpacing(6)
        self.matrix_header_layout = matrix_header_layout

        self.matrix_title = QLabel("Select a Message to View Evolution Matrix")
        self.matrix_title.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {C.TEXT_STRONG};")
        matrix_header_layout.addWidget(self.matrix_title)

        # Interactive Diff Filter Pills
        diff_legend = QHBoxLayout()
        diff_legend.setSpacing(4)

        def _make_diff_pill(text: str, bg: str, fg: str, border: str, active_border: str, tip: str) -> QPushButton:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip(tip)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {bg};
                    color: {fg};
                    font-size: 10px;
                    font-weight: bold;
                    border: 1px solid {border};
                    border-radius: 10px;
                    padding: 2px 8px;
                }}
                QPushButton:hover {{
                    border-color: {active_border};
                }}
                QPushButton:checked {{
                    border: 2px solid {active_border};
                }}
            """)
            btn.toggled.connect(lambda _checked: self._rebuild_matrix_model())
            return btn

        self.btn_diff_added = _make_diff_pill(
            "🟢 Added",
            C.DIFF_ADDED_BG,
            C.SUCCESS_DARK,
            C.SUCCESS_BORDER,
            C.SUCCESS,
            "Toggle filter: show only fields added between selected releases",
        )
        diff_legend.addWidget(self.btn_diff_added)

        self.btn_diff_modified = _make_diff_pill(
            "🟡 Modified",
            C.WARNING_YELLOW_SOFT,
            C.WARNING_YELLOW_TEXT,
            C.HIGHLIGHT_YELLOW,
            C.WARNING_ACCENT,
            "Toggle filter: show only fields modified between selected releases",
        )
        diff_legend.addWidget(self.btn_diff_modified)

        self.btn_diff_removed = _make_diff_pill(
            "🔴 Removed",
            C.DANGER_SOFT_HOVER,
            C.DANGER_DARK,
            C.DANGER_BORDER,
            C.DANGER,
            "Toggle filter: show only fields removed between selected releases",
        )
        diff_legend.addWidget(self.btn_diff_removed)

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

        self.depth_combo = QComboBox()
        self.depth_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.depth_combo.setToolTip("Limit nested Information Element expansion depth in the matrix")
        self.depth_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.depth_combo.addItem("All Levels", None)
        self.depth_combo.addItem("Top-Level Only", 0)
        self.depth_combo.addItem("Depth ≤ 1", 1)
        self.depth_combo.addItem("Depth ≤ 2", 2)
        self.depth_combo.setVisible(False)
        self.depth_combo.currentIndexChanged.connect(lambda _: self._rebuild_matrix_model())
        matrix_header_layout.addWidget(self.depth_combo)

        self.open_spec_btn = QPushButton("Open Spec ▾")
        self.open_spec_btn.setCursor(Qt.PointingHandCursor)
        self.open_spec_btn.setToolTip("Open active 3GPP specification document in Microsoft Word (Clause copied to clipboard)")
        self.open_spec_btn.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 500;
                color: {C.ACCENT_PRIMARY};
                background-color: {C.SKY_SURFACE};
                border: 1px solid {C.SKY_BORDER};
                border-radius: 4px;
                padding: 2px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_SOFT};
                border-color: {C.ACCENT_SKY};
                color: {C.ACCENT_SKY_DARK};
            }}
            QPushButton:disabled {{
                color: {C.TEXT_FAINT};
                background-color: {C.BG_SURFACE};
                border-color: {C.BORDER_SUBTLE};
            }}
        """)
        apply_lucide_icon(self.open_spec_btn, "book-open", color=C.ACCENT_PRIMARY)
        self.open_spec_btn.setEnabled(False)
        self.open_spec_btn.clicked.connect(self._on_open_spec_btn_clicked)
        matrix_header_layout.addWidget(self.open_spec_btn)

        self.copy_matrix_btn = QPushButton("📋 Copy Table ▾")
        self.copy_matrix_btn.setCursor(Qt.PointingHandCursor)
        self.copy_matrix_btn.setToolTip("Copy entire evolution matrix to clipboard (Word-compatible table, Markdown table, or TSV)")
        self.copy_matrix_btn.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 500;
                color: {C.TEXT_BODY};
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_PRIMARY};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton:disabled {{
                color: {C.TEXT_FAINT};
                background-color: {C.BG_SURFACE};
                border-color: {C.BORDER_SUBTLE};
            }}
        """)
        apply_lucide_icon(self.copy_matrix_btn, "copy")
        self.copy_matrix_btn.setEnabled(False)

        copy_matrix_menu = QMenu(self)
        copy_matrix_menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)
        act_cp_word = copy_matrix_menu.addAction("📋 Copy for Word / Outlook (Rich Table)")
        act_cp_word.triggered.connect(lambda: self._copy_matrix_formatted(mode="word"))
        act_cp_md = copy_matrix_menu.addAction("📝 Copy as Markdown Table")
        act_cp_md.triggered.connect(lambda: self._copy_matrix_formatted(mode="markdown"))
        copy_matrix_menu.addSeparator()
        act_cp_tsv = copy_matrix_menu.addAction("📊 Copy as TSV (Excel / Plain Text)")
        act_cp_tsv.triggered.connect(lambda: self._copy_matrix_formatted(mode="tsv"))

        self.copy_matrix_btn.setMenu(copy_matrix_menu)
        matrix_header_layout.addWidget(self.copy_matrix_btn)

        matrix_card_layout.addWidget(matrix_header_frame)

        # Message Context Info Banner (Direction pill + concise prose summary)
        self.msg_context_banner = QFrame()
        self.msg_context_banner.setObjectName("msgContextBanner")
        self.msg_context_banner.setStyleSheet(f"""
            QFrame#msgContextBanner {{
                background-color: {C.BG_SURFACE};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
            }}
        """)
        msg_ctx_layout = QHBoxLayout(self.msg_context_banner)
        msg_ctx_layout.setContentsMargins(10, 5, 10, 5)
        msg_ctx_layout.setSpacing(8)

        self.msg_direction_badge = QLabel()
        self.msg_direction_badge.setStyleSheet(f"""
            QLabel {{
                background-color: {C.SKY_SOFT};
                color: {C.ACCENT_SKY_DARK};
                font-size: 11px;
                font-weight: 600;
                border: 1px solid {C.SKY_BORDER};
                border-radius: 4px;
                padding: 2px 8px;
            }}
        """)
        self.msg_direction_badge.setVisible(False)
        msg_ctx_layout.addWidget(self.msg_direction_badge, 0, Qt.AlignTop)

        self.msg_meta_badge = QLabel()
        self.msg_meta_badge.setStyleSheet(f"""
            QLabel {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                font-size: 10px;
                font-weight: 600;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 7px;
            }}
        """)
        self.msg_meta_badge.setVisible(False)
        msg_ctx_layout.addWidget(self.msg_meta_badge, 0, Qt.AlignTop)

        self.msg_description_lbl = QLabel()
        self.msg_description_lbl.setWordWrap(True)
        self.msg_description_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.msg_description_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px; line-height: 1.35;")
        self.msg_description_lbl.setVisible(False)
        msg_ctx_layout.addWidget(self.msg_description_lbl, 1)

        self.msg_context_banner.setVisible(False)
        matrix_card_layout.addWidget(self.msg_context_banner)

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
        m_empty_title.setStyleSheet(f"font-weight: bold; font-size: 13px; color: {C.TEXT_STRONG};")
        m_empty_title.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.addWidget(m_empty_title)

        m_empty_sub = QLabel(
            "Choose a message or IE from the library on the left to inspect its evolution across 3GPP releases.\n"
            "Field presence, ASN.1 syntax, and format changes will be diff-highlighted here."
        )
        m_empty_sub.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; line-height: 1.4;")
        m_empty_sub.setAlignment(Qt.AlignCenter)
        matrix_empty_layout.addWidget(m_empty_sub)

        self.matrix_stack.addWidget(self.matrix_empty_widget)

        self.matrix_table = QTableView()
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QTableView.SelectRows)
        self.matrix_table.setSelectionMode(QTableView.SingleSelection)
        register_scaled_row_height(self.matrix_table, 24)
        self.matrix_table.verticalHeader().setVisible(False)
        self.matrix_table.setStyleSheet(f"""
            QTableView {{
                border: none;
                gridline-color: {C.BG_CANVAS};
                background-color: {C.BG_CARD};
                alternate-background-color: {C.BG_SURFACE};
            }}
            QTableView::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.TEXT_PRIMARY};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_BODY};
                font-weight: bold;
                font-size: 11px;
                border: none;
                border-bottom: 1px solid {C.BORDER_STRONG};
                border-right: 1px solid {C.BORDER_SUBTLE};
                padding: 4px 6px;
            }}
        """)

        h_header = self.matrix_table.horizontalHeader()
        h_header.setHighlightSections(False)
        h_header.setStretchLastSection(False)

        self.matrix_table.clicked.connect(self._on_table_cell_clicked)
        self.matrix_table.doubleClicked.connect(self._on_table_cell_double_clicked)
        self.matrix_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.matrix_table.customContextMenuRequested.connect(self._on_matrix_context_menu)
        self.matrix_table.setItemDelegateForColumn(1, NASMatrixTreeDelegate(self.matrix_table))
        self.matrix_stack.addWidget(self.matrix_table)

        matrix_card_layout.addWidget(self.matrix_stack)
        right_splitter.addWidget(matrix_card)

        # Inspector Card (Bottom Right)
        self.inspector = NASInspectorWidget(db=self.db)
        self.inspector.jump_to_message_requested.connect(self._jump_to_message)
        self.inspector.filter_by_ie_requested.connect(self._on_filter_by_ie_requested)
        self.inspector.import_spec_requested.connect(self._on_import_cross_ref_spec_requested)
        self.inspector.open_spec_requested.connect(
            lambda s, v, c, fmt: self.open_specification(s, v, c, fmt)
        )
        self.inspector.show_in_specs_tab_requested.connect(self.navigate_to_specs_tab_requested)
        self.inspector.search_in_spec_search_requested.connect(self.search_in_spec_search_requested)
        self.inspector.send_plantuml_to_editor_requested.connect(self.send_plantuml_to_editor_requested)
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([380, 320])
        return right_splitter
