"""UI layout construction and KPI summary chips for `MeetingAgreementsDialog`."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSplitter,
    QTableView,
    QTextBrowser,
    QVBoxLayout,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)


class MeetingAgreementsUiSetupMixin:
    """Mixin responsible for building widgets, filters, master-detail panes, and KPI chips."""

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # 1. Header Ribbon with Title & KPI Chips
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 8px;
                padding: 6px 12px;
            }}
        """)
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(4, 4, 4, 4)

        title_layout = QVBoxLayout()
        title_lbl = QLabel(f"<b>⚖️ Agreements & Working Assumptions</b> — {self.meeting_name}")
        title_lbl.setStyleSheet(f"font-size: 15px; color: {C.TEXT_PRIMARY}; font-weight: bold;")
        sub_lbl = QLabel("Chairman's Notes extracted agreements, decisions, and working assumptions")
        sub_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        title_layout.addWidget(title_lbl)
        title_layout.addWidget(sub_lbl)
        header_layout.addLayout(title_layout)

        header_layout.addStretch()

        # KPI Badges
        self.kpi_layout = QHBoxLayout()
        self.kpi_layout.setSpacing(8)
        header_layout.addLayout(self.kpi_layout)

        main_layout.addWidget(header_frame)

        # 2. Search & Filter Bar
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search agreements text, TDoc ID, AI, WI, or company...")
        self.search_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.search_input, stretch=3)

        # Kind Filter Pills
        self.combo_kind = QComboBox()
        self.combo_kind.addItem("All Kinds", "ALL")
        has_pcr = any("pcr" in str(a.get("kind", "")).lower() for a in self.raw_agreements)
        has_cr = any(
            str(a.get("kind", "")).lower().endswith("cr") and "pcr" not in str(a.get("kind", "")).lower()
            for a in self.raw_agreements
        )
        has_ls = any("ls" in str(a.get("kind", "")).lower() for a in self.raw_agreements)

        self.combo_kind.addItem("🟢 Technical Agreements", "AGREEMENT")
        if has_pcr:
            self.combo_kind.addItem("📝 Agreed pCRs Only", "PCR")
        if has_cr:
            self.combo_kind.addItem("📄 Approved CRs Only", "CR")
        if has_ls:
            self.combo_kind.addItem("📤 Approved LS Only", "LS")
        self.combo_kind.addItem("🟡 Working Assumptions Only", "WA")
        self.combo_kind.addItem("🔵 FFS / Open Issues Only", "FFS")
        self.combo_kind.setStyleSheet(f"""
            QComboBox {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
                font-weight: 500;
                color: {C.TEXT_STRONG};
            }}
        """)
        self.combo_kind.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_kind)

        # AI Filter Combo
        self.combo_ai = QComboBox()
        self.combo_ai.addItem("All Agenda Items", "ALL")
        unique_ais = sorted({str(ag.get("ai", "")).strip() for ag in self.raw_agreements if ag.get("ai")})
        for ai in unique_ais:
            self.combo_ai.addItem(f"AI {ai}", ai)
        self.combo_ai.setStyleSheet(self.combo_kind.styleSheet())
        self.combo_ai.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_ai)

        main_layout.addLayout(filter_layout)

        # 3. Master-Detail Splitter
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        # Master: Agreements Table
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.setAlternatingRowColors(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(f"""
            QTableView {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                gridline-color: {C.BG_CANVAS};
                selection-background-color: {C.INFO_SOFT_STRONG};
                selection-color: {C.INFO_NAVY};
                font-size: 12px;
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                font-weight: bold;
                padding: 6px;
                border: 1px solid {C.BORDER_SUBTLE};
            }}
        """)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 150)
        header.resizeSection(1, 110)
        header.resizeSection(2, 90)
        header.resizeSection(3, 130)
        header.setSectionResizeMode(4, QHeaderView.Stretch)

        self.table.selectionModel().selectionChanged.connect(self._on_selection_changed)
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        splitter.addWidget(self.table)

        # Detail: Rich Viewer Pane
        detail_container = QFrame()
        detail_container.setStyleSheet(f"""
            QFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(10, 8, 10, 8)
        detail_layout.setSpacing(6)

        detail_bar = QHBoxLayout()
        self.detail_title_lbl = QLabel("<b>Agreement Details</b>")
        self.detail_title_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_PRIMARY};")
        detail_bar.addWidget(self.detail_title_lbl)
        detail_bar.addStretch()

        self.btn_copy_single = QPushButton("📋 Copy Agreement")
        self.btn_copy_single.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_single.setToolTip("Copy this agreement text with formal citation to clipboard.")
        self.btn_copy_single.clicked.connect(self._copy_current_agreement)
        detail_bar.addWidget(self.btn_copy_single)

        self.btn_locate_table = QPushButton("🎯 Show in Meeting Table")
        self.btn_locate_table.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_locate_table.setToolTip("Scroll to and highlight this TDoc in the main meeting window.")
        self.btn_locate_table.clicked.connect(self._locate_current_tdoc)
        detail_bar.addWidget(self.btn_locate_table)

        detail_layout.addLayout(detail_bar)

        self.viewer = QTextBrowser()
        self.viewer.setOpenExternalLinks(True)
        self.viewer.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 4px;
                padding: 10px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                color: {C.TEXT_PRIMARY};
            }}
        """)
        detail_layout.addWidget(self.viewer)

        splitter.addWidget(detail_container)
        splitter.setSizes([260, 260])
        main_layout.addWidget(splitter, stretch=1)

        # 4. Bottom Action Bar
        bottom_layout = QHBoxLayout()

        self.btn_copy_all = QPushButton("📋 Copy All (Markdown)")
        self.btn_copy_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_all.setToolTip("Copy all currently visible agreements formatted as Markdown.")
        self.btn_copy_all.clicked.connect(self._copy_all_markdown)
        bottom_layout.addWidget(self.btn_copy_all)

        self.btn_export_md = QPushButton("💾 Export to Agreements.md...")
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.setToolTip("Save agreements to a standalone Markdown (.md) document.")
        self.btn_export_md.clicked.connect(self._export_to_markdown_file)
        bottom_layout.addWidget(self.btn_export_md)

        bottom_layout.addStretch()

        self.status_count_lbl = QLabel("")
        self.status_count_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-weight: bold;")
        bottom_layout.addWidget(self.status_count_lbl)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_close.clicked.connect(self.accept)
        bottom_layout.addWidget(btn_close)

        main_layout.addLayout(bottom_layout)
        self._update_status_count()

    def _update_kpi_counters(self):
        """Builds visually appealing KPI chip badges at the top of the dialog."""
        while self.kpi_layout.count():
            item = self.kpi_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        tot = len(self.raw_agreements)
        pcr_cnt = sum(
            1 for a in self.raw_agreements
            if "pcr" in str(a.get("kind", "")).lower() or "p-cr" in str(a.get("kind", "")).lower()
        )
        cr_cnt = sum(
            1 for a in self.raw_agreements
            if str(a.get("kind", "")).lower().endswith("cr") and "pcr" not in str(a.get("kind", "")).lower()
        )
        ls_cnt = sum(1 for a in self.raw_agreements if "ls" in str(a.get("kind", "")).lower())
        wa_cnt = sum(1 for a in self.raw_agreements if "working assumption" in str(a.get("kind", "")).lower())
        ffs_cnt = sum(
            1 for a in self.raw_agreements
            if "ffs" in str(a.get("kind", "")).lower() or "open issue" in str(a.get("kind", "")).lower()
        )
        ag_cnt = max(0, tot - pcr_cnt - cr_cnt - ls_cnt - wa_cnt - ffs_cnt)
        ai_cnt = len({str(a.get("ai", "")).strip() for a in self.raw_agreements if a.get("ai")})

        def _make_chip(label: str, count: int, bg: str, fg: str, border: str) -> QLabel:
            lbl = QLabel(f"<b>{count}</b> {label}")
            lbl.setStyleSheet(f"""
                QLabel {{
                    background-color: {bg};
                    color: {fg};
                    border: 1px solid {border};
                    border-radius: 4px;
                    padding: 3px 8px;
                    font-size: 11px;
                }}
            """)
            return lbl

        self.kpi_layout.addWidget(_make_chip("Total", tot, C.BG_CANVAS, C.TEXT_BODY, C.BORDER_STRONG))
        if ag_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("Agreements", ag_cnt, C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM)
            )
        if pcr_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("Agreed pCRs", pcr_cnt, C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM)
            )
        if cr_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("Approved CRs", cr_cnt, C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM)
            )
        if ls_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("Approved LS", ls_cnt, C.INDIGO_SOFT, C.WG_SA_TEXT, C.WG_SA_BORDER)
            )
        if wa_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("Working Assumptions", wa_cnt, C.WARNING_SOFT_ALT, C.WARNING_DARK, C.WARNING_BORDER_STRONG)
            )
        if ffs_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("FFS / Open", ffs_cnt, C.INFO_SOFT_STRONG, C.INFO_DARK, C.INFO_BORDER_MEDIUM)
            )
        if ai_cnt > 0:
            self.kpi_layout.addWidget(
                _make_chip("AIs Covered", ai_cnt, C.VIOLET_SOFT, C.PURPLE_DARK, C.VIOLET_BORDER)
            )
