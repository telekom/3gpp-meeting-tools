"""UI setup mixin for SpecTDocReportDialog."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import BADGE_STYLE_PRIMARY, BUTTON_STYLE_TOOLBAR_SECONDARY


class SpecTDocReportUiSetupMixin:
    """Builds the header KPI ribbon, filter controls, contribution table, drawer, and Markdown tab."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # -------------------------------------------------------------
        # 1. HEADER RIBBON & KPI SUMMARY CARD
        # -------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        spec_pill = QLabel(f"<b>TS {self.spec_number}</b>")
        spec_pill.setStyleSheet(BADGE_STYLE_PRIMARY + " font-size: 14px; padding: 4px 10px;")

        ver_pill = QLabel(f"<b>v{self.version}</b>")
        ver_pill.setStyleSheet(f"""
            background-color: {C.WG_RAN_BG};
            color: {C.WG_RAN_TEXT};
            border: 1px solid {C.WG_RAN_BORDER};
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
            padding: 3px 8px;
        """)

        title_text = self.spec_title or "Integrated Technical Contributions Dossier"
        self.title_lbl = QLabel(title_text)
        self.title_lbl.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.TEXT_STRONG};")
        self.title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top_row.addWidget(spec_pill)
        top_row.addWidget(ver_pill)
        top_row.addWidget(self.title_lbl, stretch=1)

        # KPI Badges
        total_cnt = len(self.all_entries)
        sources = {r.get("source") for r in self.all_entries if r.get("source")}
        approved_cnt = sum(
            1
            for r in self.all_entries
            if "approv" in r.get("status", "") or "agree" in r.get("status", "")
        )

        self.kpi_lbl = QLabel(
            f"<span style='color: {C.TEXT_SECONDARY};'>Total:</span> <b>{total_cnt}</b> TDocs &nbsp;|&nbsp; "
            f"<span style='color: {C.TEXT_SECONDARY};'>Sources:</span> <b>{len(sources)}</b> Companies &nbsp;|&nbsp; "
            f"<span style='color: {C.SUCCESS_DARK};'>Approved/Agreed:</span> <b>{approved_cnt}</b>"
        )
        self.kpi_lbl.setStyleSheet(
            f"font-size: 12px; background-color: {C.BG_CANVAS}; padding: 4px 10px; border-radius: 4px;"
        )
        top_row.addWidget(self.kpi_lbl)

        header_layout.addLayout(top_row)

        # Filter bar
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            "🔍 Filter by TDoc, title, company, agenda item, keyword..."
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumHeight(scale_px(28))
        self.search_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 12px;
            }}
            QLineEdit:focus {{
                border-color: {C.ACCENT_SKY};
            }}
        """)
        self.search_input.textChanged.connect(self._apply_filters)
        filter_row.addWidget(self.search_input, stretch=2)

        filter_row.addWidget(QLabel("<b>Company:</b>"))
        self.company_combo = QComboBox()
        self.company_combo.setMinimumHeight(scale_px(28))
        self.company_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.company_combo, stretch=1)

        filter_row.addWidget(QLabel("<b>Type:</b>"))
        self.type_combo = QComboBox()
        self.type_combo.setMinimumHeight(scale_px(28))
        self.type_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.type_combo)

        filter_row.addWidget(QLabel("<b>Status:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.setMinimumHeight(scale_px(28))
        self.status_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.status_combo)

        header_layout.addLayout(filter_row)
        main_layout.addWidget(header_card)

        # -------------------------------------------------------------
        # 2. MAIN TABBED VIEWER: [Table View] | [Markdown Dossier]
        # -------------------------------------------------------------
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabBar::tab {
                font-weight: bold;
                padding: 6px 16px;
            }
        """)

        # Tab 1: Interactive Table
        tab_table = QWidget()
        tab_table_layout = QVBoxLayout(tab_table)
        tab_table_layout.setContentsMargins(0, 8, 0, 0)
        tab_table_layout.setSpacing(8)

        # Splitter between Table and Detail Drawer
        self.splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            [
                "TDoc",
                "Type",
                "Status",
                "Source (Company)",
                "AI",
                "Title",
                "Technical Abstract / Remarks",
            ]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                gridline-color: {C.BG_CANVAS};
                font-size: 12px;
            }}
            QTableWidget::item:selected {{
                background-color: {C.SKY_SOFT};
                color: {C.TEXT_PRIMARY};
            }}
        """)

        h_hdr = self.table.horizontalHeader()
        h_hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(3, QHeaderView.Interactive)
        h_hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(5, QHeaderView.Stretch)
        h_hdr.setSectionResizeMode(6, QHeaderView.Stretch)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)

        self.splitter.addWidget(self.table)

        # Bottom Abstract / Solution Drawer
        self.drawer_frame = QFrame()
        self.drawer_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
            }}
        """)
        drawer_layout = QVBoxLayout(self.drawer_frame)
        drawer_layout.setContentsMargins(10, 8, 10, 8)
        drawer_layout.setSpacing(4)

        drawer_hdr = QHBoxLayout()
        self.drawer_title_lbl = QLabel("<b>Technical Details &amp; Proposed Solution</b>")
        self.drawer_title_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_STRONG};")
        drawer_hdr.addWidget(self.drawer_title_lbl)
        drawer_hdr.addStretch()

        self.btn_open_tdoc_drawer = QPushButton("📂 Open TDoc (.docx)")
        self.btn_open_tdoc_drawer.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_tdoc_drawer.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc_drawer.clicked.connect(self._open_selected_tdoc)
        drawer_hdr.addWidget(self.btn_open_tdoc_drawer)

        drawer_layout.addLayout(drawer_hdr)

        self.drawer_text = QTextEdit()
        self.drawer_text.setReadOnly(True)
        self.drawer_text.setPlaceholderText(
            "Select any contribution above to inspect its technical abstract, proposed solution, and secretary remarks..."
        )
        self.drawer_text.setStyleSheet(f"""
            QTextEdit {{
                border: none;
                background-color: {C.BG_SURFACE};
                font-size: 12px;
                color: {C.TEXT_BODY};
                line-height: 1.4;
            }}
        """)
        drawer_layout.addWidget(self.drawer_text)

        self.splitter.addWidget(self.drawer_frame)
        self.splitter.setSizes([450, 150])

        tab_table_layout.addWidget(self.splitter)
        self.tab_widget.addTab(tab_table, "📊 Interactive Contribution Table")

        # Tab 2: Markdown Dossier
        tab_md = QWidget()
        tab_md_layout = QVBoxLayout(tab_md)
        tab_md_layout.setContentsMargins(8, 8, 8, 8)
        tab_md_layout.setSpacing(8)

        self.md_browser = QTextEdit()
        self.md_browser.setReadOnly(True)
        self.md_browser.setPlainText(self.markdown_report)
        self.md_browser.setStyleSheet(f"""
            QTextEdit {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                color: {C.TEXT_STRONG};
            }}
        """)
        tab_md_layout.addWidget(self.md_browser)
        self.tab_widget.addTab(tab_md, "📝 Technical Dossier (Markdown)")

        main_layout.addWidget(self.tab_widget, stretch=1)

        # -------------------------------------------------------------
        # 3. FOOTER ACTIONS
        # -------------------------------------------------------------
        footer_layout = QHBoxLayout()

        self.count_lbl = QLabel(
            f"Showing {len(self.filtered_entries)} of {len(self.all_entries)} TDocs"
        )
        self.count_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_MUTED};")
        footer_layout.addWidget(self.count_lbl)
        footer_layout.addStretch()

        self.btn_copy_md = QPushButton("📋 Copy Markdown Dossier")
        self.btn_copy_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_md.setCursor(Qt.PointingHandCursor)
        self.btn_copy_md.clicked.connect(self._copy_markdown)
        footer_layout.addWidget(self.btn_copy_md)

        self.btn_export_md = QPushButton("💾 Save Markdown (.md)")
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.setCursor(Qt.PointingHandCursor)
        self.btn_export_md.clicked.connect(self._export_markdown_file)
        footer_layout.addWidget(self.btn_export_md)

        self.btn_export_csv = QPushButton("📊 Export CSV")
        self.btn_export_csv.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_csv.setCursor(Qt.PointingHandCursor)
        self.btn_export_csv.clicked.connect(self._export_csv_file)
        footer_layout.addWidget(self.btn_export_csv)

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.close)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)
