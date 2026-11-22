"""UI setup mixin for SpecSmartImportDialog."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C


class _SmartImportUiMixin:
    """Builds search bar, filter chips, version delta table, and footer for smart import."""

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Header Frame: Search input & spec info
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
            }}
        """)
        h_layout = QVBoxLayout(header_frame)
        h_layout.setContentsMargins(10, 10, 10, 10)
        h_layout.setSpacing(8)

        # Search Bar Row
        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        lbl_prompt = QLabel("🔍 Specification:")
        lbl_prompt.setStyleSheet(f"font-weight: bold; font-size: 12px; color: {C.TEXT_STRONG};")
        search_row.addWidget(lbl_prompt)

        self.spec_input = QLineEdit()
        self.spec_input.setPlaceholderText("Enter spec number (e.g. 23.501, 38.300, 24.501, 38.211)...")
        self.spec_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 5px 8px;
                font-size: 12px;
                font-weight: 500;
                color: {C.TEXT_PRIMARY};
            }}
            QLineEdit:focus {{
                border-color: {C.ACCENT_SKY};
            }}
        """)
        self.spec_input.setClearButtonEnabled(True)
        self.spec_input.returnPressed.connect(self._on_find_releases)
        search_row.addWidget(self.spec_input, stretch=1)

        self.btn_find = QPushButton("🔍 Find Releases")
        self.btn_find.setCursor(Qt.PointingHandCursor)
        self.btn_find.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.ACCENT_SKY};
                color: {C.BG_CARD};
                border: 1px solid {C.ACCENT_SKY};
                border-radius: 4px;
                padding: 6px 14px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C.ACCENT_SKY_DARK};
            }}
        """)
        self.btn_find.clicked.connect(self._on_find_releases)
        search_row.addWidget(self.btn_find)
        h_layout.addLayout(search_row)

        # Spec Metadata Strip
        self.info_strip = QWidget()
        info_layout = QHBoxLayout(self.info_strip)
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(8)

        self.lbl_spec_title = QLabel("Enter a 3GPP specification number above to load available releases.")
        self.lbl_spec_title.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px; font-weight: 500;")
        self.lbl_spec_title.setWordWrap(True)
        info_layout.addWidget(self.lbl_spec_title, stretch=1)

        self.lbl_status_badge = QLabel("")
        self.lbl_status_badge.setStyleSheet(f"""
            background-color: {C.BG_CANVAS};
            color: {C.TEXT_BODY};
            border: 1px solid {C.BORDER_STRONG};
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 10px;
            font-weight: bold;
        """)
        self.lbl_status_badge.setVisible(False)
        info_layout.addWidget(self.lbl_status_badge)

        h_layout.addWidget(self.info_strip)
        layout.addWidget(header_frame)

        # Quick Filter Chips Bar
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(6)

        lbl_select = QLabel("Select:")
        lbl_select.setStyleSheet(f"font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY};")
        chips_layout.addWidget(lbl_select)

        self.btn_chip_updates = QPushButton("⚡ Updates && New Only")
        self.btn_chip_updates.setToolTip("Select only releases that have updates or are not yet indexed")
        self.btn_chip_updates.setCursor(Qt.PointingHandCursor)
        self.btn_chip_updates.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 600;
                color: {C.ORANGE_DARK};
                background-color: {C.WG_CT_BG};
                border: 1px solid {C.ORANGE_BORDER};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.ORANGE_SOFT};
                border-color: {C.ORANGE_BORDER};
            }}
        """)
        self.btn_chip_updates.clicked.connect(self._select_updates_and_new)
        chips_layout.addWidget(self.btn_chip_updates)

        self.btn_chip_all = QPushButton("⭐ All Releases")
        self.btn_chip_all.setToolTip("Select the latest version of all releases")
        self.btn_chip_all.setCursor(Qt.PointingHandCursor)
        self.btn_chip_all.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 600;
                color: {C.ACCENT_SKY_DARK};
                background-color: {C.SKY_SURFACE};
                border: 1px solid {C.SKY_BORDER};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_SOFT};
            }}
        """)
        self.btn_chip_all.clicked.connect(self._select_all)
        chips_layout.addWidget(self.btn_chip_all)

        self.btn_chip_r18 = QPushButton("Rel-18+")
        self.btn_chip_r18.setToolTip("Select Release 18 and newer (5G-Advanced & 6G)")
        self.btn_chip_r18.setCursor(Qt.PointingHandCursor)
        self.btn_chip_r18.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 600;
                color: {C.INDIGO_DARK};
                background-color: {C.WG_SA_BG};
                border: 1px solid {C.WG_SA_BORDER};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.INDIGO_SOFT};
            }}
        """)
        self.btn_chip_r18.clicked.connect(lambda: self._select_min_release(18))
        chips_layout.addWidget(self.btn_chip_r18)

        self.btn_chip_r17 = QPushButton("Rel-17+")
        self.btn_chip_r17.setToolTip("Select Release 17 and newer")
        self.btn_chip_r17.setCursor(Qt.PointingHandCursor)
        self.btn_chip_r17.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 600;
                color: {C.TEXT_SECONDARY};
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
            }}
        """)
        self.btn_chip_r17.clicked.connect(lambda: self._select_min_release(17))
        chips_layout.addWidget(self.btn_chip_r17)

        self.btn_chip_clear = QPushButton("◻️ Clear")
        self.btn_chip_clear.setToolTip("Deselect all releases")
        self.btn_chip_clear.setCursor(Qt.PointingHandCursor)
        self.btn_chip_clear.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 500;
                color: {C.TEXT_MUTED};
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_PRIMARY};
            }}
        """)
        self.btn_chip_clear.clicked.connect(self._clear_selection)
        chips_layout.addWidget(self.btn_chip_clear)

        chips_layout.addStretch()

        self.lbl_selection_summary = QLabel("Selected: 0 release(s)")
        self.lbl_selection_summary.setStyleSheet(f"font-weight: bold; font-size: 11px; color: {C.ACCENT_SKY};")
        chips_layout.addWidget(self.lbl_selection_summary)

        layout.addLayout(chips_layout)

        # Table of Available Releases
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Select", "Release", "Latest Version", "Status", "Release Date", "Filename"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                gridline-color: {C.BG_CANVAS};
                font-size: 11px;
            }}
            QTableWidget::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.TEXT_PRIMARY};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                border: none;
                border-bottom: 1px solid {C.BORDER_STRONG};
                padding: 4px 6px;
                font-weight: bold;
                font-size: 11px;
                color: {C.TEXT_BODY};
            }}
        """)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        layout.addWidget(self.table)

        # Options & Ingestion Bar
        opt_layout = QHBoxLayout()
        self.chk_auto_supersede = QCheckBox("Auto-supersede older minor versions within the same release (recommended)")
        self.chk_auto_supersede.setChecked(True)
        self.chk_auto_supersede.setToolTip(
            "When checked, ingesting newer versions (e.g. v20.3.0) will automatically remove older\n"
            "superseded versions (e.g. v20.2.0) of that release from the search index."
        )
        self.chk_auto_supersede.setStyleSheet(f"font-size: 11px; color: {C.TEXT_BODY}; font-weight: 500;")
        opt_layout.addWidget(self.chk_auto_supersede)
        opt_layout.addStretch()
        layout.addLayout(opt_layout)

        # Action Buttons
        btn_box = QHBoxLayout()
        btn_box.setSpacing(8)
        btn_box.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 6px 14px;
                font-size: 11px;
                font-weight: 500;
                color: {C.TEXT_BODY};
            }}
            QPushButton:hover {{
                background-color: {C.BG_SURFACE};
            }}
        """)
        self.btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cancel)

        self.btn_ingest = QPushButton("📥 Ingest Selected Releases (0)")
        self.btn_ingest.setEnabled(False)
        self.btn_ingest.setCursor(Qt.PointingHandCursor)
        self.btn_ingest.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.ACCENT_SKY};
                color: {C.BG_CARD};
                border: 1px solid {C.ACCENT_SKY};
                border-radius: 4px;
                padding: 6px 16px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C.ACCENT_SKY_DARK};
            }}
            QPushButton:disabled {{
                background-color: {C.BORDER_STRONG};
                color: {C.TEXT_FAINT};
                border-color: {C.BORDER_STRONG};
            }}
        """)
        self.btn_ingest.clicked.connect(self._on_ingest_confirmed)
        btn_box.addWidget(self.btn_ingest)

        layout.addLayout(btn_box)
