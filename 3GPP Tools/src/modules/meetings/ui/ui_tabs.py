# --- File: modules/meetings/ui/ui_tabs.py ---
import logging
import os
import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QEvent, Qt, pyqtSignal, QDate, QPoint, QTimer
from PyQt5.QtGui import QPen, QColor, QBrush
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLineEdit, QComboBox, QTableView, QHeaderView,
    QMenu, QLabel, QCheckBox, QDateEdit,
    QMessageBox, QFrame, QFileDialog, QDialog, QStyledItemDelegate, QStyle
)

from core.network.session import NetworkConfigDialog
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_WARNING,
    SEGMENT_BTN_STYLE,
    SPEC_FILTER_COMBO_STYLE,
    SPEC_FILTER_INPUT_STYLE,
)
from modules.meetings.core.agenda_manager import AgendaDownloaderThread
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.ui.dialogs import MeetingInfoDialog, AddMeetingDialog, MeetingsConfigDialog
from modules.meetings.ui.models import MeetingsTableModel
from modules.meetings.ui.search_controller import GlobalSearchController
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_window import TDocsWindow, _open_terminal
from modules.word_tools.core.word_comparator import WordComparatorThread
from modules.meetings.core.tdocs_merger import TDocsMergerThread
from modules.meetings.ui.contribution_dialog import ContributionReportDialog

# Delegates
from modules.meetings.ui.meetings_delegates import TDocsButtonDelegate, DynamicHoverMenuDelegate, WGTagDelegate
from modules.meetings.ui.meetings_menu import populate_meeting_context_menu
from modules.meetings.ui.meetings_actions import (
    open_last_meeting,
    start_tdocs_caching,
    on_tdocs_caching_finished,
    export_merged_tdocs,
    on_merger_finished,
    open_contribution_report,
    open_meeting_from_report,
)


class MeetingsTab(QWidget):
    update_db_requested = pyqtSignal(bool, bool, bool)
    update_specific_requested = pyqtSignal(list, bool, bool, bool)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.settings = MeetingsSettings()
        self.search_controller = GlobalSearchController(self)

        self.tdoc_windows = {}
        self._active_dl_threads = {}
        self._current_tsg = ""

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        self.search_controller.connect_signals()

        self._load_filters()
        self._update_last_meeting_btn()
        self.refresh_table()

    def _save_filters(self):
        filters = {
            "wg": self.wg_filter.getCheckedItems(),
            "adhoc": self.adhoc_filter.currentText(),
            "type": self.type_filter.currentText(),
            "search": self.search_input.text().strip(),
            "enable_dates": self.enable_dates_cb.isChecked(),
            "date_from": self.date_from.date().toString("yyyy-MM-dd"),
            "date_to": self.date_to.date().toString("yyyy-MM-dd"),
            "date_preset": self.date_preset_combo.currentText(),
            "tsg": getattr(self, "_current_tsg", ""),
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        filters = self.settings.get_filters()
        if not filters:
            return

        self.wg_filter.blockSignals(True)
        self.adhoc_filter.blockSignals(True)
        self.type_filter.blockSignals(True)
        self.search_input.blockSignals(True)
        self.enable_dates_cb.blockSignals(True)
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)
        self.date_preset_combo.blockSignals(True)

        if "tsg" in filters:
            self._current_tsg = filters["tsg"]
            self._update_tsg_button_styles()

        if "wg" in filters:
            saved_wgs = filters["wg"]
            if isinstance(saved_wgs, list):
                for i in range(1, self.wg_filter.model().rowCount()):
                    item = self.wg_filter.model().item(i)
                    if item.text() in saved_wgs:
                        item.setCheckState(Qt.Checked)
                    else:
                        item.setCheckState(Qt.Unchecked)
                self.wg_filter.updateText()

        if "adhoc" in filters:
            idx = self.adhoc_filter.findText(filters["adhoc"])
            if idx >= 0:
                self.adhoc_filter.setCurrentIndex(idx)

        if "type" in filters:
            idx = self.type_filter.findText(filters["type"])
            if idx >= 0:
                self.type_filter.setCurrentIndex(idx)

        if "search" in filters:
            self.search_input.setText(filters["search"])

        if "date_preset" in filters:
            preset_name = filters["date_preset"]
            idx = self.date_preset_combo.findText(preset_name)
            if idx >= 0:
                self.date_preset_combo.setCurrentIndex(idx)
                is_custom = (preset_name == "Custom Range...")
                self.date_from.setVisible(is_custom)
                self.arrow_lbl.setVisible(is_custom)
                self.date_to.setVisible(is_custom)

        if "enable_dates" in filters:
            self.enable_dates_cb.setChecked(filters["enable_dates"])
            self._toggle_date_inputs(filters["enable_dates"])

        if "date_from" in filters:
            d = QDate.fromString(filters["date_from"], "yyyy-MM-dd")
            if d.isValid():
                self.date_from.setDate(d)

        if "date_to" in filters:
            d = QDate.fromString(filters["date_to"], "yyyy-MM-dd")
            if d.isValid():
                self.date_to.setDate(d)

        self.wg_filter.blockSignals(False)
        self.adhoc_filter.blockSignals(False)
        self.type_filter.blockSignals(False)
        self.search_input.blockSignals(False)
        self.enable_dates_cb.blockSignals(False)
        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)
        self.date_preset_combo.blockSignals(False)

    def _browse_cache_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Cache Directory", self.dl_dir_input.text())
        if directory:
            normalized_dir = str(Path(directory))
            self.dl_dir_input.setText(normalized_dir)
            self.settings.save_settings(normalized_dir)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # --- Settings Dialog & Persistent Reference Bridge ---
        self.config_dialog = MeetingsConfigDialog(self.settings, self)
        self.dl_dir_input = self.config_dialog.dl_dir_input
        self.chk_wg = self.config_dialog.chk_wg
        self.chk_dyna = self.config_dialog.chk_dyna
        self.chk_docs = self.config_dialog.chk_docs
        self.delete_all_btn = getattr(self.config_dialog, "btn_clear", None)

        # ==========================================
        # --- ROW 1: Quick Actions & Primary Toolbar ---
        # ==========================================
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(8)

        # 1. Quick Launch Last Meeting
        self.btn_open_last = QPushButton("🚀 Open Last Meeting")
        self.btn_open_last.setObjectName("primaryBtn")
        self.btn_open_last.setFixedHeight(30)
        self.btn_open_last.setCursor(Qt.PointingHandCursor)
        self.btn_open_last.clicked.connect(self._open_last_meeting)
        row1_layout.addWidget(self.btn_open_last)

        # Vertical Divider
        sep_v1 = QFrame()
        sep_v1.setFrameShape(QFrame.VLine)
        sep_v1.setFrameShadow(QFrame.Sunken)
        row1_layout.addWidget(sep_v1)

        # 2. Quick TDoc Jump
        tdoc_lbl = QLabel("⚡ <b>TDoc:</b>")
        tdoc_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        row1_layout.addWidget(tdoc_lbl)

        self.global_tdoc_input = QLineEdit()
        self.global_tdoc_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.global_tdoc_input.setPlaceholderText("Jump to TDoc (e.g., S2-2605740)...")
        self.global_tdoc_input.setFixedHeight(30)
        self.global_tdoc_input.setMinimumWidth(210)
        self.global_tdoc_input.setClearButtonEnabled(True)
        self.global_tdoc_input.setToolTip("Type a TDoc number. Press Enter to open or use the quick buttons.")
        row1_layout.addWidget(self.global_tdoc_input)

        self.btn_open_tdoc = QPushButton("📄 Doc")
        self.btn_open_tdoc.setObjectName("primaryBtn")
        self.btn_open_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc.setFixedHeight(28)
        self.btn_open_tdoc.setVisible(False)
        self.btn_open_tdoc.setToolTip("Download and open the specified TDoc directly in Microsoft Word.")
        row1_layout.addWidget(self.btn_open_tdoc)

        self.btn_open_meeting = QPushButton("🗓️ Mtg")
        self.btn_open_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_meeting.setCursor(Qt.PointingHandCursor)
        self.btn_open_meeting.setFixedHeight(28)
        self.btn_open_meeting.setVisible(False)
        self.btn_open_meeting.setToolTip("Open the full meeting table containing this TDoc.")
        row1_layout.addWidget(self.btn_open_meeting)

        self.tdoc_metadata_warning = QLabel()
        self.tdoc_metadata_warning.setStyleSheet("color: #B45309; font-size: 11px; font-weight: bold;")
        self.tdoc_metadata_warning.setVisible(False)
        row1_layout.addWidget(self.tdoc_metadata_warning)

        self.btn_update_tdoc_metadata = QPushButton("🔄 Update metadata")
        self.btn_update_tdoc_metadata.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.btn_update_tdoc_metadata.setFixedHeight(26)
        self.btn_update_tdoc_metadata.setVisible(False)
        self.btn_update_tdoc_metadata.setToolTip("Deep-scrape Docs/ folders to rebuild missing first/last TDoc ranges.")
        self.btn_update_tdoc_metadata.clicked.connect(self._update_incomplete_tdoc_metadata)
        row1_layout.addWidget(self.btn_update_tdoc_metadata)

        row1_layout.addStretch()

        # 3. Network Sync Menu
        self.sync_menu_btn = QPushButton("🔄 Sync ▾")
        self.sync_menu_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.sync_menu_btn.setFixedHeight(30)
        self.sync_menu_btn.setCursor(Qt.PointingHandCursor)
        self.sync_menu_btn.setToolTip("Synchronization and FTP scraping options...")

        sync_menu = QMenu(self)
        sync_act_all = sync_menu.addAction("🔄 Sync All Meetings", lambda: self.update_db_requested.emit(
            self.chk_wg.isChecked(), self.chk_docs.isChecked(), self.chk_dyna.isChecked()
        ))
        sync_act_all.setToolTip("Run multi-phase FTP scrape according to configured scraper options.")

        sync_act_add = sync_menu.addAction("➕ Add / Fetch Specific Meeting...", self._open_add_meeting_dialog)
        sync_act_add.setToolTip("Manually query, preview, and import a specific meeting via Portal ID or FTP URL.")

        sync_menu.addSeparator()
        sync_act_cfg = sync_menu.addAction("⚙️ Scraper && Cache Preferences...", self.config_dialog.exec_)
        sync_act_cfg.setToolTip("Configure local cache directory and FTP scraper phase toggles.")

        self.sync_menu_btn.setMenu(sync_menu)
        row1_layout.addWidget(self.sync_menu_btn)
        self.update_btn = self.sync_menu_btn  # Compatibility alias
        self.btn_add_meeting = self.sync_menu_btn  # Compatibility alias

        # 4. Reports & Analytics Menu
        self.reports_menu_btn = QPushButton("📊 Reports ▾")
        self.reports_menu_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.reports_menu_btn.setFixedHeight(30)
        self.reports_menu_btn.setCursor(Qt.PointingHandCursor)
        self.reports_menu_btn.setToolTip("Reports, analytics, and consolidated exports...")

        reports_menu = QMenu(self)
        rep_act_contrib = reports_menu.addAction("📊 Contributions Report...", self._open_contribution_report)
        rep_act_contrib.setToolTip("Filter and audit company contributions across multiple meetings.")

        rep_act_merge = reports_menu.addAction("📥 Export Merged TDocs (Excel)...", self._export_merged_tdocs)
        rep_act_merge.setToolTip("Download and consolidate TDoc lists from filtered meetings into a single Excel workbook.")

        self.reports_menu_btn.setMenu(reports_menu)
        row1_layout.addWidget(self.reports_menu_btn)
        self.btn_contribution_report = self.reports_menu_btn  # Compatibility alias
        self.btn_export_merged = self.reports_menu_btn  # Compatibility alias

        # 5. Settings Button
        self.btn_config = QPushButton("⚙️ Settings")
        self.btn_config.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_config.setFixedHeight(30)
        self.btn_config.setCursor(Qt.PointingHandCursor)
        self.btn_config.setToolTip("Open Meetings & Cache Configuration...")
        self.btn_config.clicked.connect(self.config_dialog.exec_)
        row1_layout.addWidget(self.btn_config)

        main_layout.addLayout(row1_layout)

        # ==========================================
        # --- ROW 2: Search & Filter Engine ---
        # ==========================================
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(6)

        search_lbl = QLabel("<b>🔍 Search:</b>")
        search_lbl.setStyleSheet("color: #334155; font-size: 11px;")
        row2_layout.addWidget(search_lbl)

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.search_input.setPlaceholderText("Filter number, city, or title...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFixedHeight(28)
        self.search_input.setMinimumWidth(180)
        self.search_input.setToolTip("Filter visible meetings by meeting number, host city, or title.")
        self.search_input.textChanged.connect(self.refresh_table)
        row2_layout.addWidget(self.search_input, stretch=2)

        # TSG Segmented Chips (All, SA, RAN, CT)
        tsg_lbl = QLabel("TSG:")
        tsg_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 2px;")
        row2_layout.addWidget(tsg_lbl)

        self.tsg_buttons = {}
        for t_label, t_val in [("All", ""), ("SA", "SA"), ("RAN", "RAN"), ("CT", "CT")]:
            btn = QPushButton(t_label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setFixedHeight(26)
            btn.setToolTip(f"Filter meetings under TSG {t_label}" if t_val else "Show meetings across all TSGs")
            btn.clicked.connect(lambda _, tv=t_val: self._on_tsg_button_clicked(tv))
            self.tsg_buttons[t_val] = btn
            row2_layout.addWidget(btn)
        self._update_tsg_button_styles()

        # Working Group Selector
        wg_lbl = QLabel("WG:")
        wg_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 2px;")
        row2_layout.addWidget(wg_lbl)

        self.wg_filter = CheckableComboBox("Working Group")
        self.wg_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_filter.setMinimumWidth(125)
        self.wg_filter.setFixedHeight(28)
        self.wg_filter.setToolTip("Filter meetings by Working Group (e.g., SA2, RAN1).")
        self.wg_filter.selectionChanged.connect(self.refresh_table)
        row2_layout.addWidget(self.wg_filter)

        # Ad-Hoc / Nature Dropdown
        self.adhoc_filter = QComboBox()
        self.adhoc_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.adhoc_filter.addItems(["All Meetings", "Regular", "Ad-Hoc / BIS"])
        self.adhoc_filter.setFixedHeight(28)
        self.adhoc_filter.setToolTip("Filter by regular vs. ad-hoc meetings.")
        self.adhoc_filter.currentTextChanged.connect(self.refresh_table)
        row2_layout.addWidget(self.adhoc_filter)

        # In-Person / Electronic Dropdown
        self.type_filter = QComboBox()
        self.type_filter.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.type_filter.addItems(["All Formats", "In-Person", "Electronic"])
        self.type_filter.setFixedHeight(28)
        self.type_filter.setToolTip("Filter by in-person vs. electronic meetings.")
        self.type_filter.currentTextChanged.connect(self.refresh_table)
        row2_layout.addWidget(self.type_filter)

        # Dates Filter: Preset Combo + Date Range Pickers
        dates_lbl = QLabel("Dates:")
        dates_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 2px;")
        row2_layout.addWidget(dates_lbl)

        self.date_preset_combo = QComboBox()
        self.date_preset_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.date_preset_combo.addItems(["All Dates", "Upcoming", "This Year", "Past 12 Months", "Custom Range..."])
        self.date_preset_combo.setFixedHeight(28)
        self.date_preset_combo.setToolTip("Quickly scope meetings by date presets or select a custom range.")
        self.date_preset_combo.currentIndexChanged.connect(self._on_date_preset_changed)
        row2_layout.addWidget(self.date_preset_combo)

        # Hidden checkbox used internally for backward compatibility
        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.setVisible(False)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setFixedHeight(28)
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.setToolTip("Filter meetings starting on or after this date.")
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)
        self.date_from.setVisible(False)
        row2_layout.addWidget(self.date_from)

        self.arrow_lbl = QLabel("➔")
        self.arrow_lbl.setAlignment(Qt.AlignCenter)
        self.arrow_lbl.setStyleSheet("color: #94A3B8; font-weight: bold;")
        self.arrow_lbl.setVisible(False)
        row2_layout.addWidget(self.arrow_lbl)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setFixedHeight(28)
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.setToolTip("Filter meetings ending on or before this date.")
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)
        self.date_to.setVisible(False)
        row2_layout.addWidget(self.date_to)

        # Reset Filters Button
        self.btn_reset_filters = QPushButton("✕ Clear")
        self.btn_reset_filters.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_reset_filters.setFixedHeight(28)
        self.btn_reset_filters.setCursor(Qt.PointingHandCursor)
        self.btn_reset_filters.setToolTip("Reset all filters to defaults.")
        self.btn_reset_filters.clicked.connect(self._reset_filters)
        row2_layout.addWidget(self.btn_reset_filters)

        main_layout.addLayout(row2_layout)

        # ==========================================
        # --- ROW 3: Full-Width Meetings Table ---
        # ==========================================
        self.table = QTableView()
        self.table.setMouseTracking(True)
        self.table_model = MeetingsTableModel()
        self.table.setModel(self.table_model)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        self.table.setStyleSheet("""
            QTableView {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                gridline-color: #F1F5F9;
                background-color: #FFFFFF;
                alternate-background-color: #F8FAFC;
                selection-background-color: #EBF3FC;
                selection-color: #1E293B;
            }
            QTableView::item {
                padding: 4px 6px;
                border-bottom: 1px solid #F1F5F9;
            }
            QTableView::item:selected {
                background-color: #EBF3FC;
                color: #1E293B;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: bold;
                font-size: 11px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
                padding: 6px 8px;
            }
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

        header.resizeSection(0, 38)
        header.resizeSection(1, 90)
        header.resizeSection(2, 75)
        header.resizeSection(3, 95)
        # col 4 Stretch (Location)
        header.resizeSection(5, 145)
        header.resizeSection(6, 185)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_right_click_menu)
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        main_layout.addWidget(self.table)

        # ==========================================
        # --- ROW 4: Status Strip & Floating Cart ---
        # ==========================================
        bottom_layout = QHBoxLayout()
        bottom_layout.setContentsMargins(4, 2, 4, 2)
        bottom_layout.setSpacing(10)

        self.status_lbl = QLabel("Showing 0 meetings")
        self.status_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        bottom_layout.addWidget(self.status_lbl)

        bottom_layout.addStretch()

        # Floating Comparison Cart Pill (hidden when empty)
        self.cart_frame = QFrame()
        self.cart_frame.setObjectName("cartFrame")
        self.cart_frame.setStyleSheet("""
            QFrame#cartFrame {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 2px 8px;
            }
            QLabel { color: #1E293B; border: none; font-size: 11px; }
        """)
        cart_layout = QHBoxLayout(self.cart_frame)
        cart_layout.setContentsMargins(8, 2, 8, 2)
        cart_layout.setSpacing(6)

        cart_layout.addWidget(QLabel("<b>⚖️ Cart:</b>"))
        self.lbl_slot_a = QLabel("<i style='color:#64748B;'>[Slot A Empty]</i>")
        self.lbl_slot_b = QLabel("<i style='color:#64748B;'>[Slot B Empty]</i>")
        cart_layout.addWidget(self.lbl_slot_a)
        cart_layout.addWidget(QLabel("<b>vs</b>"))
        cart_layout.addWidget(self.lbl_slot_b)

        self.btn_compare = QPushButton("⚖️ Compare in Word")
        self.btn_compare.setObjectName("primaryBtn")
        self.btn_compare.setEnabled(False)
        self.btn_compare.setFixedHeight(24)
        self.btn_compare.setToolTip("Launch an invisible Word instance to generate a visual redline diff.")
        self.btn_compare.clicked.connect(self._run_comparison)
        cart_layout.addWidget(self.btn_compare)

        self.btn_clear_cart = QPushButton("✖ Clear")
        self.btn_clear_cart.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear_cart.setFixedHeight(24)
        self.btn_clear_cart.setToolTip("Clear all items from the comparison cart.")
        self.btn_clear_cart.clicked.connect(ComparisonManager.get_instance().clear_cart)
        cart_layout.addWidget(self.btn_clear_cart)

        self.cart_frame.setVisible(False)
        bottom_layout.addWidget(self.cart_frame)

        main_layout.addLayout(bottom_layout)

        ComparisonManager.get_instance().cart_updated.connect(self._update_cart_ui)
        self._populate_filters()

    # ==========================================
    # --- FILTER HANDLERS ---
    # ==========================================
    def _on_tsg_button_clicked(self, tsg_val: str):
        if getattr(self, "_current_tsg", "") == tsg_val and tsg_val != "":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg_val

        self._update_tsg_button_styles()

        model = self.wg_filter.model()
        self.wg_filter.blockSignals(True)
        try:
            for i in range(1, model.rowCount()):
                item = model.item(i)
                wg_name = item.text().upper()
                if not self._current_tsg:
                    item.setCheckState(Qt.Unchecked)
                elif self._current_tsg == "SA" and wg_name.startswith("SA"):
                    item.setCheckState(Qt.Checked)
                elif self._current_tsg == "RAN" and (wg_name.startswith("RAN") or wg_name.startswith("GERAN")):
                    item.setCheckState(Qt.Checked)
                elif self._current_tsg == "CT" and (wg_name.startswith("CT") or wg_name.startswith("CN")):
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

        self.refresh_table()

    def _update_tsg_button_styles(self):
        curr = getattr(self, "_current_tsg", "")
        for t_val, btn in getattr(self, "tsg_buttons", {}).items():
            btn.setChecked(curr == t_val)

    def _on_date_preset_changed(self):
        preset = self.date_preset_combo.currentText()
        today = QDate.currentDate()
        self.enable_dates_cb.blockSignals(True)
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)

        is_custom = (preset == "Custom Range...")
        self.date_from.setVisible(is_custom)
        self.arrow_lbl.setVisible(is_custom)
        self.date_to.setVisible(is_custom)

        try:
            if preset == "All Dates":
                self.enable_dates_cb.setChecked(False)
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Upcoming":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(today)
                self.date_to.setDate(today.addYears(3))
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "This Year":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(QDate(today.year(), 1, 1))
                self.date_to.setDate(QDate(today.year(), 12, 31))
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Past 12 Months":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(today.addYears(-1))
                self.date_to.setDate(today)
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Custom Range...":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setEnabled(True)
                self.date_to.setEnabled(True)
        finally:
            self.enable_dates_cb.blockSignals(False)
            self.date_from.blockSignals(False)
            self.date_to.blockSignals(False)

        self.refresh_table()

    def _reset_filters(self):
        self._current_tsg = ""
        self._update_tsg_button_styles()
        self.search_input.clear()
        self.adhoc_filter.setCurrentIndex(0)
        self.type_filter.setCurrentIndex(0)
        self.date_preset_combo.setCurrentIndex(0)
        self.date_from.setVisible(False)
        self.arrow_lbl.setVisible(False)
        self.date_to.setVisible(False)

        model = self.wg_filter.model()
        self.wg_filter.blockSignals(True)
        try:
            for i in range(1, model.rowCount()):
                model.item(i).setCheckState(Qt.Unchecked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

        self.refresh_table()

    def _on_table_double_clicked(self, index):
        if not index.isValid():
            return
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        status = row_data.get('tdoc_btn_status', 'na')
        if status == 'open':
            self._open_tdocs_window(row_data, row_data.get('tdoc_filepath'))
        elif status == 'get':
            self._download_and_open_tdocs(row_data, index.row())
        else:
            self.show_meeting_info(row_data)

    def refresh_quick_tdoc_lookup(self):
        """Re-evaluate the current Quick TDoc Jump after meeting metadata changes."""
        text = self.global_tdoc_input.text().strip()
        if text:
            self.search_controller.on_tdoc_input_changed(text)

    def _set_tdoc_metadata_warning(self, meetings: list):
        """Update the inline Quick TDoc Jump metadata-gap warning."""
        if not meetings:
            self.tdoc_metadata_warning.clear()
            self.tdoc_metadata_warning.setVisible(False)
            self.btn_update_tdoc_metadata.setVisible(False)
            self.btn_update_tdoc_metadata.setEnabled(True)
            self.btn_update_tdoc_metadata.setText("🔄 Update metadata")
            return

        refreshable = [m for m in meetings if (m.get("url_key") or "").strip()]
        unavailable = len(meetings) - len(refreshable)

        wg_name = meetings[0].get("wg_name", "")
        start_date = meetings[0].get("start_date", "")
        year = start_date[:4] if len(start_date) >= 4 else ""

        count = len(meetings)
        noun = "meeting" if count == 1 else "meetings"
        message = (
            f"⚠ No indexed TDoc range matched. {count} {wg_name} {noun}"
            f"{f' from {year}' if year else ''} have incomplete TDoc metadata."
        )
        if refreshable:
            message += " A Docs/ metadata refresh may make this TDoc discoverable."
        if unavailable:
            missing_noun = "meeting has" if unavailable == 1 else "meetings have"
            message += f" {unavailable} {missing_noun} no FTP path and cannot be refreshed automatically."

        self.tdoc_metadata_warning.setText(message)
        self.tdoc_metadata_warning.setVisible(True)
        self.btn_update_tdoc_metadata.setVisible(bool(refreshable))
        self.btn_update_tdoc_metadata.setEnabled(bool(refreshable))
        self.btn_update_tdoc_metadata.setText("🔄 Update metadata")

    def _update_incomplete_tdoc_metadata(self):
        """
        Run only scraper Phase 2 for same-WG/same-year meetings whose TDoc
        range is incomplete.
        """
        meetings = getattr(self.search_controller, "current_incomplete_meetings", [])
        refreshable = [m for m in meetings if (m.get("url_key") or "").strip()]
        if not refreshable:
            return

        targets = [
            {"wg": meeting.get("wg_name"), "meeting": meeting.get("meeting_number")}
            for meeting in refreshable
            if meeting.get("wg_name") and meeting.get("meeting_number")
        ]
        if not targets:
            return

        self.tdoc_metadata_warning.setText(
            f"⏳ Requested Docs/ metadata refresh for {len(targets)} "
            f"{'meeting' if len(targets) == 1 else 'meetings'}. "
            "Re-run the Quick TDoc Jump after synchronization completes."
        )
        self.update_specific_requested.emit(targets, False, True, False)

    def _open_add_meeting_dialog(self):
        dialog = AddMeetingDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            self.refresh_table()

    def _open_network_config(self):
        NetworkConfigDialog(self).exec_()

    def _confirm_delete_all(self):
        if QMessageBox.question(self, 'Confirm Clear', "Delete ALL meetings? Cannot be undone.",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_all_meetings()
            self.refresh_table()

    def _confirm_delete_specific(self, targets: list):
        if QMessageBox.question(self, 'Confirm', f"Delete {len(targets)} meeting(s)?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_specific_meetings(targets)
            self.refresh_table()

    def _emit_multi_delete(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets:
            self._confirm_delete_specific(targets)

    def _toggle_date_inputs(self, checked):
        self.date_from.setEnabled(checked)
        self.date_to.setEnabled(checked)

    def _update_cart_ui(self, slot_a: dict, slot_b: dict):
        has_items = bool(slot_a or slot_b)
        self.cart_frame.setVisible(has_items)

        if slot_a:
            self.lbl_slot_a.setText(f"<b style='color:#1E5C99;'>{slot_a.get('name', 'Doc A')}</b>")
        else:
            self.lbl_slot_a.setText("<i style='color:#64748B;'>[Slot A Empty]</i>")

        if slot_b:
            self.lbl_slot_b.setText(f"<b style='color:#1E5C99;'>{slot_b.get('name', 'Doc B')}</b>")
        else:
            self.lbl_slot_b.setText("<i style='color:#64748B;'>[Slot B Empty]</i>")

        self.btn_compare.setEnabled(bool(slot_a and slot_b))

    def _run_comparison(self):
        mgr = ComparisonManager.get_instance()
        if mgr.slot_a and mgr.slot_b:
            self.btn_compare.setText("⏳ Comparing...")
            self.btn_compare.setEnabled(False)

            self.cmp_thread = WordComparatorThread(mgr.slot_a['path'], mgr.slot_b['path'])
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setText("⚖️ Compare in Word"))
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setEnabled(True))
            self.cmp_thread.start()

    def _populate_filters(self):
        selected_wgs = set(self.wg_filter.getCheckedItems())
        wgs = self.db.get_working_groups()

        self.wg_filter.blockSignals(True)
        try:
            self.wg_filter.updateItems(wgs)
            model = self.wg_filter.model()
            for i in range(1, model.rowCount()):
                item = model.item(i)
                item.setCheckState(Qt.Checked if item.text() in selected_wgs else Qt.Unchecked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

    def refresh_table(self):
        self.save_filters_timer.start()

        selected_wgs = self.wg_filter.getCheckedItems()
        date_from = self.date_from.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None
        date_to = self.date_to.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None
        adhoc_val = self.adhoc_filter.currentText()
        type_val = self.type_filter.currentText()

        data = self.db.search_meetings(
            wg_name=selected_wgs, search_term=self.search_input.text().strip(),
            location=None, date_from=date_from, date_to=date_to,
            adhoc_filter=adhoc_val, type_filter=type_val
        )

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        cache_path = Path(current_cache)
        cached_folders = set(f.name for f in cache_path.iterdir() if f.is_dir()) if cache_path.exists() else set()

        cached_count = 0
        live_count = 0
        today = QDate.currentDate()

        for row in data:
            mtg_id = row.get("mtg_id")
            s = row.get("start_date", "")
            e = row.get("end_date", "")
            if s and e:
                try:
                    d1 = QDate.fromString(s, "yyyy-MM-dd")
                    d2 = QDate.fromString(e, "yyyy-MM-dd")
                    if d1.isValid() and d2.isValid() and d1 <= today <= d2:
                        live_count += 1
                except Exception:
                    pass

            if not mtg_id:
                row['tdoc_btn_status'] = 'na'
                continue

            folder_name = row.get("folder_name") or row.get("meeting_number", "")
            filepath = None

            if folder_name in cached_folders:
                agenda_dir = cache_path / folder_name / "Agenda"
                if agenda_dir.exists():
                    filepath = next((f for f in agenda_dir.iterdir() if (
                                "tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower()) and f.name.endswith(
                        ".xlsx")), None)
                    if not filepath:
                        fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                        filepath = fallback if fallback.exists() else None

            if filepath:
                row['tdoc_btn_status'] = 'open'
                row['tdoc_filepath'] = str(filepath)
                cached_count += 1
            else:
                row['tdoc_btn_status'] = 'get'

        self.table_model.update_data(data)

        # Update status bar summary
        total_count = len(data)
        summary_parts = [f"Showing <b>{total_count}</b> meeting{'s' if total_count != 1 else ''}"]
        if cached_count > 0:
            summary_parts.append(f"<b>{cached_count}</b> cached")
        if live_count > 0:
            summary_parts.append(f"<span style='color:#059669; font-weight:bold;'>🟢 {live_count} live</span>")

        self.status_lbl.setText(" · ".join(summary_parts))

    def _emit_multi_sync(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets:
            self.update_specific_requested.emit(targets, self.chk_wg.isChecked(), self.chk_docs.isChecked(),
                                                self.chk_dyna.isChecked())

    def show_right_click_menu(self, pos: QPoint):
        index = self.table.indexAt(pos)
        if index.isValid():
            row_data = self.table_model.data(index, Qt.UserRole)
            if not row_data:
                return
            menu = QMenu(self)
            self._populate_dynamic_menu(menu, row_data, index.row())
            menu.exec_(self.table.viewport().mapToGlobal(pos))

    def show_meeting_info(self, data: dict):
        MeetingInfoDialog(data, self).exec_()

    def _get_tdoc_list_path(self, row_data: dict) -> Path:
        mtg_id = row_data.get("mtg_id")
        if not mtg_id:
            return None

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        agenda_dir = Path(current_cache) / folder_name / "Agenda"

        if agenda_dir.exists() and agenda_dir.is_dir():
            for file_path in agenda_dir.iterdir():
                filename = file_path.name.lower()
                if (filename.startswith("tdoc_list_meeting_") or filename.startswith(
                        "tdocs_list_")) and filename.endswith(".xlsx"):
                    return file_path

        return agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"

    def _open_tdocs_window(self, mtg_info: dict, filepath: str):
        self.settings.save_last_meeting(mtg_info)
        self._update_last_meeting_btn()
        mtg_id = mtg_info.get("mtg_id")

        mtg_info["is_active_sync"] = self.db.is_active_sync_meeting(
            mtg_info.get("wg_name", ""),
            mtg_info.get("start_date", ""),
            mtg_info.get("end_date", ""),
            mtg_info.get("is_electronic", 0)
        )

        if mtg_id in self.tdoc_windows:
            existing_win = self.tdoc_windows[mtg_id]
            if existing_win and existing_win.isVisible():
                existing_win.raise_()
                existing_win.activateWindow()
                return
            else:
                existing_win.close()
                existing_win.deleteLater()
                del self.tdoc_windows[mtg_id]

        tdocs_data = TDocsParser.parse_tdocs_excel(filepath)
        if not tdocs_data:
            QMessageBox.warning(self, "Parse Error", "Could not read data from the Excel file.")
            return

        window = TDocsWindow(mtg_info, tdocs_data, filepath)
        window.global_action_requested.connect(self._handle_global_action_from_window)
        self.tdoc_windows[mtg_id] = window
        window.show()

    def closeEvent(self, event):
        if hasattr(self, 'contribution_dialog') and self.contribution_dialog:
            self.contribution_dialog.close()

        for win in list(self.tdoc_windows.values()):
            if win:
                win.close()
        self.tdoc_windows.clear()

        tab_threads = [
            getattr(self, 'agenda_dl_thread', None),
            getattr(self, 'cacher_thread', None),
            getattr(self, 'merger_thread', None),
            getattr(self, 'cmp_thread', None),
        ]
        if hasattr(self, '_active_dl_threads'):
            tab_threads.extend(self._active_dl_threads.values())

        for th in filter(None, tab_threads):
            if hasattr(th, 'is_cancelled'):
                th.is_cancelled = True
            if th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(150)

        super().closeEvent(event)

    def _on_inline_download_finished(self, success: bool, result: str, mtg_id: str, row_data: dict, row_idx: int):
        if mtg_id in self._active_dl_threads:
            del self._active_dl_threads[mtg_id]

        self.refresh_table()
        if success:
            self._open_tdocs_window(row_data, result)
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to download TDocs List:\n{result}")

    def _update_last_meeting_btn(self):
        last_id, last_num, last_wg = self.settings.get_last_meeting()

        if last_id and last_num and not last_wg:
            results = self.db.search_meetings(search_term=last_num)
            matched = next((m for m in results if m.get("mtg_id") == last_id), None)
            if matched:
                last_wg = matched.get("wg_name")

        if last_num:
            display_tag = f"{last_wg}#{last_num}" if last_wg else f"{last_num}"
            self.btn_open_last.setText(f"🚀 Last Meeting ({display_tag})")
            self.btn_open_last.setToolTip(f"Instantly resume working on meeting {display_tag}")
            self.btn_open_last.setEnabled(True)
        else:
            self.btn_open_last.setText("🚀 Open Last Meeting")
            self.btn_open_last.setToolTip("No recent meeting history found. Please open a meeting first.")
            self.btn_open_last.setEnabled(False)

    def _populate_dynamic_menu(self, menu: QMenu, row_data: dict, row_idx: int):
        populate_meeting_context_menu(self, menu, row_data, row_idx)

    def _download_agenda_csv(self, row_data: dict):
        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        agenda_dir = Path(current_cache) / folder_name / "Agenda"

        candidate_urls = []
        raw_url = row_data.get("url_key", "")
        if raw_url:
            full_ftp = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            candidate_urls.append(full_ftp)

        docs_url = row_data.get("docs_folder_url", "")
        if docs_url:
            candidate_urls.append(re.sub(r'/Docs/?$', '', docs_url, flags=re.IGNORECASE))

        self.agenda_dl_thread = AgendaDownloaderThread(candidate_urls, agenda_dir, self)
        self.agenda_dl_thread.finished.connect(self._on_agenda_download_finished)
        self.agenda_dl_thread.start()

    def _on_agenda_download_finished(self, success: bool, msg: str):
        if success:
            QMessageBox.information(self, "Agenda Download", msg)
        else:
            QMessageBox.warning(self, "Agenda Download", msg)

    def _download_and_open_tdocs(self, row_data: dict, row_idx: int = -1):
        mtg_id = row_data.get("mtg_id")
        row_data['tdoc_btn_status'] = 'fetching'

        if not isinstance(row_idx, int) or row_idx < 0:
            row_idx = -1
            if hasattr(self.table_model, '_data'):
                for i, r in enumerate(self.table_model._data):
                    if r.get('mtg_id') == mtg_id:
                        row_idx = i
                        break

        if isinstance(row_idx, int) and row_idx >= 0:
            index = self.table_model.index(row_idx, 1)
            self.table_model.dataChanged.emit(index, index)

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        local_path = Path(current_cache) / folder_name

        thread = TDocsDownloaderThread(mtg_id, local_path, self)
        self._active_dl_threads[mtg_id] = thread
        thread.finished.connect(
            lambda success, res, m_id: self._on_inline_download_finished(success, res, m_id, row_data, row_idx))
        thread.start()

    def _open_last_meeting(self):
        open_last_meeting(self)

    def _start_tdocs_caching(self, docs_url: str, local_path: Path):
        start_tdocs_caching(self, docs_url, local_path)

    def _on_tdocs_caching_finished(self, success: bool, msg: str):
        on_tdocs_caching_finished(self, success, msg)

    def _handle_global_action_from_window(self, tdoc_str: str, action: str):
        self.global_tdoc_input.setText(tdoc_str)
        self.search_controller.on_tdoc_input_changed(tdoc_str)

        if not self.search_controller.current_found_meeting:
            logging.warning(f"⚠️ MeetingsTab couldn't find meeting for '{tdoc_str}'")
            QMessageBox.warning(self, "Not Found",
                                f"Could not find '{tdoc_str}' in the global database. Ensure the meeting is synced.")
            return

        if action == 'open_meeting':
            self.search_controller.action_open_meeting_list()
        elif action == 'open_doc':
            self.search_controller.action_open_tdoc_only()
        elif action == 'add_to_cart':
            self.search_controller.action_add_to_cart()

    def _export_merged_tdocs(self):
        export_merged_tdocs(self)

    def _on_merger_finished(self, success: bool, msg: str):
        on_merger_finished(self, success, msg)

    def _open_contribution_report(self):
        open_contribution_report(self)

    def _open_meeting_from_report(self, wg_name: str, meeting_number: str):
        open_meeting_from_report(self, wg_name, meeting_number)

    def open_cr_pcr_for_meeting(self, row_data: dict):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog
        mtg_info = {
            "wg_name": row_data.get("wg_name", ""),
            "meeting_number": row_data.get("meeting_number", ""),
            "location": row_data.get("location", ""),
            "dates": row_data.get("dates", ""),
            "start_date": row_data.get("start_date", ""),
            "end_date": row_data.get("end_date", ""),
            "title": row_data.get("title", ""),
        }
        dlg = CrPcrDialog(
            spec_num="",
            meeting_info=mtg_info,
            parent=self,
        )
        dlg.show()
