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
    QMenu, QLabel, QCheckBox, QDateEdit, QSplitter,
    QMessageBox, QFrame, QFileDialog, QDialog, QStyledItemDelegate, QStyle
)

from core.network.session import NetworkConfigDialog
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_WARNING
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



# Delegates extracted to meetings_delegates.py (re-exported for backward compatibility)
from modules.meetings.ui.meetings_delegates import TDocsButtonDelegate, DynamicHoverMenuDelegate
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
            "date_to": self.date_to.date().toString("yyyy-MM-dd")
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

    def _browse_cache_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Cache Directory", self.dl_dir_input.text())
        if directory:
            normalized_dir = str(Path(directory))
            self.dl_dir_input.setText(normalized_dir)
            self.settings.save_settings(normalized_dir)


    def _setup_ui(self):
        main_layout = QHBoxLayout(self)
        self.splitter = QSplitter(Qt.Horizontal)

        # --- Settings Dialog & Persistent Reference Bridge ---
        self.config_dialog = MeetingsConfigDialog(self.settings, self)
        self.dl_dir_input = self.config_dialog.dl_dir_input
        self.chk_wg = self.config_dialog.chk_wg
        self.chk_dyna = self.config_dialog.chk_dyna
        self.chk_docs = self.config_dialog.chk_docs

        # --- Left Side: Table View ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        self.table = QTableView()
        self.table.setMouseTracking(True)
        self.table_model = MeetingsTableModel()
        self.table.setModel(self.table_model)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        self.table.setStyleSheet(
            "QTableView { border: 1px solid #CBD5E1; gridline-color: #F1F5F9; background-color: #FFFFFF; } "
            "QTableView::item:selected { background-color: #EBF3FC; color: #1E293B; }"
        )

        self.hover_delegate = DynamicHoverMenuDelegate(self)
        self.table.setItemDelegateForColumn(0, self.hover_delegate)

        self.tdocs_delegate = TDocsButtonDelegate(self)
        self.table.setItemDelegateForColumn(1, self.tdocs_delegate)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.resizeSection(0, 40)
        header.resizeSection(1, 90)
        header.resizeSection(2, 60)
        header.resizeSection(3, 90)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.resizeSection(5, 90)
        header.resizeSection(6, 90)
        header.resizeSection(7, 110)
        header.resizeSection(8, 110)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_right_click_menu)
        left_layout.addWidget(self.table)

        # --- Comparison Cart ---
        self.cart_frame = QFrame()
        self.cart_frame.setStyleSheet("""
                QFrame {
                    background-color: #F8FAFC;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                }
                QLabel { color: #1E293B; border: none; }
            """)
        cart_layout = QHBoxLayout(self.cart_frame)
        cart_layout.setContentsMargins(12, 8, 12, 8)

        cart_layout.addWidget(QLabel("<b>⚖️ Comparison Cart:</b>"))
        self.lbl_slot_a = QLabel("<i style='color:#64748B;'>[Slot A Empty]</i>")
        self.lbl_slot_b = QLabel("<i style='color:#64748B;'>[Slot B Empty]</i>")

        cart_layout.addSpacing(8)
        cart_layout.addWidget(self.lbl_slot_a)
        cart_layout.addWidget(QLabel(" <b>VS</b> "))
        cart_layout.addWidget(self.lbl_slot_b)
        cart_layout.addStretch()

        self.btn_compare = QPushButton("⚖️ Compare in Word")
        self.btn_compare.setObjectName("primaryBtn")
        self.btn_compare.setEnabled(False)
        self.btn_compare.setToolTip("Launch an invisible Word instance to generate a visual redline diff.")
        self.btn_compare.clicked.connect(self._run_comparison)

        self.btn_clear_cart = QPushButton("✖ Clear")
        self.btn_clear_cart.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear_cart.setToolTip("Clear all items from the comparison cart.")
        self.btn_clear_cart.clicked.connect(ComparisonManager.get_instance().clear_cart)

        cart_layout.addWidget(self.btn_compare)
        cart_layout.addWidget(self.btn_clear_cart)
        left_layout.addWidget(self.cart_frame)
        self.splitter.addWidget(left_widget)

        ComparisonManager.get_instance().cart_updated.connect(self._update_cart_ui)

        # --- Right Side: Refactored Filter & Action Panel ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(6, 4, 6, 6)

        # 1. Quick Launch
        self.btn_open_last = QPushButton("🚀 Open Last Meeting")
        self.btn_open_last.setObjectName("primaryBtn")
        self.btn_open_last.setFixedHeight(34)
        self.btn_open_last.clicked.connect(self._open_last_meeting)
        right_layout.addWidget(self.btn_open_last)

        # 2. Smart Global TDoc Jump Card
        jump_frame = QFrame()
        jump_frame.setStyleSheet("""
                QFrame {
                    background-color: #F8FAFC;
                    border: 1px solid #E2E8F0;
                    border-radius: 6px;
                }
            """)
        jump_layout = QVBoxLayout(jump_frame)
        jump_layout.setContentsMargins(8, 6, 8, 6)
        jump_layout.setSpacing(4)

        jump_header = QLabel("⚡ <b>Quick TDoc Jump</b>")
        jump_header.setStyleSheet("font-size: 11px; color: #475569; border: none;")
        jump_layout.addWidget(jump_header)

        global_search_layout = QHBoxLayout()
        global_search_layout.setSpacing(4)

        # Quick TDoc Jump Card
        self.global_tdoc_input = QLineEdit()
        self.global_tdoc_input.setPlaceholderText("e.g., S2-2605740")
        self.global_tdoc_input.setToolTip(
            "Type a valid TDoc number. Press Enter to open the document, or use the quick buttons.")

        self.btn_open_tdoc = QPushButton("📄 Doc")
        self.btn_open_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc.setFixedHeight(26)
        self.btn_open_tdoc.setObjectName("primaryBtn")
        self.btn_open_tdoc.setVisible(False)
        self.btn_open_tdoc.setToolTip("Download and open the specified TDoc directly in Microsoft Word.")

        self.btn_open_meeting = QPushButton("🗓️ Mtg")
        self.btn_open_meeting.setCursor(Qt.PointingHandCursor)
        self.btn_open_meeting.setFixedHeight(26)
        self.btn_open_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_meeting.setVisible(False)
        self.btn_open_meeting.setToolTip("Open the full meeting table containing this TDoc.")

        global_search_layout.addWidget(self.global_tdoc_input, 1)
        global_search_layout.addWidget(self.btn_open_tdoc)
        global_search_layout.addWidget(self.btn_open_meeting)
        jump_layout.addLayout(global_search_layout)

        # Shown only when strict Quick TDoc Jump cannot resolve the document
        # and same-year meetings have incomplete Docs-derived TDoc metadata.
        self.tdoc_metadata_warning = QLabel()
        self.tdoc_metadata_warning.setWordWrap(True)
        self.tdoc_metadata_warning.setStyleSheet(
            "color: #B45309; font-size: 10px; border: none; background: transparent;"
        )
        self.tdoc_metadata_warning.setVisible(False)
        jump_layout.addWidget(self.tdoc_metadata_warning)

        self.btn_update_tdoc_metadata = QPushButton("🔄 Update TDoc metadata")
        self.btn_update_tdoc_metadata.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.btn_update_tdoc_metadata.setFixedHeight(26)
        self.btn_update_tdoc_metadata.setVisible(False)
        self.btn_update_tdoc_metadata.setToolTip(
            "Deep-scrape the Docs/ folders of the affected meetings to rebuild their first/last TDoc ranges."
        )
        self.btn_update_tdoc_metadata.clicked.connect(self._update_incomplete_tdoc_metadata)
        jump_layout.addWidget(self.btn_update_tdoc_metadata)

        right_layout.addWidget(jump_frame)

        # 3. Filter & Search Controls
        filter_title = QLabel("🔍 <b>Filters</b>")
        filter_title.setStyleSheet("font-size: 12px; color: #1E293B; margin-top: 2px;")
        right_layout.addWidget(filter_title)

        # Working Group Selector
        self.wg_filter = CheckableComboBox("Working Group")
        self.wg_filter.setToolTip("Filter meetings by Working Group (e.g., SA2, RAN1).")
        self.wg_filter.selectionChanged.connect(self.refresh_table)
        right_layout.addWidget(self.wg_filter)

        # Compact Side-by-Side Combos (Ad-Hoc & Meeting Type)
        type_row = QHBoxLayout()
        type_row.setSpacing(6)

        self.adhoc_filter = QComboBox()
        self.adhoc_filter.addItems(["All Meetings", "Regular", "Ad-Hoc / BIS"])
        self.adhoc_filter.setToolTip("Filter by regular vs. ad-hoc meetings.")
        self.adhoc_filter.currentTextChanged.connect(self.refresh_table)

        self.type_filter = QComboBox()
        self.type_filter.addItems(["All Types", "In-Person", "Electronic"])
        self.type_filter.setToolTip("Filter by in-person vs. electronic meetings.")
        self.type_filter.currentTextChanged.connect(self.refresh_table)

        type_row.addWidget(self.adhoc_filter, 1)
        type_row.addWidget(self.type_filter, 1)
        right_layout.addLayout(type_row)

        # Search Box
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter number, city, or name...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setToolTip("Filter visible meetings by meeting number, host city, or title.")
        self.search_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.search_input)

        # Compact Date Range Selection
        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.setToolTip("Filter meetings within specific dates.")
        self.enable_dates_cb.toggled.connect(self._toggle_date_inputs)
        self.enable_dates_cb.toggled.connect(self.refresh_table)
        right_layout.addWidget(self.enable_dates_cb)

        dates_row = QHBoxLayout()
        dates_row.setSpacing(4)

        # Date Pickers
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.setToolTip("Filter meetings starting on or after this date.")
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)

        arrow_lbl = QLabel("➔")
        arrow_lbl.setAlignment(Qt.AlignCenter)
        arrow_lbl.setStyleSheet("color: #94A3B8; font-weight: bold;")

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.setToolTip("Filter meetings ending on or before this date.")
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)

        dates_row.addWidget(self.date_from, 1)
        dates_row.addWidget(arrow_lbl)
        dates_row.addWidget(self.date_to, 1)
        right_layout.addLayout(dates_row)
        self.enable_dates_cb.setChecked(True)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        sep1.setFrameShadow(QFrame.Sunken)
        right_layout.addWidget(sep1)

        # 4. Synchronization & Data Management
        sync_lbl = QLabel("<b>Sync & Database</b>")
        sync_lbl.setStyleSheet("font-size: 11px; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px;")
        right_layout.addWidget(sync_lbl)

        sync_btn_row = QHBoxLayout()
        sync_btn_row.setSpacing(4)

        # Sync & Database Buttons
        self.update_btn = QPushButton("🔄 Sync All Meetings")
        self.update_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.update_btn.setToolTip(
            "Run multi-phase scrape across 3GPP FTP servers according to configured scraper options.")
        self.update_btn.clicked.connect(lambda: self.update_db_requested.emit(
            self.chk_wg.isChecked(), self.chk_docs.isChecked(), self.chk_dyna.isChecked()
        ))
        sync_btn_row.addWidget(self.update_btn, 1)

        self.btn_config = QPushButton("⚙️")
        self.btn_config.setFixedWidth(32)
        self.btn_config.setToolTip("Configure local cache directory and scraping options...")
        self.btn_config.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_config.clicked.connect(self.config_dialog.exec_)
        sync_btn_row.addWidget(self.btn_config)
        right_layout.addLayout(sync_btn_row)

        self.btn_add_meeting = QPushButton("➕ Add / Fetch Meeting...")
        self.btn_add_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_add_meeting.setToolTip(
            "Manually query, preview, and import a specific meeting via Portal ID, number, or FTP URL.")
        self.btn_add_meeting.clicked.connect(self._open_add_meeting_dialog)
        right_layout.addWidget(self.btn_add_meeting)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setFrameShadow(QFrame.Sunken)
        right_layout.addWidget(sep2)

        # 5. Reports & Analytics
        reports_lbl = QLabel("<b>Reports & Exports</b>")
        reports_lbl.setStyleSheet("font-size: 11px; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px;")
        right_layout.addWidget(reports_lbl)

        self.btn_contribution_report = QPushButton("📊 Contributions Report...")
        self.btn_contribution_report.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_contribution_report.setToolTip("Filter and audit company contributions across multiple meetings.")
        self.btn_contribution_report.clicked.connect(self._open_contribution_report)
        right_layout.addWidget(self.btn_contribution_report)

        # Export & Destructive Actions
        self.btn_export_merged = QPushButton("📥 Export Merged TDocs (Excel)")
        self.btn_export_merged.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_merged.setToolTip(
            "Download and consolidate TDoc lists from all currently filtered meetings into a single Excel workbook.")
        self.btn_export_merged.clicked.connect(self._export_merged_tdocs)
        right_layout.addWidget(self.btn_export_merged)

        right_layout.addStretch()

        # 6. Destructive Actions (De-emphasized at the bottom)
        self.delete_all_btn = QPushButton("🗑️ Clear All Meetings...")
        self.delete_all_btn.setToolTip("Permanently erase all meeting records and metadata from the local database.")
        self.delete_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.delete_all_btn.clicked.connect(self._confirm_delete_all)
        right_layout.addWidget(self.delete_all_btn)

        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([750, 250])
        main_layout.addWidget(self.splitter)
        self._populate_filters()

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
            self.btn_update_tdoc_metadata.setText("🔄 Update TDoc metadata")
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
        self.btn_update_tdoc_metadata.setText("🔄 Update TDoc metadata")

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

        # update_specific_requested(targets, sync_wg, sync_docs, sync_dyna)
        #
        # Do not leave this local button disabled here: MeetingsTab emits the
        # request, but synchronization is owned by the outer meetings plugin
        # and this controller has no completion signal to re-enable it safely.
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
        self.lbl_slot_a.setText(
            f"<b style='color:#1E5C99;'>{slot_a['name']}</b>" if slot_a else "<i style='color:#64748B;'>[Slot A Empty]</i>")
        self.lbl_slot_b.setText(
            f"<b style='color:#1E5C99;'>{slot_b['name']}</b>" if slot_b else "<i style='color:#64748B;'>[Slot B Empty]</i>")
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
        """
        Refresh the available WG choices without discarding the user's current
        selection. Database synchronization may discover new WGs, but a UI
        refresh must not silently reset existing filters.
        """
        selected_wgs = set(self.wg_filter.getCheckedItems())
        wgs = self.db.get_working_groups()

        self.wg_filter.blockSignals(True)
        try:
            self.wg_filter.updateItems(wgs)

            # CheckableComboBox reserves row 0 for its "all/summary" entry.
            # Restore only selections that still exist in the refreshed list.
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

        for row in data:
            mtg_id = row.get("mtg_id")
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
            else:
                row['tdoc_btn_status'] = 'get'

        self.table_model.update_data(data)

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
            self.btn_open_last.setText(f"🚀 Open Last Meeting ({display_tag})")
            self.btn_open_last.setToolTip(f"Instantly resume working on meeting {display_tag}")
        else:
            self.btn_open_last.setText("🚀 Open Last Meeting")
            self.btn_open_last.setToolTip("No recent meeting history found. Please open a meeting first.")

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
