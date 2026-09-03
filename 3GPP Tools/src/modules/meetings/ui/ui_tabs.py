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
from modules.meetings.ui.dialogs import MeetingInfoDialog, AddMeetingDialog
from modules.meetings.ui.models import MeetingsTableModel
from modules.meetings.ui.search_controller import GlobalSearchController
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_window import TDocsWindow
from modules.word_tools.core.word_comparator import WordComparatorThread
from modules.meetings.core.tdocs_merger import TDocsMergerThread


class TDocsButtonDelegate(QStyledItemDelegate):
    def __init__(self, parent_tab):
        super().__init__(parent_tab.table)
        self.parent_tab = parent_tab

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        row_data = index.model().data(index, Qt.UserRole)
        if not row_data:
            return

        status = row_data.get('tdoc_btn_status', 'na')
        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        if status == 'na':
            bg, border, text_col, text = "#F8FAFC", "#E2E8F0", "#94A3B8", "N/A"
        elif status == 'open':
            bg, border, text_col, text = "#ECFDF5", "#A7F3D0", "#065F46", "✔ Open"
        elif status == 'fetching':
            bg, border, text_col, text = "#FFFBEB", "#FDE68A", "#B45309", "⏳ Fetching"
        else:
            bg, border, text_col, text = "#EBF3FC", "#BFDBFE", "#1E5C99", "⬇ Get"

        rect = option.rect
        btn_rect = rect.adjusted(4, 6, -4, -6)

        painter.setBrush(QBrush(QColor(bg)))
        painter.setPen(QPen(QColor(border), 1))
        painter.drawRoundedRect(btn_rect, 6, 6)

        font = painter.font()
        font.setPointSize(8)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QPen(QColor(text_col)))
        painter.drawText(btn_rect, Qt.AlignCenter, text)

        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            row_data = model.data(index, Qt.UserRole)
            if not row_data:
                return False

            status = row_data.get('tdoc_btn_status', 'na')
            if status == 'open':
                self.parent_tab._open_tdocs_window(row_data, row_data.get('tdoc_filepath'))
            elif status == 'get':
                self.parent_tab._download_and_open_tdocs(row_data, index.row())
            return True

        return super().editorEvent(event, model, option, index)


class DynamicHoverMenuDelegate(QStyledItemDelegate):
    def __init__(self, parent_tab):
        super().__init__(parent_tab.table)
        self.parent_tab = parent_tab

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        painter.save()
        font = painter.font()
        font.setPointSize(16)
        font.setBold(True)
        painter.setFont(font)

        if option.state & QStyle.State_MouseOver:
            painter.setPen(QColor("#1E5C99"))
        else:
            painter.setPen(QColor("#94A3B8"))

        painter.drawText(option.rect, Qt.AlignCenter, "⋮")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            row_data = model.data(index, Qt.UserRole)
            if not row_data:
                return False

            menu = QMenu(self.parent_tab)
            self.parent_tab._populate_dynamic_menu(menu, row_data, index.row())
            menu.exec_(event.globalPos())
            return True

        return super().editorEvent(event, model, option, index)


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

        # --- Right Side: Filter & Sync Panel ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setAlignment(Qt.AlignTop)
        right_layout.setSpacing(6)

        self.btn_open_last = QPushButton("🚀 Open Last Meeting")
        self.btn_open_last.setObjectName("primaryBtn")
        self.btn_open_last.clicked.connect(self._open_last_meeting)
        right_layout.addWidget(self.btn_open_last)

        title_lbl = QLabel("<b>Filter & Search</b>")
        title_lbl.setStyleSheet("font-size: 13px; margin-top: 4px; color: #1E293B;")
        right_layout.addWidget(title_lbl)

        right_layout.addWidget(QLabel("Working Group:"))
        self.wg_filter = CheckableComboBox("Working Group")
        self.wg_filter.setToolTip("Filter meetings by Working Group (e.g., SA2, RAN1).")
        self.wg_filter.selectionChanged.connect(self.refresh_table)

        self.adhoc_filter = QComboBox()
        self.adhoc_filter.addItems(["All Meetings", "Regular", "Ad-Hoc / BIS"])
        self.adhoc_filter.setToolTip("Show all meetings, or filter by Regular vs. Ad-Hoc/BIS.")
        self.adhoc_filter.currentTextChanged.connect(self.refresh_table)

        self.type_filter = QComboBox()
        self.type_filter.addItems(["All Types", "In-Person", "Electronic"])
        self.type_filter.setToolTip("Filter by In-Person or Electronic (eMeetings).")
        self.type_filter.currentTextChanged.connect(self.refresh_table)

        right_layout.addWidget(self.wg_filter)
        right_layout.addWidget(self.adhoc_filter)
        right_layout.addWidget(self.type_filter)

        right_layout.addWidget(QLabel("Search (No. or Name):"))
        self.search_input = QLineEdit()
        self.search_input.setToolTip("Search across meeting numbers, locations, and names.")
        self.search_input.textChanged.connect(self.refresh_table)
        right_layout.addWidget(self.search_input)

        # Smart Global TDoc Search
        right_layout.addWidget(QLabel("Global TDoc Search:"))
        global_search_layout = QHBoxLayout()
        self.global_tdoc_input = QLineEdit()
        self.global_tdoc_input.setPlaceholderText("e.g., S2-2605740")
        self.global_tdoc_input.setMinimumWidth(120)
        self.global_tdoc_input.setToolTip("Type a valid TDoc. Press Enter to instantly download and open it.")

        self.btn_open_tdoc = QPushButton("📄 Doc")
        self.btn_open_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc.setFixedHeight(28)
        self.btn_open_tdoc.setObjectName("primaryBtn")
        self.btn_open_tdoc.setVisible(False)

        self.btn_open_meeting = QPushButton("🗓️ Mtg")
        self.btn_open_meeting.setCursor(Qt.PointingHandCursor)
        self.btn_open_meeting.setFixedHeight(28)
        self.btn_open_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_meeting.setVisible(False)

        global_search_layout.addWidget(self.global_tdoc_input)
        global_search_layout.addWidget(self.btn_open_tdoc)
        global_search_layout.addWidget(self.btn_open_meeting)
        right_layout.addLayout(global_search_layout)

        self.enable_dates_cb = QCheckBox("Filter by Date Range")
        self.enable_dates_cb.setToolTip("Enable to filter meetings within a specific date range.")
        self.enable_dates_cb.toggled.connect(self._toggle_date_inputs)
        self.enable_dates_cb.toggled.connect(self.refresh_table)
        right_layout.addWidget(self.enable_dates_cb)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))
        self.date_from.dateChanged.connect(self.refresh_table)
        self.date_from.setEnabled(False)
        right_layout.addWidget(self.date_from)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate().addYears(1))
        self.date_to.dateChanged.connect(self.refresh_table)
        self.date_to.setEnabled(False)
        right_layout.addWidget(self.date_to)
        self.enable_dates_cb.setChecked(True)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        right_layout.addWidget(line)

        # Scrape Configuration Toggle
        self.scrape_toggle_btn = QPushButton("⚙️ Scrape Configuration (Click to Expand)")
        self.scrape_toggle_btn.setCheckable(True)
        self.scrape_toggle_btn.setCursor(Qt.PointingHandCursor)
        self.scrape_toggle_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)

        self.scrape_frame = QFrame()
        self.scrape_frame.setVisible(False)
        scrape_layout = QVBoxLayout(self.scrape_frame)
        scrape_layout.setContentsMargins(10, 4, 0, 4)

        self.chk_wg = QCheckBox("Check for New Folders")
        self.chk_wg.setChecked(True)
        self.chk_dyna = QCheckBox("Update Metadata")
        self.chk_dyna.setChecked(True)
        self.chk_docs = QCheckBox("Deep Scrape 'Docs/'")
        self.chk_docs.setChecked(True)

        scrape_layout.addWidget(self.chk_wg)
        scrape_layout.addWidget(self.chk_dyna)
        scrape_layout.addWidget(self.chk_docs)

        self.scrape_toggle_btn.toggled.connect(self.scrape_frame.setVisible)
        right_layout.addWidget(self.scrape_toggle_btn)
        right_layout.addWidget(self.scrape_frame)

        # Local Cache Folder Selector
        right_layout.addWidget(QLabel("Local Cache Directory:"))
        cache_layout = QHBoxLayout()
        self.dl_dir_input = QLineEdit()
        self.dl_dir_input.setText(self.settings.cache_dir)
        self.dl_dir_input.editingFinished.connect(lambda: self.settings.save_settings(self.dl_dir_input.text().strip()))

        browse_btn = QPushButton("...")
        browse_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        browse_btn.setFixedWidth(32)
        browse_btn.clicked.connect(self._browse_cache_dir)

        cache_layout.addWidget(self.dl_dir_input)
        cache_layout.addWidget(browse_btn)
        right_layout.addLayout(cache_layout)
        right_layout.addStretch()

        # Action Buttons
        self.update_btn = QPushButton("🔄 Sync All Meetings")
        self.update_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.update_btn.clicked.connect(lambda: self.update_db_requested.emit(
            self.chk_wg.isChecked(), self.chk_docs.isChecked(), self.chk_dyna.isChecked()
        ))
        right_layout.addWidget(self.update_btn)

        self.btn_add_meeting = QPushButton("➕ Add / Fetch Meeting...")
        self.btn_add_meeting.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_add_meeting.clicked.connect(self._open_add_meeting_dialog)
        right_layout.addWidget(self.btn_add_meeting)

        self.btn_export_merged = QPushButton("📥 Export Merged TDocs (Excel)")
        self.btn_export_merged.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_merged.clicked.connect(self._export_merged_tdocs)
        right_layout.addWidget(self.btn_export_merged)

        self.delete_all_btn = QPushButton("🗑️ Clear All Meetings")
        self.delete_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.delete_all_btn.clicked.connect(self._confirm_delete_all)
        right_layout.addWidget(self.delete_all_btn)

        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([750, 250])
        main_layout.addWidget(self.splitter)
        self._populate_filters()

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
            self.cmp_thread.ui_log_msg.connect(self._handle_compare_log)
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setText("⚖️ Compare in Word"))
            self.cmp_thread.finished.connect(lambda: self.btn_compare.setEnabled(True))
            self.cmp_thread.start()

    def _handle_compare_log(self, msg: str, level: int):
        if level == logging.ERROR:
            print(f"🔴 {msg}")
        elif level == logging.WARNING:
            print(f"🟠 {msg}")
        else:
            print(f"🔵 {msg}")

    def _populate_filters(self):
        wgs = self.db.get_working_groups()
        self.wg_filter.blockSignals(True)
        self.wg_filter.updateItems(wgs)
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
        menu.clear()
        selected_rows = self.table.selectionModel().selectedRows()
        if len(selected_rows) > 1 and any(r.row() == row_idx for r in selected_rows):
            menu.addAction(f"🔄 Sync selected ({len(selected_rows)} meetings)").triggered.connect(
                lambda _, rows=selected_rows: self._emit_multi_sync(rows))
            menu.addSeparator()
            menu.addAction(f"🗑️ Delete selected ({len(selected_rows)} meetings)").triggered.connect(
                lambda _, rows=selected_rows: self._emit_multi_delete(rows))
        else:
            menu.addAction("ℹ️ Meeting Info").triggered.connect(lambda _, d=row_data: self.show_meeting_info(d))

            mtg_id = row_data.get("mtg_id")
            if mtg_id:
                menu.addAction("🖥️ 3GU Meeting Portal").triggered.connect(
                    lambda _, m=mtg_id: webbrowser.open(f"https://portal.3gpp.org/Home.aspx#/meeting?MtgId={m}")
                )

                end_date_str = row_data.get("end_date", "")
                end_date = QDate.fromString(end_date_str, "yyyy-MM-dd")
                current_date = QDate.currentDate()

                if not end_date.isValid() or end_date >= current_date:
                    menu.addAction("📝 Reserve / Contribute TDoc").triggered.connect(
                        lambda _, m=mtg_id: webbrowser.open(
                            f"https://portal.3gpp.org/ngppapp/CreateTdoc.Aspx?mode=create&meetingId={m}")
                    )

            menu.addAction("🔄 Sync this Meeting").triggered.connect(lambda: self.update_specific_requested.emit(
                [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}], self.chk_wg.isChecked(),
                self.chk_docs.isChecked(), self.chk_dyna.isChecked()))
            menu.addSeparator()

            raw_url = row_data.get("url_key", "")
            if raw_url:
                full_ftp_url = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url}"
                menu.addAction("🌐 Open Main Folder (FTP)").triggered.connect(
                    lambda _, u=full_ftp_url: webbrowser.open(u))

            docs_url = row_data.get("docs_folder_url")
            if docs_url:
                menu.addAction("📂 Open Documents Folder").triggered.connect(lambda _, u=docs_url: webbrowser.open(u))

            folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")

            if folder_name:
                current_cache = self.dl_dir_input.text().strip() if hasattr(self,
                                                                            'dl_dir_input') else self.settings.cache_dir
                local_path = Path(current_cache) / folder_name

                if local_path.exists() and local_path.is_dir():
                    menu.addAction("📁 Open Local Cache Folder").triggered.connect(
                        lambda _, p=str(local_path): os.startfile(p) if hasattr(os, 'startfile') else webbrowser.open(
                            f"file:///{p}")
                    )

                    menu.addAction("📋 Download Agenda CSV").triggered.connect(
                        lambda _, d=row_data: self._download_agenda_csv(d)
                    )

                if mtg_id:
                    status = row_data.get('tdoc_btn_status', 'na')
                    if status == 'open':
                        menu.addAction("📗 Open TDocs List").triggered.connect(
                            lambda _, d=row_data: self._open_tdocs_window(d, d.get('tdoc_filepath'))
                        )
                    elif status == 'get':
                        menu.addAction("⬇️ Get TDocs List").triggered.connect(
                            lambda _, d=row_data, r=row_idx: self._download_and_open_tdocs(d, r)
                        )

                if docs_url:
                    menu.addAction("📥 Cache TDocs (Docs/)").triggered.connect(
                        lambda _, u=docs_url, p=local_path: self._start_tdocs_caching(u, p)
                    )

            wg_name = row_data.get("wg_name", "")
            meeting_name = row_data.get("name", "")
            start_date = row_data.get("start_date", "")
            end_date = row_data.get("end_date", "")
            is_elec = row_data.get("is_electronic", 0)

            if self.db.is_active_sync_meeting(wg_name, start_date, end_date, is_elec):
                menu.addSeparator()
                sync_wg = "SA3LI" if wg_name == "SA3" and "LI" in meeting_name.upper() else wg_name
                sync_base_url = f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}"

                menu.addAction("🔄 Open SYNC folder (FTP)").triggered.connect(
                    lambda _, u=sync_base_url: webbrowser.open(u))
                menu.addAction("📂 Open SYNC Documents folder").triggered.connect(
                    lambda _, u=f"{sync_base_url}/Docs": webbrowser.open(u))

            menu.addSeparator()
            menu.addAction("🗑️ Delete this Meeting").triggered.connect(lambda: self._confirm_delete_specific(
                [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}]))

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
        try:
            last_id, last_num, last_wg = self.settings.get_last_meeting()

            if not last_id or not last_num:
                QMessageBox.information(self, "No History",
                                        "No recent meeting history found. Please open a meeting first.")
                return

            wg_filter = [last_wg] if last_wg else None
            results = self.db.search_meetings(wg_name=wg_filter, search_term=last_num)
            target_meeting = next((m for m in results if m.get("mtg_id") == last_id), None)

            if not target_meeting:
                results = self.db.search_meetings(search_term=last_num)
                target_meeting = next((m for m in results if m.get("mtg_id") == last_id), None)

            display_tag = f"{last_wg}#{last_num}" if last_wg else f"'{last_num}'"

            if not target_meeting:
                QMessageBox.warning(self, "Not Found",
                                    f"Meeting {display_tag} could not be found in the database.\nIt may have been cleared or the database was updated.")
                return

            filepath = self._get_tdoc_list_path(target_meeting)

            if filepath and filepath.exists():
                self._open_tdocs_window(target_meeting, str(filepath))
            else:
                self._download_and_open_tdocs(target_meeting)

        except Exception as e:
            QMessageBox.critical(self, "Launch Error", f"Could not open last meeting:\n{e}")

    def _start_tdocs_caching(self, docs_url: str, local_path: Path):
        from modules.meetings.core.tdocs_cacher import TDocsCacherThread
        self.update_btn.setText("⏳ Caching TDocs...")
        self.update_btn.setEnabled(False)
        self.cacher_thread = TDocsCacherThread(docs_url, local_path, self)
        self.cacher_thread.finished.connect(self._on_tdocs_caching_finished)
        self.cacher_thread.start()

    def _on_tdocs_caching_finished(self, success: bool, msg: str):
        self.update_btn.setText("🔄 Sync All Meetings")
        self.update_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Caching Complete", msg)
        else:
            QMessageBox.warning(self, "Caching Failed", msg)

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
        meetings_data = getattr(self.table_model, '_data', [])
        if not meetings_data:
            QMessageBox.warning(self, "No Data", "There are no meetings currently visible in the table to export.")
            return

        default_name = f"Merged_TDocs_{QDate.currentDate().toString('yyyy-MM-dd')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged TDocs",
                                                   str(Path.home() / "Desktop" / default_name), "Excel Files (*.xlsx)")
        if not save_path:
            return

        reply = QMessageBox.question(self, 'Force Download?',
                                     "Do you want to force a fresh download of all Excel files from 3GPP?\n\n"
                                     "• Select 'Yes' to fetch the absolute latest updates.\n"
                                     "• Select 'No' to use your local cache for files you've already downloaded.",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        force_download = (reply == QMessageBox.Yes)

        self.btn_export_merged.setText("⏳ Merging...")
        self.btn_export_merged.setEnabled(False)

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        self.merger_thread = TDocsMergerThread(meetings_data, force_download, save_path, current_cache, self)
        self.merger_thread.progress.connect(lambda msg: self.btn_export_merged.setText(f"⏳ {msg[:20]}..."))
        self.merger_thread.finished.connect(self._on_merger_finished)
        self.merger_thread.start()

    def _on_merger_finished(self, success: bool, msg: str):
        self.btn_export_merged.setText("📥 Export Merged TDocs (Excel)")
        self.btn_export_merged.setEnabled(True)

        if success:
            QMessageBox.information(self, "Export Complete", msg)
            if QMessageBox.question(self, "Open File", "Would you like to open the generated Excel file now?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                path_match = re.search(r'Saved to:\n(.*)', msg)
                if path_match:
                    filepath = path_match.group(1).strip()
                    try:
                        os.startfile(filepath) if hasattr(os, 'startfile') else webbrowser.open(f"file:///{filepath}")
                    except Exception as e:
                        logging.error(f"Could not open merged Excel: {e}")
        else:
            QMessageBox.warning(self, "Export Failed", msg)