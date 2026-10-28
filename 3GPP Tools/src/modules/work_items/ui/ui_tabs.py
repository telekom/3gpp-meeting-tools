# --- File: src/modules/work_items/ui/ui_tabs.py ---
import webbrowser
from datetime import datetime
from pathlib import Path
import logging

from PyQt5.QtCore import Qt, QModelIndex, QTimer, QEvent, QRect, pyqtSignal
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableView, QHeaderView,
    QPushButton, QProgressBar, QMessageBox, QLineEdit, QMenu,
    QApplication, QComboBox, QDialog, QFrame,
    QScrollArea, QFormLayout
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEGMENT_BTN_STYLE,
    TOGGLE_PILL_STYLE,
    FAV_PILL_STYLE,
    DISMISSIBLE_TAG_STYLE,
    CLEAR_ALL_BTN_STYLE,
    SPEC_FILTER_COMBO_STYLE,
    SPEC_FILTER_INPUT_STYLE,
    WG_DESCRIPTIONS,
    get_wg_colors,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import WorkItemsScraperThread, TargetedWIScraperThread
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.work_items.ui.work_items_dialogs import WorkItemInfoDialog, LinkedSpecsDialog, AddWorkItemDialog

logger = logging.getLogger(__name__)


from modules.work_items.ui.work_items_delegates import (
    TrackingDelegate,
    WgBadgeDelegate,
    WidDelegate,
    ReleaseBadgeDelegate,
    StatusBadgeDelegate,
    SpecsDelegate,
    TimelineDelegate,
    RemarksDelegate,
)
from modules.work_items.ui.work_items_model import WorkItemsTableModel

__all__ = [
    "TrackingDelegate",
    "WgBadgeDelegate",
    "WidDelegate",
    "ReleaseBadgeDelegate",
    "StatusBadgeDelegate",
    "SpecsDelegate",
    "TimelineDelegate",
    "RemarksDelegate",
    "WorkItemsTableModel",
    "WorkItemsTab",
]


class WorkItemsTab(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db_path = db_path
        self.db = WorkItemsDatabase(db_path)
        self.settings = WorkItemsSettings()

        self._chunk_size = 60
        self._current_offset = 0
        self._total_count = 0
        self._loaded_data = []
        self._is_loading_chunk = False

        # Tracking state (Personal Interest & Delegation Assignment)
        self.starred, self.following = self.settings.get_tracking_sets()
        self.starred_only = False
        self.following_only = False
        self._current_tsg = ""
        self.tsg_buttons = {}
        self._rendered_tag_labels = None

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(350)
        self.search_timer.timeout.connect(lambda: self.refresh_table(reset_pagination=True))

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        self._populate_filters()
        self._load_filters()
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _save_filters(self):
        filters = {
            "search": self.search_input.text().strip(),
            "tsg": self._current_tsg,
            "releases": self.release_combo.getCheckedItems(),
            "wgs": self.wg_combo.getCheckedItems(),
            "status": self.status_combo.currentData() or "active",
            "starred_only": self.starred_only,
            "following_only": self.following_only,
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        filters = self.settings.get_filters()
        self.search_input.blockSignals(True)
        self.release_combo.blockSignals(True)
        self.wg_combo.blockSignals(True)
        self.status_combo.blockSignals(True)
        try:
            if not filters:
                idx = self.status_combo.findData("active")
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)
                return

            if "search" in filters:
                self.search_input.setText(filters["search"])

            if "tsg" in filters:
                self._current_tsg = filters["tsg"]

            if "releases" in filters:
                self._apply_checked_items(self.release_combo, filters["releases"])

            if "wgs" in filters:
                self._apply_checked_items(self.wg_combo, filters["wgs"])

            if "status" in filters:
                idx = self.status_combo.findData(filters["status"])
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)
            else:
                idx = self.status_combo.findData("active")
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)

            if "starred_only" in filters:
                self.starred_only = bool(filters["starred_only"])
            if "following_only" in filters:
                self.following_only = bool(filters["following_only"])
        finally:
            self.search_input.blockSignals(False)
            self.release_combo.blockSignals(False)
            self.wg_combo.blockSignals(False)
            self.status_combo.blockSignals(False)

    def _apply_checked_items(self, combo, items_to_check):
        model = combo.model()
        row_count = model.rowCount()
        if row_count == 0:
            return

        model.blockSignals(True)
        try:
            if not items_to_check:
                for row in range(row_count):
                    if hasattr(model, 'item'):
                        item = model.item(row)
                        if item:
                            item.setCheckState(Qt.Checked)
                    else:
                        index = model.index(row, 0)
                        model.setData(index, Qt.Checked, Qt.CheckStateRole)
            else:
                items_to_check_stripped = {str(x).strip() for x in items_to_check}
                all_checked = True
                for row in range(1, row_count):
                    if hasattr(model, 'item'):
                        item = model.item(row)
                        if item:
                            user_val = str(item.data(Qt.UserRole) or item.text()).strip()
                            checked = (user_val in items_to_check_stripped or item.text().strip() in items_to_check_stripped)
                            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                            if not checked:
                                all_checked = False
                    else:
                        index = model.index(row, 0)
                        text = str(model.data(index, Qt.DisplayRole)).strip()
                        checked = text in items_to_check_stripped
                        model.setData(index, Qt.Checked if checked else Qt.Unchecked, Qt.CheckStateRole)
                        if not checked:
                            all_checked = False

                if hasattr(model, 'item'):
                    root_item = model.item(0)
                    if root_item:
                        root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
                else:
                    root_index = model.index(0, 0)
                    model.setData(root_index, Qt.Checked if all_checked else Qt.Unchecked, Qt.CheckStateRole)
        finally:
            model.blockSignals(False)

        if hasattr(combo, 'updateText'):
            combo.updateText()
        elif hasattr(combo, 'update'):
            combo.update()
        elif hasattr(combo, 'repaint'):
            combo.repaint()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        # --- ROW 1: HEADER & ACTION BUTTONS ---
        header_layout = QHBoxLayout()
        header_lbl = QLabel("<b>📋 3GPP Work Items (WIs)</b>")
        header_lbl.setStyleSheet("font-size: 16px; color: #1E293B;")

        self.add_wi_btn = QPushButton("➕ Add WI")
        self.add_wi_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_wi_btn.setToolTip("Add and scrape an individual Work Item by ID or Portal URL without running a full sync.")
        self.add_wi_btn.setCursor(Qt.PointingHandCursor)
        self.add_wi_btn.clicked.connect(lambda: self._open_add_wi_dialog())

        self.sync_tracked_btn = QPushButton("🎯 Update Tracked (0)")
        self.sync_tracked_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.sync_tracked_btn.setToolTip("Quickly update details and remarks for only your Starred and Following Work Items.")
        self.sync_tracked_btn.setCursor(Qt.PointingHandCursor)
        self.sync_tracked_btn.clicked.connect(self._sync_tracked_wis)

        self.sync_btn = QPushButton("🔄 Sync 3GPP WIs")
        self.sync_btn.setObjectName("primaryBtn")
        self.sync_btn.setToolTip("Click to download and synchronize 3GPP Work Items in parallel from the server.")
        self.sync_btn.clicked.connect(self._start_sync)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(180)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #64748B; font-style: italic; font-size: 11px;")

        header_layout.addWidget(header_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.status_lbl)
        header_layout.addWidget(self.progress_bar)
        header_layout.addWidget(self.add_wi_btn)
        header_layout.addWidget(self.sync_tracked_btn)
        header_layout.addWidget(self.sync_btn)
        main_layout.addLayout(header_layout)

        # --- ROW 2: INLINE SEARCH & FILTER BAR WITH TSG QUICK SCOPE ---
        search_layout = QHBoxLayout()
        search_layout.setSpacing(8)

        search_layout.addWidget(QLabel("<b>🔍 Search:</b>"))

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.search_input.setPlaceholderText("Search Code, Acronym, or Name...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setToolTip("Filter the table instantly by typing keywords or WI code.")
        self.search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.search_input, stretch=3)

        # TSG Quick-Scope buttons
        tsg_lbl = QLabel("TSG:")
        tsg_lbl.setStyleSheet("color: #475569; font-weight: bold; font-size: 11px;")
        search_layout.addWidget(tsg_lbl)

        self.tsg_buttons = {}
        for tsg in ["All", "SA", "RAN", "CT"]:
            btn = QPushButton(tsg)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(f"Filter Work Items to {tsg if tsg != 'All' else 'all'} Technical Specification Groups")
            btn.clicked.connect(lambda _, t=tsg: self._on_tsg_clicked(t))
            self.tsg_buttons[tsg] = btn
            search_layout.addWidget(btn)
        self.tsg_buttons["All"].setChecked(True)

        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(130)
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(120)
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.wg_combo)

        self.status_combo = QComboBox()
        self.status_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.status_combo.setToolTip("Filter by Work Item completion status (End Date)")
        self.status_combo.setMinimumWidth(140)
        self.status_combo.addItem("🟢 Active / Ongoing", "active")
        self.status_combo.addItem("All Statuses", "all")
        self.status_combo.addItem("🏁 Finished", "finished")
        self.status_combo.currentIndexChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.status_combo)

        self.clear_filters_btn = QPushButton("❌ Clear")
        self.clear_filters_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_filters_btn.setToolTip("Reset all search fields and active filters.")
        self.clear_filters_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filters_btn.clicked.connect(self._clear_all_filters)
        search_layout.addWidget(self.clear_filters_btn)

        main_layout.addLayout(search_layout)

        # --- ROW 3: TRACKING FILTER CHIPS, ACTIVE DISMISSIBLE TAGS & RESULTS COUNTER ---
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(6)

        chips_lbl = QLabel("Filter:")
        chips_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        chips_layout.addWidget(chips_lbl)

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

        # Divider line
        div_lbl = QLabel("|")
        div_lbl.setStyleSheet("color: #CBD5E1; margin: 0 4px;")
        chips_layout.addWidget(div_lbl)

        # Active Dismissible Tags Container
        self.active_tags_widget = QWidget()
        self.active_tags_container_layout = QHBoxLayout(self.active_tags_widget)
        self.active_tags_container_layout.setContentsMargins(0, 0, 0, 0)
        self.active_tags_container_layout.setSpacing(4)
        chips_layout.addWidget(self.active_tags_widget)

        # Clear all active tags button
        self.clear_all_btn = QPushButton("✕ Clear All")
        self.clear_all_btn.setCursor(Qt.PointingHandCursor)
        self.clear_all_btn.setStyleSheet(CLEAR_ALL_BTN_STYLE)
        self.clear_all_btn.setToolTip("Clear all active search and filter constraints.")
        self.clear_all_btn.setVisible(False)
        self.clear_all_btn.clicked.connect(self._clear_all_filters)
        chips_layout.addWidget(self.clear_all_btn)

        chips_layout.addStretch()

        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet("font-weight: bold; color: #64748B; font-size: 11px;")
        chips_layout.addWidget(self.count_label)

        main_layout.addLayout(chips_layout)

        # --- TABLE VIEW ---
        self.table = QTableView()
        self.table_model = WorkItemsTableModel()
        self.table_model.set_tracking_sets(self.starred, self.following)
        self.table.setModel(self.table_model)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setShowGrid(True)
        self.table.setStyleSheet("""
            QTableView {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                gridline-color: #F1F5F9;
                selection-background-color: #EBF3FC;
                selection-color: #1E293B;
            }
            QTableView::item {
                padding: 2px 4px;
                border-bottom: 1px solid #F1F5F9;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: bold;
                font-size: 11px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
                border-right: 1px solid #E2E8F0;
                padding: 5px 6px;
            }
        """)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.doubleClicked.connect(self._on_row_double_clicked)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)

        # Column 0: Track (Fixed)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 64)
        self.tracking_delegate = TrackingDelegate(self.table)
        self.tracking_delegate.star_toggled.connect(self._toggle_star)
        self.tracking_delegate.follow_toggled.connect(self._toggle_follow)
        self.table.setItemDelegateForColumn(0, self.tracking_delegate)

        # Column 1: Code
        self.table.setColumnWidth(1, 75)

        # Column 2: Acronym
        self.table.setColumnWidth(2, 185)

        # Column 3: Name (Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        # Column 4: WG
        self.table.setColumnWidth(4, 95)
        self.wg_delegate = WgBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(4, self.wg_delegate)

        # Column 5: Latest WID
        self.table.setColumnWidth(5, 105)
        self.wid_delegate = WidDelegate(self.table)
        self.wid_delegate.link_clicked.connect(self._log_and_forward_action)
        self.table.setItemDelegateForColumn(5, self.wid_delegate)

        # Column 6: Release
        self.table.setColumnWidth(6, 75)
        self.release_delegate = ReleaseBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(6, self.release_delegate)

        # Column 7: Status
        self.table.setColumnWidth(7, 95)
        self.status_delegate = StatusBadgeDelegate(self.table)
        self.table.setItemDelegateForColumn(7, self.status_delegate)

        # Column 8: Specs
        self.table.setColumnWidth(8, 65)
        self.specs_delegate = SpecsDelegate(self.table)
        self.specs_delegate.specs_clicked.connect(self._on_specs_clicked)
        self.table.setItemDelegateForColumn(8, self.specs_delegate)

        # Column 9: Timeline
        self.table.setColumnWidth(9, 135)
        self.timeline_delegate = TimelineDelegate(self.table)
        self.table.setItemDelegateForColumn(9, self.timeline_delegate)

        # Column 10: Remarks (Stretch)
        header.setSectionResizeMode(10, QHeaderView.Stretch)
        self.remarks_delegate = RemarksDelegate(self.table)
        self.table.setItemDelegateForColumn(10, self.remarks_delegate)

        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll)
        main_layout.addWidget(self.table)

    def _on_tsg_clicked(self, tsg: str):
        if tsg == "All":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg

        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _update_active_tags(self):
        """Re-render active filter tags dynamically with 1-click removal."""
        tags = []

        query = self.search_input.text().strip()
        if query:
            disp_q = query if len(query) <= 18 else f"{query[:15]}..."
            tags.append((f'🔍 "{disp_q}" ✕', lambda: self.search_input.clear()))

        if self._current_tsg:
            tags.append((f'TSG: {self._current_tsg} ✕', lambda: self._on_tsg_clicked("All")))

        if hasattr(self.wg_combo, "isFiltered") and self.wg_combo.isFiltered():
            selected_wgs = self.wg_combo.getCheckedItems()
            if selected_wgs:
                w_str = ", ".join(selected_wgs)
                disp_w = w_str if len(w_str) <= 18 else f"{w_str[:15]}..."
                tags.append((f'WG: {disp_w} ✕', lambda: (self._apply_checked_items(self.wg_combo, []), self.refresh_table(True))))

        if hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered():
            selected_rels = self.release_combo.getCheckedItems()
            if selected_rels:
                r_str = ", ".join(selected_rels)
                disp_r = r_str if len(r_str) <= 18 else f"{r_str[:15]}..."
                tags.append((f'Rel: {disp_r} ✕', lambda: (self._apply_checked_items(self.release_combo, []), self.refresh_table(True))))

        cur_status = self.status_combo.currentData()
        if cur_status and cur_status != "active":
            status_name = "Finished" if cur_status == "finished" else "All Statuses"
            tags.append((f'Status: {status_name} ✕', lambda: self.status_combo.setCurrentIndex(0)))

        if self.following_only:
            tags.append(('🎯 Following ✕', self._toggle_following_filter))

        if self.starred_only:
            tags.append(('⭐ Starred ✕', self._toggle_starred_filter))

        tag_labels = [label for label, _ in tags]
        if getattr(self, "_rendered_tag_labels", None) == tag_labels:
            return
        self._rendered_tag_labels = tag_labels

        while self.active_tags_container_layout.count():
            item = self.active_tags_container_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        has_tags = len(tags) > 0
        if hasattr(self, "clear_all_btn"):
            self.clear_all_btn.setVisible(has_tags)

        for label, callback in tags:
            tag_btn = QPushButton(label)
            tag_btn.setCursor(Qt.PointingHandCursor)
            tag_btn.setToolTip("Click to remove this filter constraint")
            tag_btn.setStyleSheet(DISMISSIBLE_TAG_STYLE)
            tag_btn.clicked.connect(callback)
            self.active_tags_container_layout.addWidget(tag_btn)

    def _toggle_starred_filter(self):
        self.starred_only = not self.starred_only
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _toggle_following_filter(self):
        self.following_only = not self.following_only
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _update_chip_styles(self):
        starred_count = len(self.starred)
        following_count = len(self.following)
        tracked_count = len(self.starred | self.following)

        self.starred_chip.setText(f"⭐ Starred ({starred_count})")
        self.following_chip.setText(f"🎯 Following ({following_count})")
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setText(f"🎯 Update Tracked ({tracked_count})")
            self.sync_tracked_btn.setEnabled(tracked_count > 0)

        self.starred_chip.setChecked(self.starred_only)
        self.following_chip.setChecked(self.following_only)

        if hasattr(self, "tsg_buttons"):
            for tsg_val, btn in self.tsg_buttons.items():
                btn.setChecked((tsg_val == "All" and not self._current_tsg) or (tsg_val == self._current_tsg))

    def _get_active_filter_codes(self):
        """Returns list of WI codes if filter chips are active, or None if no tracking filter is active."""
        if not self.starred_only and not self.following_only:
            return None
        codes = set()
        if self.starred_only:
            codes |= self.starred
        if self.following_only:
            codes |= self.following
        return sorted(list(codes))

    def _toggle_star(self, wi_code: str):
        code = str(wi_code).strip()
        if not code:
            return
        if code in self.starred:
            self.starred.remove(code)
        else:
            self.starred.add(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()

        if self.starred_only:
            self.refresh_table(reset_pagination=True)
        else:
            self.table_model.refresh_row_by_code(code)

    def _toggle_follow(self, wi_code: str):
        code = str(wi_code).strip()
        if not code:
            return
        if code in self.following:
            self.following.remove(code)
        else:
            self.following.add(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()

        if self.following_only:
            self.refresh_table(reset_pagination=True)
        else:
            self.table_model.refresh_row_by_code(code)

    def _batch_set_starred(self, wi_codes: list, add: bool):
        for c in wi_codes:
            code = str(c).strip()
            if not code:
                continue
            if add:
                self.starred.add(code)
            else:
                self.starred.discard(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _batch_set_following(self, wi_codes: list, add: bool):
        for c in wi_codes:
            code = str(c).strip()
            if not code:
                continue
            if add:
                self.following.add(code)
            else:
                self.following.discard(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _sync_tracked_wis(self):
        tracked = sorted(list(self.starred | self.following))
        if not tracked:
            QMessageBox.information(
                self, "No Tracked Work Items",
                "You have no Starred or Following Work Items to update.\n\n"
                "Star or Follow Work Items first using the Track column or right-click menu."
            )
            return
        self._start_targeted_sync(tracked)

    def _clear_all_filters(self):
        self.search_input.clear()
        self._current_tsg = ""
        self._apply_checked_items(self.release_combo, [])
        self._apply_checked_items(self.wg_combo, [])
        self.status_combo.setCurrentIndex(0)
        self.starred_only = False
        self.following_only = False
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table(reset_pagination=True)

    def _on_scroll(self, value):
        max_val = self.table.verticalScrollBar().maximum()
        if value >= max_val - 5 and not self._is_loading_chunk:
            if len(self._loaded_data) < self._total_count:
                self._load_next_chunk()

    def _log_and_forward_action(self, text, action):
        self.global_action_requested.emit(text, action)

    def _populate_filters(self):
        options = self.db.get_filter_options()

        self.release_combo.blockSignals(True)
        self.release_combo.updateItems(options.get('releases', []))
        self.release_combo.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.updateItems(options.get('groups', []))
        self.wg_combo.blockSignals(False)

    def _resolve_effective_wgs(self) -> list:
        if hasattr(self.wg_combo, "isFiltered") and self.wg_combo.isFiltered():
            return self.wg_combo.getCheckedItems()
        elif self._current_tsg:
            all_wgs = self.db.get_filter_options().get('groups', [])
            if self._current_tsg == "SA":
                return [g for g in all_wgs if g.startswith("SA")]
            elif self._current_tsg == "RAN":
                return [g for g in all_wgs if g.startswith("RAN") or g.startswith("GERAN")]
            elif self._current_tsg == "CT":
                return [g for g in all_wgs if g.startswith("CT") or g.startswith("CN")]
        return []

    def refresh_table(self, reset_pagination: bool = True):
        self.save_filters_timer.start()

        if reset_pagination:
            self._current_offset = 0
            self._loaded_data = []

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems() if (hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered()) else []
        selected_wgs = self._resolve_effective_wgs()
        selected_status = self.status_combo.currentData() or "active"
        filter_codes = self._get_active_filter_codes()

        self._total_count = self.db.count_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            wi_codes=filter_codes
        )

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset,
            wi_codes=filter_codes
        )

        self._loaded_data = chunk
        self.table_model.set_search_term(search_term)
        self.table_model.update_data(self._loaded_data)
        self._update_count_label()
        self._update_chip_styles()
        self._update_active_tags()

    def _load_next_chunk(self):
        self._is_loading_chunk = True
        self._current_offset += self._chunk_size

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems() if (hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered()) else []
        selected_wgs = self._resolve_effective_wgs()
        selected_status = self.status_combo.currentData() or "active"
        filter_codes = self._get_active_filter_codes()

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset,
            wi_codes=filter_codes
        )

        if chunk:
            self.table_model.append_data(chunk)
            self._loaded_data.extend(chunk)
            self._update_count_label()

        self._is_loading_chunk = False

    def _update_count_label(self):
        loaded = len(self._loaded_data)
        query = self.search_input.text().strip()
        exact_count = 0
        if query:
            exact_count = sum(1 for r in self._loaded_data if str(r.get("acronym") or "").strip().casefold() == query.casefold())

        prefix = f'{self._total_count} matches for "{query}"' if query else f"{self._total_count} Work Items"
        if query and exact_count:
            prefix += f" · Exact acronym: {exact_count}"
        if self._total_count > loaded:
            self.count_label.setText(f"{prefix} · showing {loaded} (scroll for more)")
        else:
            self.count_label.setText(prefix)

    def _on_row_double_clicked(self, index):
        if not index.isValid():
            return
        row_data = self.table_model.get_row_data(index.row())
        wi_code = row_data.get('code')
        if wi_code:
            self._show_wi_info(wi_code)

    def _on_specs_clicked(self, row_idx: int):
        row_data = self.table_model.get_row_data(row_idx)
        wi_code = row_data.get('code')
        acronym = row_data.get('acronym', '')
        if wi_code:
            specs = self.db.get_linked_specs_for_wi(wi_code)
            dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
            dialog.exec_()

    def _show_wi_info(self, wi_code: str):
        details = self.db.get_work_item_details(wi_code)
        if details:
            code_str = str(wi_code).strip()
            dialog = WorkItemInfoDialog(
                details,
                parent=self,
                is_starred=code_str in self.starred,
                is_following=code_str in self.following,
                on_toggle_star=self._toggle_star,
                on_toggle_follow=self._toggle_follow,
            )
            dialog.exec_()

    def _open_add_wi_dialog(self, initial_query: str = ""):
        dialog = AddWorkItemDialog(self.db, initial_query=initial_query, parent=self)
        if dialog.exec_() == QDialog.Accepted and dialog.saved_wi_code:
            code = dialog.saved_wi_code

            # Update tracking if requested
            tracking_changed = False
            if dialog.is_starred_requested and code not in self.starred:
                self.starred.add(code)
                tracking_changed = True
            if dialog.is_following_requested and code not in self.following:
                self.following.add(code)
                tracking_changed = True

            if tracking_changed:
                self.settings.save_tracking_sets(self.starred, self.following)
                self.table_model.set_tracking_sets(self.starred, self.following)
                self._update_chip_styles()

            # Refresh dropdown filter options (new releases / WGs)
            self._populate_filters()

            # Clear search and active filter constraints so the new WI is clearly displayed
            self._clear_all_filters()
            self.search_input.setText(code)
            self.refresh_table(reset_pagination=True)
            self.status_lbl.setText(f"✅ Added Work Item #{code}")
            logger.info(f"✅ Added Work Item #{code} to database.")

    def _start_sync(self):
        logger.info("🔄 [Work Items] Starting full Work Items synchronization.")
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Syncing...")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(False)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.scraper_thread = WorkItemsScraperThread(self.db_path, self)
        self.scraper_thread.progress.connect(self._update_progress)
        self.scraper_thread.finished_sync.connect(self._on_sync_finished)
        self.scraper_thread.start()

    def _update_progress(self, current: int, total: int, msg: str):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_lbl.setText(msg)

    def _on_sync_finished(self, success: bool, msg: str):
        if success:
            logger.info(f"✅ [Work Items] {msg}")
        else:
            logger.error(f"❌ [Work Items] {msg}")

        self.sync_btn.setEnabled(True)
        self.sync_btn.setText("🔄 Sync 3GPP WIs")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(True)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(len(self.starred | self.following) > 0)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("")

        prev_releases = self.release_combo.getCheckedItems()
        prev_wgs = self.wg_combo.getCheckedItems()
        self._populate_filters()
        if prev_releases:
            self.release_combo.blockSignals(True)
            self._apply_checked_items(self.release_combo, prev_releases)
            self.release_combo.blockSignals(False)
        if prev_wgs:
            self.wg_combo.blockSignals(True)
            self._apply_checked_items(self.wg_combo, prev_wgs)
            self.wg_combo.blockSignals(False)
        self.refresh_table(reset_pagination=True)

        if success:
            QMessageBox.information(self, "Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Sync Failed", msg)

    def _show_context_menu(self, position):
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            menu = QMenu(self)
            add_wi_action = menu.addAction("➕ Add Work Item...")
            action = menu.exec_(self.table.viewport().mapToGlobal(position))
            if action == add_wi_action:
                self._open_add_wi_dialog()
            return

        menu = QMenu(self)
        len_indexes = len(selected_indexes)
        wi_code_list = [
            str(self.table_model.get_row_data(e.row()).get("code") or "").strip()
            for e in selected_indexes
        ]
        wi_code_list = [code for code in wi_code_list if code]

        if len_indexes == 1:
            row_idx = selected_indexes[0].row()
            row_data = self.table_model.get_row_data(row_idx)
            wi_code = str(row_data.get('code') or "").strip()
            acronym = row_data.get('acronym', '')
            name = row_data.get('name', '')
            wg = row_data.get('wg_names', '')
            rel = row_data.get('release', '')

            if not wi_code:
                return

            # Tracking Actions
            is_starred = wi_code in self.starred
            is_following = wi_code in self.following

            star_label = "⭐ Remove from Starred" if is_starred else "⭐ Add to Starred (Personal Interest)"
            follow_label = "🎯 Unfollow Work Item" if is_following else "🎯 Follow Work Item (Assigned Duty)"

            star_action = menu.addAction(star_label)
            follow_action = menu.addAction(follow_label)
            menu.addSeparator()

            info_action = menu.addAction("ℹ️ View WI Details")
            wi_page_action = menu.addAction("🌐 Open WI Page (Portal)")
            local_specs_action = menu.addAction("📑 View Linked Specifications in DB")
            specs_action = menu.addAction("📂 Specifications Resulting from this WI (Web)")
            crs_action = menu.addAction("📄 CRs Related to this WI (Web)")
            menu.addSeparator()
            citation_action = menu.addAction("📋 Copy WI Citation")
            menu.addSeparator()
            add_wi_action = menu.addAction("➕ Add Work Item...")
            update_action = menu.addAction("🔄 Update WI")
            delete_action = menu.addAction("🗑️ Delete this Work Item")
            batch_star_action = None
            batch_unstar_action = None
            batch_follow_action = None
            batch_unfollow_action = None
        else:
            star_action = None
            follow_action = None
            info_action = None
            wi_page_action = None
            local_specs_action = None
            specs_action = None
            crs_action = None
            citation_action = None
            wi_code = None

            batch_star_action = menu.addAction(f"⭐ Star selected ({len_indexes} WIs)")
            batch_unstar_action = menu.addAction(f"☆ Unstar selected ({len_indexes} WIs)")
            batch_follow_action = menu.addAction(f"🎯 Follow selected ({len_indexes} WIs)")
            batch_unfollow_action = menu.addAction(f"🚫 Unfollow selected ({len_indexes} WIs)")
            menu.addSeparator()

            add_wi_action = menu.addAction("➕ Add Work Item...")
            update_action = menu.addAction(f"🔄 Update WIs ({len_indexes} WIs)")
            delete_action = menu.addAction(f"🗑️ Delete selected Work Items ({len_indexes} WIs)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if len_indexes == 1:
            if action == star_action:
                self._toggle_star(wi_code)
            elif action == follow_action:
                self._toggle_follow(wi_code)
            elif action == info_action:
                self._show_wi_info(wi_code)
            elif action == wi_page_action:
                webbrowser.open(f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}")
            elif action == local_specs_action:
                specs = self.db.get_linked_specs_for_wi(wi_code)
                dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
                dialog.exec_()
            elif action == specs_action:
                webbrowser.open(f"https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={wi_code}")
            elif action == crs_action:
                webbrowser.open(f"https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem={wi_code}")
            elif action == citation_action:
                citation = f"WI {wi_code}: {acronym} — {name} ({rel}, {wg})"
                QApplication.clipboard().setText(citation)
            elif action == add_wi_action:
                self._open_add_wi_dialog()
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Deletion",
                    f"Are you sure you want to delete Work Item '{wi_code}' from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_item(wi_code)
                    self.starred.discard(wi_code)
                    self.following.discard(wi_code)
                    self.settings.save_tracking_sets(self.starred, self.following)
                    self.table_model.set_tracking_sets(self.starred, self.following)
                    self._update_chip_styles()
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync([wi_code])
        else:
            if action == batch_star_action:
                self._batch_set_starred(wi_code_list, True)
            elif action == batch_unstar_action:
                self._batch_set_starred(wi_code_list, False)
            elif action == batch_follow_action:
                self._batch_set_following(wi_code_list, True)
            elif action == batch_unfollow_action:
                self._batch_set_following(wi_code_list, False)
            elif action == add_wi_action:
                self._open_add_wi_dialog()
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Batch Deletion",
                    f"Are you sure you want to delete {len(wi_code_list)} selected Work Items from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_items(wi_code_list)
                    for c in wi_code_list:
                        self.starred.discard(c)
                        self.following.discard(c)
                    self.settings.save_tracking_sets(self.starred, self.following)
                    self.table_model.set_tracking_sets(self.starred, self.following)
                    self._update_chip_styles()
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync(wi_code_list)

    def _start_targeted_sync(self, wi_codes: list):
        logger.info(f"🔄 [Work Items] Starting targeted update for {len(wi_codes)} WI(s): {wi_codes}")
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Updating...")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(False)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.targeted_thread = TargetedWIScraperThread(self.db_path, wi_codes, self)
        self.targeted_thread.progress.connect(self._update_progress)
        self.targeted_thread.finished_sync.connect(self._on_sync_finished)
        self.targeted_thread.start()

    def shutdown(self):
        if hasattr(self, "search_timer"):
            self.search_timer.stop()
        if hasattr(self, "save_filters_timer"):
            self.save_filters_timer.stop()
        for th in (getattr(self, "scraper_thread", None), getattr(self, "targeted_thread", None)):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(150)

    def closeEvent(self, event):
        self.shutdown()
        super().closeEvent(event)
