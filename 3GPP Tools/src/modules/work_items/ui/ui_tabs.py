# --- File: src/modules/work_items/ui/ui_tabs.py ---
import webbrowser
from datetime import datetime
from pathlib import Path
import logging

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer, QEvent, QRect, pyqtSignal
from PyQt5.QtGui import QColor, QPalette, QBrush, QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableView, QHeaderView,
    QPushButton, QProgressBar, QMessageBox, QLineEdit, QMenu, QStyle,
    QApplication, QStyledItemDelegate, QComboBox, QDialog, QFrame,
    QScrollArea, QFormLayout
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import WorkItemsScraperThread, TargetedWIScraperThread
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.work_items.ui.work_items_dialogs import WorkItemInfoDialog, LinkedSpecsDialog, AddWorkItemDialog

logger = logging.getLogger(__name__)


class TrackingDelegate(QStyledItemDelegate):
    """Renders 1-click Star (Personal Interest) and Follow (Assigned Duty) toggles."""
    star_toggled = pyqtSignal(str)
    follow_toggled = pyqtSignal(str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        state = index.data(Qt.UserRole + 2) or {}
        is_starred = state.get("starred", False)
        is_following = state.get("following", False)

        rect = option.rect
        mid_x = rect.left() + rect.width() // 2
        star_rect = QRect(rect.left() + 3, rect.top() + 3, (rect.width() // 2) - 4, rect.height() - 6)
        follow_rect = QRect(mid_x + 1, rect.top() + 3, (rect.width() // 2) - 4, rect.height() - 6)

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        # Draw Starred pill/icon
        if is_starred:
            painter.setBrush(QColor("#FEF3C7"))
            painter.setPen(QColor("#F59E0B"))
            painter.drawRoundedRect(star_rect, 4, 4)
            painter.setPen(QColor("#B45309"))
            font = option.font
            font.setPointSize(font.pointSize() if font.pointSize() > 8 else 9)
            painter.setFont(font)
            painter.drawText(star_rect, Qt.AlignCenter, "⭐")
        else:
            painter.setPen(QColor("#CBD5E1"))
            painter.drawText(star_rect, Qt.AlignCenter, "☆")

        # Draw Following pill/icon
        if is_following:
            painter.setBrush(QColor("#DBEAFE"))
            painter.setPen(QColor("#3B82F6"))
            painter.drawRoundedRect(follow_rect, 4, 4)
            painter.setPen(QColor("#1E40AF"))
            font = option.font
            font.setPointSize(font.pointSize() if font.pointSize() > 8 else 9)
            painter.setFont(font)
            painter.drawText(follow_rect, Qt.AlignCenter, "🎯")
        else:
            painter.setPen(QColor("#CBD5E1"))
            painter.drawText(follow_rect, Qt.AlignCenter, "🎯")

        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            state = index.data(Qt.UserRole + 2) or {}
            wi_code = state.get("code", "")
            if not wi_code:
                return False
            mid_x = option.rect.left() + option.rect.width() // 2
            if event.pos().x() < mid_x:
                self.star_toggled.emit(wi_code)
            else:
                self.follow_toggled.emit(wi_code)
            return True
        return super().editorEvent(event, model, option, index)


class WidDelegate(QStyledItemDelegate):
    """Renders the Latest WID as a clickable hyperlink using the common palette."""
    link_clicked = pyqtSignal(str, str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text:
            return

        painter.save()
        font = option.font
        font.setUnderline(True)
        painter.setFont(font)

        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(QColor("#1E5C99"))

        painter.drawText(option.rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text:
                self.link_clicked.emit(text, "open_doc")
                return True
        return super().editorEvent(event, model, option, index)


class SpecsDelegate(QStyledItemDelegate):
    """Renders the linked specifications count as a standardized pill button."""
    specs_clicked = pyqtSignal(int)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            painter.save()
            painter.setPen(QColor("#94A3B8"))
            painter.drawText(option.rect, Qt.AlignCenter, "-")
            painter.restore()
            return

        badge_rect = QRect(option.rect.left() + 4, option.rect.top() + 4, option.rect.width() - 8, option.rect.height() - 8)

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(badge_rect, 10, 10)

        painter.setPen(QColor("#1E5C99"))
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text and text != "-":
                self.specs_clicked.emit(index.row())
                return True
        return super().editorEvent(event, model, option, index)


class RemarksDelegate(QStyledItemDelegate):
    """Draws the latest remark text with a clickable history badge."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.button_width = 45
        self.button_margin = 4

    def get_button_rect(self, option):
        rect = option.rect
        return QRect(
            rect.right() - self.button_width - self.button_margin,
            rect.top() + self.button_margin,
            self.button_width,
            rect.height() - (2 * self.button_margin)
        )

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)

        remarks_list = index.data(Qt.UserRole + 1)
        if not remarks_list:
            return

        latest_remark = remarks_list[0]
        count = len(remarks_list)

        text_rect = option.rect.adjusted(6, 0, -(self.button_width + 10), 0)
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(option.palette.color(QPalette.Text))

        elided_text = option.fontMetrics.elidedText(latest_remark, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        btn_rect = self.get_button_rect(option)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(btn_rect, 4, 4)
        painter.setPen(QColor("#1E5C99"))
        painter.drawText(btn_rect, Qt.AlignCenter, f"💬 {count}")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            btn_rect = self.get_button_rect(option)
            if btn_rect.contains(event.pos()):
                remarks_list = index.data(Qt.UserRole + 1)
                if remarks_list:
                    menu = QMenu()
                    for remark in remarks_list:
                        menu.addAction(remark)
                    menu.exec_(event.globalPos())
                return True
        return super().editorEvent(event, model, option, index)


class WorkItemsTableModel(QAbstractTableModel):
    """Fast tabular data model supporting chunked row insertions and tracking."""

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self.search_term = ""
        self.starred_codes = set()
        self.following_codes = set()
        self._headers = [
            "Track", "Code", "Acronym", "Name", "WG",
            "Latest WID", "Release", "Status", "Specs",
            "Start Date", "End Date", "Remarks"
        ]

    def set_tracking_sets(self, starred: set, following: set):
        self.starred_codes = {str(x).strip() for x in starred}
        self.following_codes = {str(x).strip() for x in following}

    def refresh_row_by_code(self, wi_code: str):
        target_code = str(wi_code).strip()
        for row_idx, r in enumerate(self._data):
            if str(r.get("code") or "").strip() == target_code:
                top_left = self.index(row_idx, 0)
                bottom_right = self.index(row_idx, len(self._headers) - 1)
                self.dataChanged.emit(top_left, bottom_right, [Qt.UserRole + 2, Qt.BackgroundRole, Qt.ToolTipRole])
                break

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]
        wi_code = str(row.get("code") or "").strip()

        if role == Qt.UserRole + 2 and col_name == "Track":
            return {
                "code": wi_code,
                "starred": wi_code in self.starred_codes,
                "following": wi_code in self.following_codes,
            }

        if role in (Qt.DisplayRole, Qt.UserRole):
            if col_name == "Track":
                return ""

            key_map = {
                "Code": "code", "Acronym": "acronym", "Name": "name",
                "WG": "wg_names", "Latest WID": "latest_wid", "Release": "release",
                "Specs": "spec_count", "Start Date": "start_date", "End Date": "end_date",
                "Remarks": "remarks"
            }
            if col_name == "Remarks" and role == Qt.DisplayRole:
                return ""

            if col_name == "Status" and role == Qt.DisplayRole:
                end_date = str(row.get("end_date") or "").strip()
                finished = bool(end_date and end_date < datetime.now().strftime("%Y-%m-%d"))
                return "Finished" if finished else "Active"

            if col_name == "Specs" and role == Qt.DisplayRole:
                cnt = row.get("spec_count", 0)
                return f"📑 {cnt}" if cnt > 0 else "-"

            val = row.get(key_map.get(col_name, ""), "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.UserRole + 1 and col_name == "Remarks":
            raw_remarks = row.get("remarks")
            if raw_remarks:
                bundled_remarks = raw_remarks.split("|||")
                parsed_remarks = []
                for item in bundled_remarks:
                    parts = item.split(":::", 1)
                    parsed_remarks.append((parts[0], parts[1]) if len(parts) == 2 else ("", item))
                parsed_remarks.sort(key=lambda x: x[0], reverse=True)
                return [item[1] for item in parsed_remarks]
            return []

        elif role == Qt.BackgroundRole:
            if col_name == "Acronym" and self.search_term:
                acronym = str(row.get("acronym") or "").strip().casefold()
                if acronym == self.search_term.casefold():
                    return QBrush(QColor("#DBEAFE"))

            # Highlight followed and starred rows subtly in Acronym and Track columns
            if col_name in ("Acronym", "Track"):
                if wi_code in self.following_codes:
                    return QBrush(QColor("#EFF6FF"))
                elif wi_code in self.starred_codes:
                    return QBrush(QColor("#FFFBEB"))

            if col_name == "Release" and str(row.get("release") or "").strip():
                return QBrush(QColor("#F1F5F9"))
            if col_name == "Status":
                end_date = str(row.get("end_date") or "").strip()
                finished = bool(end_date and end_date < datetime.now().strftime("%Y-%m-%d"))
                return QBrush(QColor("#F1F5F9" if finished else "#ECFDF5"))

        elif role == Qt.ForegroundRole:
            if col_name == "Acronym" and self.search_term and str(row.get("acronym") or "").strip().casefold() == self.search_term.casefold():
                return QBrush(QColor("#1E5C99"))
            if col_name == "Status":
                end_date = str(row.get("end_date") or "").strip()
                finished = bool(end_date and end_date < datetime.now().strftime("%Y-%m-%d"))
                return QBrush(QColor("#475569" if finished else "#065F46"))

        elif role == Qt.FontRole:
            if col_name in ["Acronym", "Release", "Status"]:
                font = QFont()
                font.setBold(True)
                return font

        elif role == Qt.ToolTipRole:
            if col_name == "Track":
                s_state = "⭐ Starred" if wi_code in self.starred_codes else "☆ Not Starred"
                f_state = "🎯 Following" if wi_code in self.following_codes else "🎯 Not Following"
                return (
                    f"<b>Tracking Status (WI #{wi_code})</b><br>"
                    f"Left (⭐): {s_state} (Personal Interest)<br>"
                    f"Right (🎯): {f_state} (Assigned Duty)<br><br>"
                    f"<i>Click left to toggle Star · Click right to toggle Follow</i>"
                )
            elif col_name == "Name":
                return f"<b>{row.get('acronym', '')}</b><br>{row.get('name', '')}"
            elif col_name == "Acronym":
                return row.get('name', '')
            elif col_name == "Specs":
                cnt = row.get("spec_count", 0)
                return f"Click to view all {cnt} linked specifications" if cnt > 0 else "No linked specifications"
            elif col_name == "Latest WID":
                wid = row.get("latest_wid", "")
                return f"Click to download/open latest Work Item Description ({wid})" if wid else ""
            elif col_name == "Remarks":
                raw = row.get("remarks", "")
                if raw:
                    entries = [r.split(":::")[-1] for r in raw.split("|||")]
                    return "<br>• " + "<br>• ".join(entries[:5])
            return None

        elif role == Qt.TextAlignmentRole:
            if col_name in ["Name", "Remarks"]:
                return Qt.AlignLeft | Qt.AlignVCenter
            return Qt.AlignCenter
        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        return None

    def get_row_data(self, row_idx: int) -> dict:
        if 0 <= row_idx < len(self._data):
            return self._data[row_idx]
        return {}

    def set_search_term(self, term: str):
        self.search_term = (term or "").strip()

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()

    def append_data(self, new_rows):
        if not new_rows:
            return
        start_row = len(self._data)
        end_row = start_row + len(new_rows) - 1
        self.beginInsertRows(QModelIndex(), start_row, end_row)
        self._data.extend(new_rows)
        self.endInsertRows()


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
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- HEADER & CONTROLS ---
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
        self.progress_bar.setFixedWidth(200)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #64748B; font-style: italic;")

        header_layout.addWidget(header_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.status_lbl)
        header_layout.addWidget(self.progress_bar)
        header_layout.addWidget(self.add_wi_btn)
        header_layout.addWidget(self.sync_tracked_btn)
        header_layout.addWidget(self.sync_btn)
        main_layout.addLayout(header_layout)

        # --- INLINE SEARCH & FILTER BAR ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("<b>🔍 Local Search:</b>"))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Code, Acronym, or Name...")
        self.search_input.setToolTip("Filter the table instantly by typing keywords.")
        self.search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.search_input)

        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(150)
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(130)
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.wg_combo)

        self.status_combo = QComboBox()
        self.status_combo.setToolTip("Filter by Work Item completion status (End Date)")
        self.status_combo.setMinimumWidth(150)
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

        # --- ROW 3: TRACKING FILTER CHIPS & RESULTS COUNTER ---
        chips_layout = QHBoxLayout()
        chips_layout.setSpacing(6)

        chips_lbl = QLabel("Filter:")
        chips_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        chips_layout.addWidget(chips_lbl)

        # Following Filter Chip (Assigned Duty)
        self.following_chip = QPushButton("🎯 Following (0)")
        self.following_chip.setCursor(Qt.PointingHandCursor)
        self.following_chip.setToolTip("Toggle view to show only Work Items you are assigned to follow.")
        self.following_chip.clicked.connect(self._toggle_following_filter)
        chips_layout.addWidget(self.following_chip)

        # Starred Filter Chip (Personal Interest)
        self.starred_chip = QPushButton("⭐ Starred (0)")
        self.starred_chip.setCursor(Qt.PointingHandCursor)
        self.starred_chip.setToolTip("Toggle view to show only your starred/favorite Work Items.")
        self.starred_chip.clicked.connect(self._toggle_starred_filter)
        chips_layout.addWidget(self.starred_chip)

        chips_layout.addStretch()

        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet("font-weight: bold; color: #64748B; margin-top: 2px;")
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
        self.table.setStyleSheet("""
            QTableView {
                border: 1px solid #CBD5E1;
                gridline-color: #F1F5F9;
                background-color: #FFFFFF;
            }
            QTableView::item:selected {
                background-color: #EBF3FC;
                color: #1E293B;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                padding: 4px;
                font-weight: bold;
                border: 1px solid #E2E8F0;
                color: #1E293B;
            }
        """)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.doubleClicked.connect(self._on_row_double_clicked)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)

        # Column 0: Track
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 62)
        self.tracking_delegate = TrackingDelegate(self.table)
        self.tracking_delegate.star_toggled.connect(self._toggle_star)
        self.tracking_delegate.follow_toggled.connect(self._toggle_follow)
        self.table.setItemDelegateForColumn(0, self.tracking_delegate)

        # Column 3: Name
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        # Column 5: Latest WID
        self.wid_delegate = WidDelegate(self.table)
        self.wid_delegate.link_clicked.connect(self._log_and_forward_action)
        self.table.setItemDelegateForColumn(5, self.wid_delegate)

        # Column 8: Specs
        self.specs_delegate = SpecsDelegate(self.table)
        self.specs_delegate.specs_clicked.connect(self._on_specs_clicked)
        self.table.setItemDelegateForColumn(8, self.specs_delegate)

        # Column 11: Remarks
        header.setSectionResizeMode(11, QHeaderView.Stretch)
        self.table.setItemDelegateForColumn(11, RemarksDelegate(self.table))

        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll)
        main_layout.addWidget(self.table)

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

        if self.starred_only:
            self.starred_chip.setStyleSheet("""
                QPushButton {
                    background-color: #FEF3C7; color: #92400E; border: 1px solid #F59E0B;
                    border-radius: 10px; font-weight: bold; padding: 3px 10px; font-size: 11px;
                }
                QPushButton:hover { background-color: #FDE68A; }
            """)
        else:
            self.starred_chip.setStyleSheet("""
                QPushButton {
                    background-color: #F0F4F8; color: #4B5563; border: 1px solid #D1D5DB;
                    border-radius: 10px; padding: 3px 10px; font-size: 11px;
                }
                QPushButton:hover { background-color: #E2E8F0; }
            """)

        if self.following_only:
            self.following_chip.setStyleSheet("""
                QPushButton {
                    background-color: #DBEAFE; color: #1E40AF; border: 1px solid #3B82F6;
                    border-radius: 10px; font-weight: bold; padding: 3px 10px; font-size: 11px;
                }
                QPushButton:hover { background-color: #BFDBFE; }
            """)
        else:
            self.following_chip.setStyleSheet("""
                QPushButton {
                    background-color: #F0F4F8; color: #4B5563; border: 1px solid #D1D5DB;
                    border-radius: 10px; padding: 3px 10px; font-size: 11px;
                }
                QPushButton:hover { background-color: #E2E8F0; }
            """)

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
        self._apply_checked_items(self.release_combo, [])
        self._apply_checked_items(self.wg_combo, [])
        self.status_combo.setCurrentIndex(0)
        self.starred_only = False
        self.following_only = False
        self._update_chip_styles()
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

    def refresh_table(self, reset_pagination: bool = True):
        self.save_filters_timer.start()

        if reset_pagination:
            self._current_offset = 0
            self._loaded_data = []

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()
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

    def _load_next_chunk(self):
        self._is_loading_chunk = True
        self._current_offset += self._chunk_size

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()
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

        filter_desc = []
        if self.following_only:
            filter_desc.append("🎯 Following")
        if self.starred_only:
            filter_desc.append("⭐ Starred")

        mode_str = f" [{' + '.join(filter_desc)}]" if filter_desc else ""
        prefix = f'{self._total_count} matches for "{query}"{mode_str}' if query else f"{self._total_count} Work Items{mode_str}"
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
