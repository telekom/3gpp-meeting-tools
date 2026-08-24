import webbrowser
from pathlib import Path
import logging

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer, QEvent, QRect, pyqtSignal
from PyQt5.QtGui import QColor, QPalette
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QTableView, QHeaderView, QPushButton, QProgressBar,
                             QMessageBox, QLineEdit, QMenu, QStyle, QApplication,
                             QStyledItemDelegate, QComboBox)

from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import WorkItemsScraperThread, TargetedWIScraperThread
from modules.work_items.core.wi_settings import WorkItemsSettings


class WidDelegate(QStyledItemDelegate):
    """Custom delegate to render the Latest WID as a clickable hyperlink."""
    link_clicked = pyqtSignal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)

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
            painter.setPen(QColor("#0078D7"))

        painter.drawText(option.rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text:
                logging.info(f"🚀 Emitting link_clicked signal for '{text}'...")
                self.link_clicked.emit(text, "open_doc")
                return True
            else:
                logging.warning("⚠️ Clicked, but no text found in cell.")
        return super().editorEvent(event, model, option, index)


class RemarksDelegate(QStyledItemDelegate):
    """Custom delegate to draw the latest remark text and a clickable '💬' history button."""
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

        text_rect = option.rect.adjusted(5, 0, -(self.button_width + 10), 0)
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(option.palette.color(QPalette.Text))

        elided_text = option.fontMetrics.elidedText(latest_remark, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        btn_rect = self.get_button_rect(option)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#E1F0FF"))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(btn_rect, 4, 4)
        painter.setPen(QColor("#0078D7"))
        painter.drawText(btn_rect, Qt.AlignCenter, f"💬 {count}")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            btn_rect = self.get_button_rect(option)
            if btn_rect.contains(event.pos()):
                remarks_list = index.data(Qt.UserRole + 1)
                if remarks_list:
                    menu = QMenu()
                    menu.setStyleSheet("""
                        QMenu { background-color: #FAFAFA; border: 1px solid #CCC; } 
                        QMenu::item { padding: 5px 20px 5px 15px; color: #333333; } 
                        QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                    """)
                    for remark in remarks_list:
                        menu.addAction(remark)
                    menu.exec_(event.globalPos())
                return True
        return super().editorEvent(event, model, option, index)


class WorkItemsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["Code", "Acronym", "Name", "Latest WID", "Release", "Start Date", "End Date", "Remarks"]

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if role == Qt.DisplayRole or role == Qt.UserRole:
            key_map = {
                "Code": "code", "Acronym": "acronym", "Name": "name",
                "Latest WID": "latest_wid", "Release": "release",
                "Start Date": "start_date", "End Date": "end_date",
                "Remarks": "remarks"
            }
            if col_name == "Remarks" and role == Qt.DisplayRole:
                return ""

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

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()


class WorkItemsTab(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db_path = db_path
        self.db = WorkItemsDatabase(db_path)
        self.settings = WorkItemsSettings()

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(400)
        self.search_timer.timeout.connect(self.refresh_table)

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        self._populate_filters()
        self._load_filters()
        self.refresh_table()

    def _save_filters(self):
        """Saves the current UI filter state including completion status."""
        filters = {
            "search": self.search_input.text().strip(),
            "releases": self.release_combo.getCheckedItems(),
            "wgs": self.wg_combo.getCheckedItems(),
            "status": self.status_combo.currentData() or "all"
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        """Restores the UI filter state from JSON on startup."""
        filters = self.settings.get_filters()
        if not filters:
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

    def _apply_checked_items(self, combo, items_to_check):
        model = combo.model()
        items_to_check_stripped = [str(x).strip() for x in items_to_check]

        for row in range(model.rowCount()):
            if hasattr(model, 'item'):
                item = model.item(row)
                if item.text().strip() in items_to_check_stripped:
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
            else:
                index = model.index(row, 0)
                text = str(model.data(index, Qt.DisplayRole)).strip()
                state = Qt.Checked if text in items_to_check_stripped else Qt.Unchecked
                model.setData(index, state, Qt.CheckStateRole)

        if hasattr(combo, 'updateText'):
            combo.updateText()
        elif hasattr(combo, 'repaint'):
            combo.repaint()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # --- HEADER & CONTROLS ---
        header_layout = QHBoxLayout()
        header_lbl = QLabel("<b>📋 3GPP Work Items (WIs)</b>")
        header_lbl.setStyleSheet("font-size: 16px; color: #333;")

        self.sync_btn = QPushButton("🔄 Sync 3GPP WIs")
        self.sync_btn.setStyleSheet("""
            QPushButton { font-weight: bold; background-color: #0078D7; color: white; padding: 5px 15px; border-radius: 4px; }
            QPushButton:hover { background-color: #005A9E; }
            QPushButton:disabled { background-color: #A0C0E0; }
        """)
        self.sync_btn.setToolTip("Click to download and synchronize 3GPP Work Items in parallel from the server.")
        self.sync_btn.clicked.connect(self._start_sync)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #666; font-style: italic;")

        header_layout.addWidget(header_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.status_lbl)
        header_layout.addWidget(self.progress_bar)
        header_layout.addWidget(self.sync_btn)
        main_layout.addLayout(header_layout)

        # --- INLINE SEARCH & FILTER BAR ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("<b>🔍 Local Search:</b>"))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Code, Acronym, or Name...")
        self.search_input.setToolTip("Filter the table instantly by typing keywords.")
        self.search_input.textChanged.connect(lambda text: self.search_timer.start())
        search_layout.addWidget(self.search_input)

        # Multi-select combos
        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(160)
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(140)
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.wg_combo)

        # Status Filter Combo
        self.status_combo = QComboBox()
        self.status_combo.setToolTip("Filter by Work Item completion status (End Date)")
        self.status_combo.setMinimumWidth(150)
        self.status_combo.addItem("All Statuses", "all")
        self.status_combo.addItem("🟢 Active / Ongoing", "active")
        self.status_combo.addItem("🏁 Finished", "finished")
        self.status_combo.currentIndexChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.status_combo)

        main_layout.addLayout(search_layout)

        # --- RESULTS COUNTER ---
        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet("font-weight: bold; color: #555555; margin-top: 5px;")
        count_layout = QHBoxLayout()
        count_layout.addStretch()
        count_layout.addWidget(self.count_label)
        main_layout.addLayout(count_layout)

        # --- TABLE VIEW ---
        self.table = QTableView()
        self.table_model = WorkItemsTableModel()
        self.table.setModel(self.table_model)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(
            "QTableView { border: 1px solid #dcdcdc; gridline-color: #f0f0f0; }"
            "QTableView::item:selected { background-color: #cce8ff; color: #000; }"
        )

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)

        self.wid_delegate = WidDelegate(self.table)
        self.wid_delegate.link_clicked.connect(self._log_and_forward_action)
        self.wid_delegate.link_clicked.connect(self.global_action_requested.emit)
        self.table.setItemDelegateForColumn(3, self.wid_delegate)

        header.setSectionResizeMode(7, QHeaderView.Stretch)
        self.table.setItemDelegateForColumn(7, RemarksDelegate(self.table))

        main_layout.addWidget(self.table)

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

    def refresh_table(self):
        self.save_filters_timer.start()

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()
        selected_status = self.status_combo.currentData() or "all"

        data = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status
        )

        self.table_model.update_data(data)
        count = len(data)
        self.count_label.setText(f"Showing {count} Work Items")

    def _start_sync(self):
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Syncing...")
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
        self.sync_btn.setEnabled(True)
        self.sync_btn.setText("🔄 Sync 3GPP WIs")
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("")

        self._populate_filters()
        self.refresh_table()

        if success:
            QMessageBox.information(self, "Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Sync Failed", msg)

    def _show_context_menu(self, position):
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #FAFAFA; border: 1px solid #CCC; } 
            QMenu::item { padding: 5px 20px 5px 15px; color: #333333; } 
            QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
            QMenu::item:disabled { color: #AAAAAA; } 
        """)

        len_indexes = len(selected_indexes)
        wi_code_list = [
            self.table_model.data(self.table_model.index(e.row(), 0), Qt.DisplayRole)
            for e in selected_indexes
        ]
        wi_code_list = [code for code in wi_code_list if code]

        if len_indexes == 1:
            wi_code = wi_code_list[0]
            if not wi_code:
                return

            wi_page_action = menu.addAction("🌐 Open WI Page")
            specs_action = menu.addAction("📂 Specifications Resulting from this WI")
            crs_action = menu.addAction("📄 CRs Related to this WI")
            update_action = menu.addAction("🔄 Update WI")
            delete_action = menu.addAction("🗑️ Delete this Work Item")
        else:
            wi_page_action = None
            specs_action = None
            crs_action = None
            wi_code = None
            update_action = menu.addAction(f"🔄 Update WIs ({len_indexes} WIs)")
            delete_action = menu.addAction(f"🗑️ Delete selected Work Items ({len_indexes} WIs)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if len_indexes == 1:
            if action == wi_page_action:
                url = f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
                webbrowser.open(url)
            elif action == specs_action:
                url = f"https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={wi_code}"
                webbrowser.open(url)
            elif action == crs_action:
                url = f"https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem={wi_code}"
                webbrowser.open(url)
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self,
                    "Confirm Deletion",
                    f"Are you sure you want to delete Work Item '{wi_code}' from the database?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_item(wi_code)
                    self.refresh_table()
            elif action == update_action:
                self._start_targeted_sync([wi_code])
        else:
            if action == delete_action:
                confirm = QMessageBox.question(
                    self,
                    "Confirm Batch Deletion",
                    f"Are you sure you want to delete {len(wi_code_list)} selected Work Items from the database?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_items(wi_code_list)
                    self.refresh_table()
            elif action == update_action:
                self._start_targeted_sync(wi_code_list)

    def _start_targeted_sync(self, wi_codes: list):
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Updating...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.targeted_thread = TargetedWIScraperThread(self.db_path, wi_codes, self)
        self.targeted_thread.progress.connect(self._update_progress)
        self.targeted_thread.finished_sync.connect(self._on_sync_finished)
        self.targeted_thread.start()