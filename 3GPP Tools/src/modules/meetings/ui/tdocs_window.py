# --- File: modules/meetings/ui/tdocs_window.py ---
import os
import webbrowser
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QComboBox, QFrame,
                             QPushButton, QMessageBox)
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QEvent, pyqtSignal


# ==========================================
# --- UPGRADED: BULLETPROOF MULTI-SELECT ---
# ==========================================
class CheckableComboBox(QComboBox):
    selectionChanged = pyqtSignal(list)

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.title = title
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.setModel(QStandardItemModel(self))

        # FIXED: Catch ALL state changes via the model, guaranteeing it works with Mouse & Keyboard
        self.model().itemChanged.connect(self._on_item_changed)
        self.view().viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj == self.view().viewport() and event.type() == QEvent.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            if index.isValid():
                item = self.model().itemFromIndex(index)
                if item:
                    # FIXED: Safe truthy check bypasses the PyQt Enum vs Integer bug!
                    new_state = Qt.Unchecked if item.checkState() else Qt.Checked
                    item.setCheckState(new_state)
            return True  # Consume event to prevent popup from closing
        return super().eventFilter(obj, event)

    def addItems(self, items):
        # Block signals briefly so it doesn't spam filters while building the dropdown
        self.model().blockSignals(True)
        for text in items:
            display_text = str(text) if text else "(Empty)"
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            item.setData(str(text), Qt.UserRole)
            self.model().appendRow(item)
        self.model().blockSignals(False)
        self.updateText()

    def getCheckedItems(self):
        checked = []
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            # Evaluates to True if it is Checked (2) or PartiallyChecked (1)
            if item.checkState():
                checked.append(item.data(Qt.UserRole))
        return checked

    def _on_item_changed(self, item):
        self.updateText()
        self.selectionChanged.emit(self.getCheckedItems())

    def updateText(self, *args):
        checked = self.getCheckedItems()
        total = self.model().rowCount()

        if total == 0:
            self.lineEdit().setText(f"{self.title}: None")
        elif len(checked) == total:
            self.lineEdit().setText(f"{self.title}: All")
        elif len(checked) == 1:
            display = checked[0] if checked[0] else "(Empty)"
            self.lineEdit().setText(f"{self.title}: {display}")
        else:
            self.lineEdit().setText(f"{self.title}: {len(checked)} selected")


class TDocsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = [
            "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "Agenda Item", "TDoc Status", "Related TDocs"
        ]

    def _format_related_tdocs(self, row_data: dict) -> str:
        parts = []
        if row_data.get("Is revision of"): parts.append(f"⬅️ Rev of: {row_data['Is revision of']}")
        if row_data.get("Revised to"): parts.append(f"➡️ Rev to: {row_data['Revised to']}")
        if row_data.get("Original LS"): parts.append(f"✉️ Orig LS: {row_data['Original LS']}")
        if row_data.get("Reply in"): parts.append(f"↩️ Reply: {row_data['Reply in']}")
        return "\n".join(parts)

    def data(self, index, role):
        if not index.isValid(): return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if role == Qt.DisplayRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row)
            return row.get(col_name, "")

        elif role == Qt.TextAlignmentRole:
            if col_name in ["TDoc", "Type", "For", "Agenda Item", "TDoc Status"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignTop
        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None


# ==========================================
# --- UPGRADED: PROXY FILTER MODEL ---
# ==========================================
class TDocsFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.global_filter = ""
        # FIXED: Using Python Sets for ultra-fast, robust lookups
        self.type_filters = set()
        self.status_filters = set()
        self.ai_filters = set()

    def setGlobalFilter(self, text):
        self.global_filter = str(text).lower().strip()
        self.invalidateFilter()

    def setTypeFilters(self, types):
        self.type_filters = set(types)
        self.invalidateFilter()

    def setStatusFilters(self, statuses):
        self.status_filters = set(statuses)
        self.invalidateFilter()

    def setAIFilters(self, ais):
        self.ai_filters = set(ais)
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # FIXED: Enforce absolute Python strings, dodging QVariant mismatch bugs

        # 1. Type Filter
        type_data = str(model.data(model.index(source_row, 3, source_parent), Qt.DisplayRole) or "")
        if type_data not in self.type_filters: return False

        # 2. Agenda Item Filter
        ai_data = str(model.data(model.index(source_row, 7, source_parent), Qt.DisplayRole) or "")
        if ai_data not in self.ai_filters: return False

        # 3. Status Filter
        status_data = str(model.data(model.index(source_row, 8, source_parent), Qt.DisplayRole) or "")
        if status_data not in self.status_filters: return False

        # 4. Global Search
        if self.global_filter:
            match_found = False
            for col in [0, 1, 2, 5, 9]:
                data = model.data(model.index(source_row, col, source_parent), Qt.DisplayRole)
                if data and self.global_filter in str(data).lower():
                    match_found = True
                    break
            if not match_found: return False

        return True


# ==========================================
# --- TDOCS WINDOW ---
# ==========================================
class TDocsWindow(QWidget):
    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.filepath = filepath
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')}"
        self.setWindowTitle(title)
        self.resize(1400, 750)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- HEADER & COUNT ---
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 18px; color: #333;")

        self.excel_btn = QPushButton("📗 Open in Excel")
        self.excel_btn.setCursor(Qt.PointingHandCursor)
        self.excel_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #0C6B0C; background-color: #E6F4E6; border: 1px solid #A3DDA3;
            }
            QPushButton:hover { background-color: #D1EED1; border: 1px solid #0C6B0C; }
        """)
        self.excel_btn.clicked.connect(self._open_excel)

        self.count_lbl = QLabel(f"Showing {len(tdocs_data)} of {len(tdocs_data)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.excel_btn)
        header_layout.addSpacing(15)
        header_layout.addWidget(self.count_lbl)
        main_layout.addLayout(header_layout)

        # --- MODERN FILTER BAR ---
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; }
            QLabel { font-weight: bold; color: #555; border: none; }
            QLineEdit, QComboBox { padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; }
            QLineEdit:focus { border: 1px solid #0078D7; }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(15, 10, 15, 10)

        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search TDoc number, title, source, or abstract...")
        self.search_input.setMinimumWidth(250)
        self.search_input.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_input)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMinimumWidth(150)
        unique_types = sorted(list(set(str(r.get("Type", "")) for r in tdocs_data)))
        self.type_combo.addItems(unique_types)
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMinimumWidth(150)
        unique_ais = sorted(list(set(str(r.get("Agenda Item", "")) for r in tdocs_data)))
        self.ai_combo.addItems(unique_ais)
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("Status")
        self.status_combo.setMinimumWidth(150)
        unique_statuses = sorted(list(set(str(r.get("TDoc Status", "")) for r in tdocs_data)))
        self.status_combo.addItems(unique_statuses)
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        main_layout.addWidget(filter_frame)

        # --- TABLE SETUP ---
        self.table = QTableView()
        self.model = TDocsTableModel(tdocs_data)

        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        # Initialize Proxy with all data selected
        self.proxy.setTypeFilters(unique_types)
        self.proxy.setAIFilters(unique_ais)
        self.proxy.setStatusFilters(unique_statuses)

        self.table.setModel(self.proxy)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet("""
            QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; }
            QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }
        """)

        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.resizeRowsToContents()

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.resizeSection(9, 150)

        main_layout.addWidget(self.table)

    # --- ACTIONS & TRIGGERS ---
    def _open_excel(self):
        try:
            if hasattr(os, 'startfile'):
                os.startfile(self.filepath)
            else:
                webbrowser.open(f"file:///{self.filepath}")
        except Exception as e:
            QMessageBox.warning(self, "Open Error", f"Could not open the Excel file:\n{e}")

    def _on_search_changed(self, text):
        self.proxy.setGlobalFilter(text)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)

    def _update_count_label(self):
        visible = self.proxy.rowCount()
        total = self.model.rowCount()
        self.count_lbl.setText(f"Showing {visible} of {total} TDocs")