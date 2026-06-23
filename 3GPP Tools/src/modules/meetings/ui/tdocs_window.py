# --- File: modules/meetings/ui/tdocs_window.py ---
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QComboBox, QFrame)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel


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
# --- NEW: PROXY FILTER MODEL ---
# ==========================================
class TDocsFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.global_filter = ""
        self.type_filter = "All Types"
        self.status_filter = "All Statuses"

    def setGlobalFilter(self, text):
        self.global_filter = text.lower().strip()
        self.invalidateFilter()

    def setTypeFilter(self, text):
        self.type_filter = text
        self.invalidateFilter()

    def setStatusFilter(self, text):
        self.status_filter = text
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # 1. Type Filter (Column Index 3)
        if self.type_filter != "All Types":
            type_idx = model.index(source_row, 3, source_parent)
            if model.data(type_idx, Qt.DisplayRole) != self.type_filter:
                return False

        # 2. Status Filter (Column Index 8)
        if self.status_filter != "All Statuses":
            status_idx = model.index(source_row, 8, source_parent)
            if model.data(status_idx, Qt.DisplayRole) != self.status_filter:
                return False

        # 3. Global Search (Searches TDoc, Title, Source, Abstract, and Related)
        if self.global_filter:
            match_found = False
            for col in [0, 1, 2, 5, 9]:
                idx = model.index(source_row, col, source_parent)
                data = model.data(idx, Qt.DisplayRole)
                if data and self.global_filter in str(data).lower():
                    match_found = True
                    break
            if not match_found:
                return False

        return True


# ==========================================
# --- UPGRADED: TDOCS WINDOW ---
# ==========================================
class TDocsWindow(QWidget):
    def __init__(self, mtg_info: dict, tdocs_data: list):
        super().__init__()
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

        self.count_lbl = QLabel(f"Showing {len(tdocs_data)} of {len(tdocs_data)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.count_lbl)
        main_layout.addLayout(header_layout)

        # --- MODERN FILTER BAR ---
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; }
            QLabel { font-weight: bold; color: #555; border: none; }
            QLineEdit, QComboBox { 
                padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF;
            }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #0078D7; }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(15, 10, 15, 10)

        # 1. Global Search
        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search TDoc number, title, source, or abstract...")
        self.search_input.setMinimumWidth(300)
        self.search_input.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_input)

        # 2. Dynamic Type Dropdown
        filter_layout.addWidget(QLabel("Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItem("All Types")
        # Extract unique types dynamically
        unique_types = sorted(list(set(r.get("Type", "") for r in tdocs_data if r.get("Type"))))
        self.type_combo.addItems(unique_types)
        self.type_combo.currentTextChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        # 3. Dynamic Status Dropdown
        filter_layout.addWidget(QLabel("Status:"))
        self.status_combo = QComboBox()
        self.status_combo.addItem("All Statuses")
        # Extract unique statuses dynamically
        unique_statuses = sorted(list(set(r.get("TDoc Status", "") for r in tdocs_data if r.get("TDoc Status"))))
        self.status_combo.addItems(unique_statuses)
        self.status_combo.currentTextChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        main_layout.addWidget(filter_frame)

        # --- TABLE SETUP ---
        self.table = QTableView()
        self.model = TDocsTableModel(tdocs_data)

        # Wrap the model in our new Proxy!
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.table.setModel(self.proxy)

        # UI Styling & Sorting
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)  # Bonus feature: Click headers to sort!
        self.table.setStyleSheet("""
            QTableView { 
                gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF;
            }
            QHeaderView::section {
                background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0;
            }
        """)

        # Formatting Columns & Rows
        self.table.verticalHeader().setDefaultSectionSize(40)  # Breathing room
        self.table.resizeRowsToContents()

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)  # TDoc
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Title
        header.resizeSection(9, 150)  # Related TDocs

        main_layout.addWidget(self.table)

    # --- FILTER TRIGGERS ---
    def _on_search_changed(self, text):
        self.proxy.setGlobalFilter(text)

    def _on_type_changed(self, text):
        self.proxy.setTypeFilter(text)

    def _on_status_changed(self, text):
        self.proxy.setStatusFilter(text)

    def _update_count_label(self):
        visible = self.proxy.rowCount()
        total = self.model.rowCount()
        self.count_lbl.setText(f"Showing {visible} of {total} TDocs")