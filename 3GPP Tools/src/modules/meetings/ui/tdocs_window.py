# --- File: modules/meetings/ui/tdocs_window.py ---
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTableView, QHeaderView, QLabel
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex


class TDocsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = [
            "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "Agenda Item", "TDoc Status", "Related TDocs"
        ]

    def _format_related_tdocs(self, row_data: dict) -> str:
        """Combines related fields into a single, visually identifiable multi-line string."""
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


class TDocsWindow(QWidget):
    def __init__(self, mtg_info: dict, tdocs_data: list):
        super().__init__()
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')}"
        self.setWindowTitle(title)
        self.resize(1400, 700)

        layout = QVBoxLayout(self)

        info_lbl = QLabel(f"<b>{title}</b> | {len(tdocs_data)} TDocs found")
        info_lbl.setStyleSheet("font-size: 16px; margin-bottom: 5px;")
        layout.addWidget(info_lbl)

        self.table = QTableView()
        self.model = TDocsTableModel(tdocs_data)
        self.table.setModel(self.model)

        # UI Styling
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableView { gridline-color: #f0f0f0; }")

        # Formatting Columns & Rows
        self.table.resizeRowsToContents()  # Crucial for multi-line "Related" column
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)  # TDoc
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Title
        header.resizeSection(9, 150)  # Related TDocs

        layout.addWidget(self.table)