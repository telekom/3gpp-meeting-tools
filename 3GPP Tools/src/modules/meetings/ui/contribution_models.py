# --- File: src/modules/meetings/ui/contribution_models.py ---
"""
Table model for contribution audit search results.
Provides restrained semantic highlighting for status cells, metadata badges,
and filter match explanations.
"""
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor, QFont


class ContributionResultsTableModel(QAbstractTableModel):
    """Contribution results with restrained semantic highlighting for status cells."""

    COLUMNS = [
        "WG", "Meeting", "TDoc", "Title", "Source", "Matched Company",
        "Type", "For", "Agenda Item", "TDoc Status", "Related WIs", "Abstract"
    ]

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self.COLUMNS)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.COLUMNS[section]
        return None

    def data(self, index, role):
        if not index.isValid() or not (0 <= index.row() < len(self._data)):
            return None

        row_dict = self._data[index.row()]
        col_name = self.COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            return str(row_dict.get(col_name, ""))

        # Compact metadata badges: restrained fills make short Type/For values
        # scannable without coloring the complete row.
        if col_name in ("Type", "For"):
            value = str(row_dict.get(col_name, "") or "").strip()
            if value:
                if role == Qt.BackgroundRole:
                    return QBrush(QColor("#EBF3FC" if col_name == "Type" else "#F1F5F9"))
                if role == Qt.ForegroundRole:
                    return QBrush(QColor("#1E5C99" if col_name == "Type" else "#475569"))
                if role == Qt.FontRole:
                    font = QFont()
                    font.setBold(True)
                    return font

        # Explain WI/AI OR searches visually. Only the matching facet cells are
        # tinted, based on the worker's explicit match annotation.
        match_origin = str(row_dict.get("_Filter Match", ""))
        if col_name == "Related WIs" and "WI" in match_origin:
            if role == Qt.BackgroundRole:
                return QBrush(QColor("#EFF6FF"))
            if role == Qt.ForegroundRole:
                return QBrush(QColor("#1D4ED8"))
            if role == Qt.FontRole:
                font = QFont()
                font.setBold(True)
                return font
        if col_name == "Agenda Item" and "AI" in match_origin:
            if role == Qt.BackgroundRole:
                return QBrush(QColor("#F5F3FF"))
            if role == Qt.ForegroundRole:
                return QBrush(QColor("#6D28D9"))
            if role == Qt.FontRole:
                font = QFont()
                font.setBold(True)
                return font

        if col_name == "TDoc Status":
            status = str(row_dict.get(col_name, "") or "").strip().lower()
            if role == Qt.BackgroundRole:
                if "agreed" in status or "approved" in status:
                    return QBrush(QColor("#E6F4EA"))
                if "noted" in status or "postponed" in status:
                    return QBrush(QColor("#FEF7E0"))
                if any(term in status for term in ("withdrawn", "rejected", "not agreed", "not-agreed")):
                    return QBrush(QColor("#FCE8E6"))
                if status:
                    return QBrush(QColor("#F1F3F4"))
            if role == Qt.ForegroundRole:
                if "agreed" in status or "approved" in status:
                    return QBrush(QColor("#137333"))
                if "noted" in status or "postponed" in status:
                    return QBrush(QColor("#B06000"))
                if any(term in status for term in ("withdrawn", "rejected", "not agreed", "not-agreed")):
                    return QBrush(QColor("#C5221F"))
                if status:
                    return QBrush(QColor("#3C4043"))
            if role == Qt.FontRole and status:
                font = QFont()
                font.setBold(True)
                return font

        if role == Qt.TextAlignmentRole:
            if col_name in ["WG", "Meeting", "TDoc", "Type", "For", "Agenda Item", "TDoc Status", "Related WIs"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.ToolTipRole:
            val = str(row_dict.get(col_name, "")).strip()
            match_origin = str(row_dict.get("_Filter Match", ""))
            if col_name == "Related WIs" and "WI" in match_origin:
                return f"Matched the active WI filter.\n{val}" if val else "Matched the active WI filter."
            if col_name == "Agenda Item" and "AI" in match_origin:
                return f"Matched the active Agenda Item filter.\n{val}" if val else "Matched the active Agenda Item filter."
            if col_name == "Abstract" and val:
                return f"<div style='width: 450px; white-space: pre-wrap;'>{val}</div>"
            elif col_name in ["Title", "Source"] and len(val) > 40:
                return f"<div style='width: 400px; white-space: pre-wrap;'>{val}</div>"
            return None

        if role == Qt.UserRole:
            return row_dict

        return None

    def update_data(self, new_data: list):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()
