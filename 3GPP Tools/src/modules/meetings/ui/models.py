from PyQt5.QtCore import QAbstractTableModel, Qt, QModelIndex

from modules.meetings.ui.dialogs import _format_meeting_info


# ==========================================
# --- TABLE MODEL ---
# ==========================================

class MeetingsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["", "WG", "Meeting", "Name", "Location", "Start Date", "End Date"]

    def data(self, index, role):
        if not index.isValid(): return None
        row_data = self._data[index.row()]

        if role == Qt.DisplayRole:
            col = index.column()
            if col == 0:
                return ""
            elif col == 1:
                return row_data.get("wg_name", "")
            elif col == 2:
                return row_data.get("meeting_number", "")
            elif col == 3:
                return row_data.get("name", "")
            elif col == 4:
                return row_data.get("location", "")
            elif col == 5:
                return row_data.get("start_date", "")
            elif col == 6:
                return row_data.get("end_date", "")

        elif role == Qt.TextAlignmentRole:
            if index.column() in [0, 1, 2, 5, 6]: return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter
        elif role == Qt.UserRole:
            return row_data
        elif role == Qt.ToolTipRole:
            return _format_meeting_info(row_data)

        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()
