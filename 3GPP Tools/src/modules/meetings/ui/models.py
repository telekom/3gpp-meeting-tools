# --- File: modules/meetings/ui/models.py ---
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from modules.meetings.ui.dialogs import _format_meeting_info


class MeetingsTableModel(QAbstractTableModel):
    HEADER_TOOLTIPS = [
        "Click the ⋮ button on any row for meeting actions (FTP, Portal, cache folder, sync, delete)",
        "TDocs list status: click to download or open the meeting Excel list",
        "Working Group (e.g., SA2, RAN1)",
        "Meeting Number (e.g., 160, 112-e)",
        "Meeting Location (City, Country)",
        "Official meeting start date",
        "Official meeting end date",
        "First allocated TDoc in Docs/ archive",
        "Last allocated TDoc in Docs/ archive"
    ]

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["", "📄", "WG", "Meeting", "Location", "Start Date", "End Date", "First TDoc", "Last TDoc"]

    def data(self, index, role):
        if not index.isValid():
            return None
        row_data = self._data[index.row()]
        col = index.column()

        if role == Qt.DisplayRole:
            if col in [0, 1]:
                return ""
            elif col == 2:
                return row_data.get("wg_name", "")
            elif col == 3:
                return row_data.get("meeting_number", "")
            elif col == 4:
                return row_data.get("location", "")
            elif col == 5:
                return row_data.get("start_date", "")
            elif col == 6:
                return row_data.get("end_date", "")
            elif col == 7:
                return row_data.get("first_tdoc", "")
            elif col == 8:
                return row_data.get("last_tdoc", "")

        elif role == Qt.TextAlignmentRole:
            if col in [0, 1, 2, 3, 5, 6, 7, 8]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.UserRole:
            return row_data

        elif role == Qt.ToolTipRole:
            # Dedicated interactive tooltips for action columns
            if col == 0:
                return "Click to open meeting actions menu (Portal links, FTP folders, cache folder, sync, delete)."
            elif col == 1:
                status = row_data.get('tdoc_btn_status', 'na')
                if status == 'open':
                    return "Local TDocs list is available. Click to open the TDocs table viewer."
                elif status == 'get':
                    return "TDocs list not cached yet. Click to download from 3GPP FTP and open."
                elif status == 'fetching':
                    return "Downloading TDocs list from 3GPP FTP..."
                return "No 3GPP Portal ID available for automatic TDoc list download."
            # Data columns retain the complete formatted meeting information card
            return _format_meeting_info(row_data)

        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                return self._headers[section]
            elif role == Qt.ToolTipRole and section < len(self.HEADER_TOOLTIPS):
                return self.HEADER_TOOLTIPS[section]
        return None

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()