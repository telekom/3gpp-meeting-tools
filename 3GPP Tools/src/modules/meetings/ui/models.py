# --- File: modules/meetings/ui/models.py ---
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QDate
from modules.meetings.ui.dialogs import _format_meeting_info
from core.ui.ui_components import WG_DESCRIPTIONS


MONTH_NAMES = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def format_meeting_dates(start_str: str, end_str: str) -> str:
    """Consolidates start and end dates into a concise human-readable range (English month names)."""
    if not start_str and not end_str:
        return "—"
    if not end_str:
        try:
            d = QDate.fromString(start_str, "yyyy-MM-dd")
            if d.isValid():
                return f"{d.day()} {MONTH_NAMES[d.month()]} {d.year()}"
        except Exception:
            pass
        return start_str
    if not start_str:
        try:
            d = QDate.fromString(end_str, "yyyy-MM-dd")
            if d.isValid():
                return f"{d.day()} {MONTH_NAMES[d.month()]} {d.year()}"
        except Exception:
            pass
        return end_str

    try:
        d1 = QDate.fromString(start_str, "yyyy-MM-dd")
        d2 = QDate.fromString(end_str, "yyyy-MM-dd")
        if d1.isValid() and d2.isValid():
            m1 = MONTH_NAMES[d1.month()]
            m2 = MONTH_NAMES[d2.month()]
            if d1.year() == d2.year():
                if d1.month() == d2.month():
                    return f"{d1.day()} – {d2.day()} {m1} {d1.year()}"
                else:
                    return f"{d1.day()} {m1} – {d2.day()} {m2} {d1.year()}"
            else:
                return f"{d1.day()} {m1} {d1.year()} – {d2.day()} {m2} {d2.year()}"
    except Exception:
        pass

    return f"{start_str} – {end_str}"


class MeetingsTableModel(QAbstractTableModel):
    HEADER_TOOLTIPS = [
        "Click the ⋮ button on any row for meeting actions (FTP, Portal, cache folder, sync, delete)",
        "TDocs list status: click to download or open the meeting Excel list",
        "Working Group (e.g., SA2, RAN1)",
        "Meeting Number (e.g., 160, 112-e)",
        "Meeting Location (City, Country or eMeeting)",
        "Official meeting dates (Start – End)",
        "Allocated TDoc range in official archive"
    ]

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["", "Action", "WG", "Meeting", "Location", "Dates", "Allocated TDocs"]

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
                loc = row_data.get("location", "")
                is_elec = bool(row_data.get("is_electronic", 0)) or "online" in (loc or "").lower()
                if is_elec:
                    return f"🌐 {loc}" if loc else "🌐 Online"
                return f"📍 {loc}" if loc else ""
            elif col == 5:
                return format_meeting_dates(row_data.get("start_date", ""), row_data.get("end_date", ""))
            elif col == 6:
                first = row_data.get("first_tdoc", "")
                last = row_data.get("last_tdoc", "")
                if first and last:
                    return f"{first} → {last}"
                elif first or last:
                    return first or last
                return "—"

        elif role == Qt.TextAlignmentRole:
            if col in [0, 1, 2, 3, 5, 6]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.UserRole:
            return row_data

        elif role == Qt.ToolTipRole:
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
            elif col == 2:
                wg_name = row_data.get("wg_name", "").upper()
                desc = WG_DESCRIPTIONS.get(wg_name, "3GPP Working Group")
                return f"{wg_name}: {desc}"
            elif col == 5:
                start_str = row_data.get("start_date", "")
                end_str = row_data.get("end_date", "")
                today = QDate.currentDate()
                status_prefix = ""
                try:
                    d1 = QDate.fromString(start_str, "yyyy-MM-dd")
                    d2 = QDate.fromString(end_str, "yyyy-MM-dd")
                    if d1.isValid() and d2.isValid():
                        if d1 <= today <= d2:
                            status_prefix = "🟢 ONGOING MEETING\n"
                        elif today < d1:
                            days = today.daysTo(d1)
                            status_prefix = f"🔵 UPCOMING MEETING (in {days} days)\n"
                        else:
                            status_prefix = "⚪ PAST MEETING\n"
                except Exception:
                    pass
                return f"{status_prefix}{_format_meeting_info(row_data)}"

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
