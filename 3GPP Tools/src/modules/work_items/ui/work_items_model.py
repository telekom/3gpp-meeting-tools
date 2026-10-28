"""Table model for the Work Items view."""

from datetime import datetime
import logging
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QBrush, QColor, QFont

from core.ui.ui_components import WG_DESCRIPTIONS

logger = logging.getLogger(__name__)


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
            "Timeline", "Remarks"
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
                "Specs": "spec_count", "Remarks": "remarks"
            }
            if col_name == "Remarks" and role == Qt.DisplayRole:
                return ""

            if col_name == "Timeline" and role == Qt.DisplayRole:
                start = str(row.get("start_date") or "").strip()
                end = str(row.get("end_date") or "").strip()
                s_ym = start[:7] if len(start) >= 7 else ""
                e_ym = end[:7] if len(end) >= 7 else ""
                if s_ym and e_ym:
                    return f"{s_ym} → {e_ym}"
                elif e_ym:
                    return f"End: {e_ym}"
                elif s_ym:
                    return f"Start: {s_ym}"
                return "-"

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

            # Unified soft row highlight for tracked items
            if wi_code in self.following_codes:
                return QBrush(QColor("#EFF6FF"))
            elif wi_code in self.starred_codes:
                return QBrush(QColor("#FFFBEB"))

        elif role == Qt.ForegroundRole:
            if col_name == "Acronym" and self.search_term and str(row.get("acronym") or "").strip().casefold() == self.search_term.casefold():
                return QBrush(QColor("#1E5C99"))
            elif col_name == "Code":
                return QBrush(QColor("#475569"))

        elif role == Qt.FontRole:
            if col_name in ["Acronym", "Latest WID"]:
                font = QFont()
                font.setBold(True)
                return font
            elif col_name == "Code":
                font = QFont("Consolas", 9)
                font.setBold(True)
                return font

        elif role == Qt.ToolTipRole:
            if col_name == "Track":
                s_state = "⭐ Starred" if wi_code in self.starred_codes else "☆ Not Starred"
                f_state = "🎯 Following" if wi_code in self.following_codes else "◎ Not Following"
                return (
                    f"<b>Tracking Status (WI #{wi_code})</b><br>"
                    f"Left (⭐): {s_state} (Personal Interest)<br>"
                    f"Right (🎯): {f_state} (Assigned Duty)<br><br>"
                    f"<i>Click left to toggle Star · Click right to toggle Follow</i>"
                )
            elif col_name == "Code":
                return f"<b>Work Item #{wi_code}</b><br>Double-click row to view full details, history, and linked specifications."
            elif col_name == "Name":
                return f"<b>{row.get('acronym', '')}</b><br>{row.get('name', '')}"
            elif col_name == "Acronym":
                return f"<b>{row.get('acronym', '')}</b><br>{row.get('name', '')}"
            elif col_name == "WG":
                raw_wgs = [w.strip() for w in str(row.get('wg_names') or '').split(',') if w.strip()]
                if raw_wgs:
                    lines = [f"• <b>{w}</b>: {WG_DESCRIPTIONS.get(w.upper(), 'Working Group')}" for w in raw_wgs]
                    return "<b>Working Groups:</b><br>" + "<br>".join(lines)
                return None
            elif col_name == "Specs":
                cnt = row.get("spec_count", 0)
                return f"Click to view all {cnt} linked specifications" if cnt > 0 else "No linked specifications"
            elif col_name == "Latest WID":
                wid = row.get("latest_wid", "")
                return f"Click to download/open latest Work Item Description ({wid})" if wid else ""
            elif col_name == "Release":
                rel = row.get("release", "")
                return f"<b>Release:</b> {rel}" if rel else None
            elif col_name == "Status":
                end_date = str(row.get("end_date") or "").strip()
                finished = bool(end_date and end_date < datetime.now().strftime("%Y-%m-%d"))
                return "<b>Status: Finished</b><br>Target completion date has passed." if finished else "<b>Status: Active / Ongoing</b><br>Normative work is actively ongoing."
            elif col_name == "Timeline":
                start = str(row.get("start_date") or "").strip() or "N/A"
                end = str(row.get("end_date") or "").strip() or "N/A"
                return f"<b>Work Item Timeline</b><br>• Start Date: {start}<br>• Target End Date: {end}"
            elif col_name == "Remarks":
                raw = row.get("remarks", "")
                if raw:
                    entries = [r.split(":::")[-1] for r in raw.split("|||")]
                    return "<b>Recent Remarks:</b><br>• " + "<br>• ".join(entries[:5])
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
