"""Table model for cross-meeting 3GPP agreements, working assumptions, FFS, and TDocs."""

from typing import List

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor, QFont

from core.ui.theme import C


class AgreementResultsTableModel(QAbstractTableModel):
    """Table model displaying extracted agreements, working assumptions, FFS, and agreed TPs/pCRs."""

    COLUMNS = [
        "WG",
        "Meeting",
        "Date",
        "AI",
        "Work Item",
        "Type",
        "Agreement / Conclusion",
        "AI / Topic Title",
        "Related TDoc",
        "Source",
        "Chair Notes File",
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._all_rows: List[dict] = []
        self._filtered_rows: List[dict] = []
        self._live_filter: str = ""
        self._bold_font = QFont()
        self._bold_font.setBold(True)

    def update_data(self, rows: List[dict]):
        self.beginResetModel()
        self._all_rows = list(rows or [])
        self._apply_live_filter_internal()
        self.endResetModel()

    def set_live_filter(self, query: str):
        self.beginResetModel()
        self._live_filter = (query or "").strip().lower()
        self._apply_live_filter_internal()
        self.endResetModel()

    def _apply_live_filter_internal(self):
        if not self._live_filter:
            self._filtered_rows = list(self._all_rows)
            return
        tokens = [t for t in self._live_filter.split() if t]
        matched = []
        for r in self._all_rows:
            haystack = " ".join([
                str(r.get("WG", "")),
                str(r.get("Meeting", "")),
                str(r.get("AI", "")),
                str(r.get("AI Title", "")),
                str(r.get("Work Item", "")),
                str(r.get("Kind", "")),
                str(r.get("Agreement Text", "")),
                str(r.get("TDoc", "")),
                str(r.get("TDoc Title", "")),
                str(r.get("Source", "")),
            ]).lower()
            if all(tok in haystack for tok in tokens):
                matched.append(r)
        self._filtered_rows = matched

    def get_filtered_rows(self) -> List[dict]:
        return self._filtered_rows

    def get_all_rows(self) -> List[dict]:
        return self._all_rows

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._filtered_rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self.COLUMNS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            if 0 <= section < len(self.COLUMNS):
                return self.COLUMNS[section]
        return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._filtered_rows)):
            return None

        row = self._filtered_rows[index.row()]
        col_name = self.COLUMNS[index.column()]

        if role == Qt.UserRole:
            return row

        if role == Qt.DisplayRole:
            if col_name == "WG":
                return row.get("WG", "")
            if col_name == "Meeting":
                return row.get("Meeting", "")
            if col_name == "Date":
                return row.get("start_date", "")
            if col_name == "AI":
                return row.get("AI", "")
            if col_name == "Work Item":
                return row.get("Work Item", "")
            if col_name == "Type":
                kind = str(row.get("Kind", "Agreement"))
                low = kind.lower()
                if "ffs" in low or "open issue" in low:
                    return f"🔵 {kind}"
                if "working assumption" in low:
                    return f"🟡 {kind}"
                if "agreed tdoc" in low or "pcr" in low:
                    return f"📄 {kind}"
                return f"🟢 {kind}"
            if col_name == "Agreement / Conclusion":
                txt = str(row.get("Agreement Text", "")).replace("\n", " ⏎ ")
                return txt if len(txt) <= 220 else txt[:217] + "..."
            if col_name == "AI / Topic Title":
                return row.get("AI Title", "")
            if col_name == "Related TDoc":
                return row.get("TDoc", "")
            if col_name == "Source":
                return row.get("Source", "")
            if col_name == "Chair Notes File":
                return row.get("Chair Notes File", "")

        if role == Qt.ToolTipRole:
            if col_name in ("Agreement / Conclusion", "Type"):
                return str(row.get("Agreement Text", ""))
            if col_name == "Related TDoc":
                tid = row.get("TDoc", "")
                ttitle = row.get("TDoc Title", "")
                return f"{tid}: {ttitle}" if ttitle else tid
            if col_name == "AI / Topic Title":
                return str(row.get("AI Title", ""))

        if role == Qt.ForegroundRole:
            kind = str(row.get("Kind", "")).lower()
            if col_name == "Type":
                if "ffs" in kind or "open issue" in kind:
                    return QBrush(QColor(C.INFO))
                if "working assumption" in kind:
                    return QBrush(QColor(C.WARNING))
                if "agreed tdoc" in kind:
                    return QBrush(QColor(C.INDIGO_DARK))
                return QBrush(QColor(C.SUCCESS))
            if col_name in ("Work Item", "Related TDoc"):
                return QBrush(QColor(C.ACCENT_PRIMARY))

        if role == Qt.BackgroundRole:
            kind = str(row.get("Kind", "")).lower()
            if col_name == "Type":
                if "ffs" in kind or "open issue" in kind:
                    return QBrush(QColor(C.INFO_SOFT))
                if "working assumption" in kind:
                    return QBrush(QColor(C.WARNING_SOFT_ALT))
                if "agreed tdoc" in kind:
                    return QBrush(QColor(C.WG_SA_BG))
                return QBrush(QColor(C.SUCCESS_SOFT))

        if role == Qt.FontRole:
            if col_name in ("Type", "Work Item", "AI"):
                return self._bold_font

        return None
