import os
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel


class EmailTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self.starred_tdocs = set()
        self.followed_ais = set()  # <--- NEW
        self._headers = ["⭐", "Status", "Local", "Date", "TDoc", "Rev", "AI", "Company", "Sender", "Short Text"]

    def update_data(self, new_data, starred_tdocs, followed_ais):
        self.beginResetModel()
        self._data = new_data
        self.starred_tdocs = set(starred_tdocs)
        self.followed_ais = set(followed_ais)
        self.endResetModel()

    def data(self, index, role):
        if not index.isValid(): return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if role == Qt.DisplayRole or role == Qt.UserRole:
            if col_name == "⭐":
                return "⭐" if row.get("tdoc_id") in self.starred_tdocs else ""
            if col_name == "Status":
                loc = row.get("outlook_location", "Source")
                return "📁 Target" if loc == "Target" else "📥 Source"
            if col_name == "Local":
                path = row.get("msg_path", "")
                return "✅ Disk" if path and os.path.exists(path) else "❌ Missing"
            if col_name == "Date": return row.get("date_received", "")[:16]
            if col_name == "TDoc": return row.get("tdoc_id", "")
            if col_name == "AI":
                ai = row.get("agenda_item", "")
                return f"👀 {ai}" if ai in self.followed_ais else ai
            if col_name == "Company": return row.get("company", "")
            if col_name == "Sender": return row.get("sender_name", "")

            # ---> NEW: Revision Shortening
            if col_name == "Rev":
                revs = row.get("revisions_mentioned", "")
                base_tdoc = row.get("tdoc_id", "")
                if revs and base_tdoc:
                    import re
                    return re.sub(re.escape(base_tdoc), "", revs).strip()
                return revs

            if col_name == "Short Text":
                text = str(row.get("short_text", ""))
                return text.replace('\n', ' ')[:80] + "..." if len(text) > 80 else text

        elif role == Qt.TextAlignmentRole:
            if col_name in ["⭐", "Status", "Local", "Date", "TDoc", "Rev", "AI", "Company"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        # ---> NEW: Hyperlink Styling (Blue text + Underline)
        elif role == Qt.ForegroundRole:
            if col_name in ["Sender", "Rev"]:
                from PyQt5.QtGui import QColor
                return QColor("#005A9E")
        elif role == Qt.FontRole:
            if col_name in ["Sender", "Rev"]:
                from PyQt5.QtGui import QFont
                font = QFont()
                font.setUnderline(True)
                return font

        # Hidden data roles
        if role == Qt.UserRole + 1 and col_name == "Status": return row.get("id", "")
        if role == Qt.UserRole + 2 and col_name == "⭐": return row.get("tdoc_id") in self.starred_tdocs
        if role == Qt.UserRole + 3 and col_name == "AI": return row.get("agenda_item") in self.followed_ais

        return None

    def get_row_data(self, row_idx: int) -> dict:
        if 0 <= row_idx < len(self._data):
            return self._data[row_idx]
        return {}

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        return None


class EmailProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.global_filter = ""
        self.ai_filters = set()
        self.company_filters = set()
        self.sender_filters = set()
        self.show_starred_only = False
        self.show_followed_only = False  # <--- NEW

    def set_filters(self, text, ais, companies, senders, starred_only, followed_only):
        self.global_filter = str(text).lower().strip()
        self.ai_filters = set(ais)
        self.company_filters = set(companies)
        self.sender_filters = set(senders)
        self.show_starred_only = starred_only
        self.show_followed_only = followed_only  # <--- NEW
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # 1. Star & Follow Filter Checks
        if self.show_starred_only:
            if not model.data(model.index(source_row, 0, source_parent), Qt.UserRole + 2): return False

        if self.show_followed_only:  # <--- NEW
            if not model.data(model.index(source_row, 6, source_parent), Qt.UserRole + 3): return False

        # 2. Combobox Filters
        if self.ai_filters and model.data(model.index(source_row, 6, source_parent),
                                          Qt.UserRole) not in self.ai_filters: return False
        if self.company_filters and model.data(model.index(source_row, 7, source_parent),
                                               Qt.UserRole) not in self.company_filters: return False
        if self.sender_filters and model.data(model.index(source_row, 8, source_parent),
                                              Qt.UserRole) not in self.sender_filters: return False

        # 3. Global Text Search
        if self.global_filter:
            row_data = model.get_row_data(source_row)
            search_pool = f"{row_data.get('subject', '')} {row_data.get('short_text', '')} {row_data.get('free_text', '')} {row_data.get('tdoc_id', '')} {row_data.get('revisions_mentioned', '')}".lower()
            if self.global_filter not in search_pool:
                return False

        return True