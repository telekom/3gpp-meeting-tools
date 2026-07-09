# --- File: src/modules/emails/ui/email_models.py ---
import os
from PyQt5.QtCore import Qt, QAbstractTableModel, QSortFilterProxyModel, QModelIndex

class EmailTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self.starred_tdocs = set()
        self.followed_ais = set()
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

        if role == Qt.UserRole + 1 and col_name == "Status": return row.get("id", "")
        if role == Qt.UserRole + 2 and col_name == "⭐": return row.get("tdoc_id") in self.starred_tdocs
        if role == Qt.UserRole + 3 and col_name == "AI": return row.get("agenda_item") in self.followed_ais

        return None

    def get_row_data(self, row_idx: int) -> dict:
        if 0 <= row_idx < len(self._data):
            return self._data[row_idx]
        return {}

    def rowCount(self, index=QModelIndex()): return len(self._data)
    def columnCount(self, index=QModelIndex()): return len(self._headers)
    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None

# ==========================================
# LEFT PANEL MODELS (TDoc Summary)
# ==========================================
class TDocSummaryModel(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self._data = []
        self.starred_tdocs = set()
        self.followed_ais = set()
        self._headers = ["⭐", "TDoc / Topic", "AI", "Emails"]

    def update_data(self, raw_email_data, starred_tdocs, followed_ais):
        self.beginResetModel()
        self.starred_tdocs = set(starred_tdocs)
        self.followed_ais = set(followed_ais)

        groups = {}
        for row in raw_email_data:
            tid = str(row.get("tdoc_id", "")).strip()
            if not tid: tid = "General / Unlinked"

            if tid not in groups:
                groups[tid] = {'tdoc_id': tid, 'ai': str(row.get('agenda_item', '')).strip(), 'count': 0}
            groups[tid]['count'] += 1

        self._data = sorted(groups.values(), key=lambda x: (x['tdoc_id'] != "General / Unlinked", x['tdoc_id']))
        self.endResetModel()

    def rowCount(self, parent=None): return len(self._data)
    def columnCount(self, parent=None): return len(self._headers)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
        row = self._data[index.row()]
        col = index.column()

        if role == Qt.DisplayRole:
            if col == 0: return "⭐" if row['tdoc_id'] in self.starred_tdocs else ""
            if col == 1: return row['tdoc_id']
            if col == 2: return row['ai']
            if col == 3: return str(row['count'])

        elif role == Qt.UserRole: return row['tdoc_id']
        elif role == Qt.TextAlignmentRole:
            if col in [0, 3]: return Qt.AlignCenter
        return None


class TDocProxyModel(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.show_starred_only = False
        self.show_followed_only = False
        self.ai_filters = set()
        self.search_text = ""

    # ---> RESTORED: Unified Filter Setter matching your UI
    def set_filters(self, starred_only, followed_only, ais, search_text):
        self.show_starred_only = starred_only
        self.show_followed_only = followed_only
        self.ai_filters = set(ais) if ais else set()
        self.search_text = search_text.lower().strip() if search_text else ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        tdoc_id = model.data(model.index(source_row, 1, source_parent))
        ai = model.data(model.index(source_row, 2, source_parent))

        tdoc_str = str(tdoc_id) if tdoc_id else ""
        ai_str = str(ai) if ai else ""

        if tdoc_str == "General / Unlinked" and not self.ai_filters and not self.search_text:
            pass
        elif tdoc_str == "General / Unlinked" and (self.ai_filters or self.search_text):
            return False

        if self.show_starred_only and tdoc_str not in model.starred_tdocs: return False
        if self.show_followed_only and ai_str not in model.followed_ais: return False
        if self.ai_filters and ai_str not in self.ai_filters: return False

        if self.search_text:
            if self.search_text not in tdoc_str.lower() and self.search_text not in ai_str.lower():
                return False

        return True


# ==========================================
# RIGHT PANEL MODELS (Email Thread)
# ==========================================
class EmailProxyModel(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.target_tdoc = None
        self.global_filter = ""
        self.company_filters = set()
        self.sender_filters = set()

    def set_target_tdoc(self, tdoc_id):
        self.target_tdoc = tdoc_id
        self.invalidateFilter()

    # ---> RESTORED: Unified Filter Setter matching your UI
    def set_filters(self, search_text, companies, senders):
        self.global_filter = search_text.lower().strip() if search_text else ""
        self.company_filters = set(companies) if companies else set()
        self.sender_filters = set(senders) if senders else set()
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        if self.target_tdoc:
            row_tdoc = str(model.data(model.index(source_row, 4, source_parent), Qt.DisplayRole)).strip()
            if not row_tdoc and self.target_tdoc == "General / Unlinked":
                pass
            elif row_tdoc != self.target_tdoc:
                return False

        if self.company_filters:
            row_comp = str(model.data(model.index(source_row, 7, source_parent), Qt.DisplayRole)).strip()
            if row_comp not in self.company_filters:
                return False

        if self.sender_filters:
            row_sender = str(model.data(model.index(source_row, 8, source_parent), Qt.DisplayRole)).strip()
            if row_sender not in self.sender_filters:
                return False

        if self.global_filter:
            company = str(model.data(model.index(source_row, 7, source_parent), Qt.DisplayRole)).lower()
            sender = str(model.data(model.index(source_row, 8, source_parent), Qt.DisplayRole)).lower()
            short_text = str(model.data(model.index(source_row, 9, source_parent), Qt.DisplayRole)).lower()

            search_pool = f"{company} {sender} {short_text}"
            if self.global_filter not in search_pool:
                return False

        return True