"""Read-only catalogue model and background picker query."""

import logging
from typing import Any, List

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, QObject, QRunnable, Qt, pyqtSignal


from modules.nas.core.nas_threads import SpecCacheIndex
from modules.spec_search.core.spec_search_db import parse_version_tuple


class SpecificationListModel(QAbstractTableModel):
    headers = ("Spec", "WG", "Releases", "Title")

    def __init__(self, parent):
        super().__init__(parent)
        self.rows = []

    def replace(self, rows):
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.headers)

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.headers[section]
        return super().headerData(section, orientation, role)

    def data(self, index, role):
        if not index.isValid():
            return None
        number, title, kind, _, group, count = self.rows[index.row()]
        column = index.column()
        if role == Qt.DisplayRole:
            return (f"{kind or 'TS'} {number}", group or "-", str(count), title or "")[column]
        if role == Qt.ToolTipRole and column == 3:
            return f"{kind or 'TS'} {number}: {title or ''}"
        if role == Qt.TextAlignmentRole and column == 2:
            return Qt.AlignCenter
        if role == Qt.UserRole and column == 0:
            return number


class CatalogueSignals(QObject):
    loaded = pyqtSignal(int, object)


class CatalogueQuery(QRunnable):
    def __init__(self, specs_db, request_id, filters):
        super().__init__()
        self.specs_db = specs_db
        self.request_id = request_id
        self.filters = filters
        self.signals = CatalogueSignals()

    def run(self):
        search_kw, series_filter, wg_filter = self.filters
        query = """
            SELECT sp.number, sp.title, sp.type, s.name AS series_name, p_grp.name AS wg_name, COUNT(f.id) AS file_count
            FROM specifications sp
            JOIN series s ON sp.series_id = s.id
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN files f ON f.spec_id = sp.id
            WHERE 1=1
        """
        params: List[Any] = []

        if series_filter:
            query += " AND s.name = ?"
            params.append(str(series_filter))

        if wg_filter:
            query += " AND p_grp.name = ?"
            params.append(str(wg_filter))

        if search_kw:
            search_pat = f"%{search_kw}%"
            query += " AND (sp.number LIKE ? OR sp.title LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR ua.tags LIKE ? OR ua.comments LIKE ?)"
            params.extend([search_pat, search_pat, search_pat, search_pat, search_pat])

        query += " GROUP BY sp.id HAVING file_count > 0 ORDER BY CAST(s.name AS INTEGER) ASC, sp.number ASC"

        try:
            with self.specs_db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                spec_rows = cursor.fetchall()
        except Exception as e:
            logging.getLogger(__name__).error(f"Error fetching specs summary: {e}")
            spec_rows = []

        self.signals.loaded.emit(self.request_id, spec_rows)


class ReleaseQuery(QRunnable):
    def __init__(self, specs_db, search_db, request_id, spec_numbers, imported_entries):
        super().__init__()
        self.specs_db = specs_db
        self.search_db = search_db
        self.request_id = request_id
        self.spec_numbers = spec_numbers
        self._imported_entries = imported_entries
        self.signals = CatalogueSignals()

    def run(self):
        try:
            placeholders = ",".join("?" for _ in self.spec_numbers)
            query = f"""
                SELECT sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date
                FROM files f
                JOIN specifications sp ON f.spec_id = sp.id
                WHERE sp.number IN ({placeholders})
            """
            with self.specs_db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, self.spec_numbers)
                file_rows = cursor.fetchall()

            if self._imported_entries is None:
                self._imported_entries = {(v["spec_number"], v["version"]) for v in self.search_db.get_imported_versions()}

            # Stable two-pass sort: version descending first, then spec number ascending
            sorted_files = sorted(file_rows, key=lambda r: parse_version_tuple(r[4]), reverse=True)
            sorted_files.sort(key=lambda r: r[0])

            cache_index = SpecCacheIndex()
            files = [(row, cache_index.find_in_listing_order(row[3], row[0])) for row in sorted_files]
        except Exception:
            logging.getLogger(__name__).exception("Failed to load specification releases")
            files = []
        self.signals.loaded.emit(self.request_id, (self._imported_entries, files))
