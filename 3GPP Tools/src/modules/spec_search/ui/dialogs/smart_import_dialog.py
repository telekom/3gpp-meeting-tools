"""Targeted intelligent import dialog for latest 3GPP specification releases."""

import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QDialog,
    QTableWidgetItem,
    QWidget,
)

from core.ui.theme import C
from modules.spec_search.core.spec_search_db import (
    SpecSearchDatabase,
    extract_major_release,
    parse_version_tuple,
)
from modules.specifications.core.database import SpecsDatabase

from .smart_import_ui import _SmartImportUiMixin


class SpecSmartImportDialog(_SmartImportUiMixin, QDialog):
    """
    Targeted Intelligent Import Dialog for a single 3GPP Specification.
    Groups all historical versions by 3GPP Release (Rel-15, Rel-16, Rel-17, Rel-18, etc.),
    identifies the newest version per release, computes delta status against the local Spec Search DB
    (Up to Date, Update Available, New Release), and allows 1-click batch ingestion.
    """

    def __init__(
        self,
        specs_db: SpecsDatabase,
        search_db: SpecSearchDatabase,
        cache_dir: Path,
        initial_spec: Optional[str] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.search_db = search_db
        self.cache_dir = cache_dir
        self.initial_spec = (initial_spec or "").strip()
        self.logger = logging.getLogger(__name__)

        self.current_spec = ""
        self.current_spec_title = ""
        self.current_spec_wg = ""
        self.releases_data: List[Dict[str, Any]] = []
        self.selected_tasks: List[Dict[str, Any]] = []
        self.auto_replace_older: bool = True

        self.setWindowTitle("⭐ Smart Import Specification (Latest per Release)")
        self.resize(880, 560)
        self._setup_ui()

        if self.initial_spec:
            self.spec_input.setText(self.initial_spec)
            self._on_find_releases()
        else:
            self.spec_input.setFocus()

    def _on_find_releases(self):
        raw = self.spec_input.text().strip()
        if not raw:
            return

        clean_spec = re.sub(r"^(?:TS|TR)\s*", "", raw, flags=re.IGNORECASE).strip()
        if re.match(r"^\d{5}$", clean_spec):
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:]}"
        elif re.match(r"^\d{6}$", clean_spec):
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:5]}-{clean_spec[5:]}"

        self.current_spec = clean_spec

        # Query metadata
        spec_details = self.specs_db.get_spec_details(clean_spec) if hasattr(self.specs_db, "get_spec_details") else {}
        self.current_spec_title = spec_details.get("title", "")
        self.current_spec_wg = spec_details.get("primary_group", "")

        files = self.specs_db.get_spec_files(clean_spec)
        if not files:
            self.table.setRowCount(0)
            self.lbl_spec_title.setText(f"❌ Specification <b>TS {clean_spec}</b> was not found in 3GPP specifications database.")
            self.lbl_spec_title.setStyleSheet(f"color: {C.DANGER_ALT}; font-size: 11px;")
            self.lbl_status_badge.setVisible(False)
            self._update_selection_count()
            return

        # Display header
        title_str = f"<b>TS {clean_spec}</b>"
        if self.current_spec_title:
            title_str += f": {self.current_spec_title}"
        if self.current_spec_wg:
            title_str += f" — <i>{self.current_spec_wg}</i>"
        self.lbl_spec_title.setText(title_str)
        self.lbl_spec_title.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 11px;")

        # Group files by release
        rel_files: Dict[str, Dict[str, Any]] = {}
        for f in files:
            v_str = str(f.get("version", "")).lstrip("vV").strip()
            if not v_str:
                continue
            rel = extract_major_release(v_str)
            if rel not in rel_files:
                rel_files[rel] = f

        # Query currently indexed versions in search_db
        indexed_rows = self.search_db.get_versions_for_spec(clean_spec)
        indexed_by_rel: Dict[str, List[str]] = {}
        for row in indexed_rows:
            v_str = str(row.get("version", "")).lstrip("vV").strip()
            rel = extract_major_release(v_str)
            indexed_by_rel.setdefault(rel, []).append(v_str)

        def _rel_sort_key(rel_str: str) -> int:
            num = rel_str.replace("Rel-", "").strip()
            return int(num) if num.isdigit() else -1

        sorted_releases = sorted(rel_files.keys(), key=_rel_sort_key, reverse=True)

        self.releases_data = []
        updates_count = 0
        new_count = 0
        current_count = 0

        for rel in sorted_releases:
            f = rel_files[rel]
            avail_ver = str(f.get("version", "")).lstrip("vV").strip()
            curr_indexed = indexed_by_rel.get(rel, [])

            if not curr_indexed:
                status = "new"
                status_text = "🔵 New Release"
                status_desc = "Not indexed in Spec Search DB yet."
                new_count += 1
            elif parse_version_tuple(curr_indexed[0]) < parse_version_tuple(avail_ver):
                status = "update"
                status_text = f"🟠 Update (v{curr_indexed[0]} → v{avail_ver})"
                status_desc = f"Currently indexed v{curr_indexed[0]} is older than available v{avail_ver}."
                updates_count += 1
            else:
                status = "current"
                status_text = f"🟢 Up to date (v{avail_ver})"
                status_desc = "Indexed version matches the latest available release."
                current_count += 1

            self.releases_data.append({
                "release": rel,
                "release_num": _rel_sort_key(rel),
                "version": avail_ver,
                "status": status,
                "status_text": status_text,
                "status_desc": status_desc,
                "upload_date": f.get("upload_date", ""),
                "filename": f.get("filename", ""),
                "url": f.get("url", ""),
            })

        self.lbl_status_badge.setText(f"{len(self.releases_data)} releases: {updates_count} update(s), {new_count} new, {current_count} current")
        self.lbl_status_badge.setStyleSheet(f"""
            background-color: {C.BG_CANVAS};
            color: {C.TEXT_BODY};
            border: 1px solid {C.BORDER_STRONG};
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 10px;
            font-weight: bold;
        """)
        self.lbl_status_badge.setVisible(True)

        self._populate_table()
        # Default: Select Updates & New
        self._select_updates_and_new()

    def _populate_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.releases_data))

        for r, item in enumerate(self.releases_data):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(r, 0, chk)

            rel_item = QTableWidgetItem(item["release"])
            rel_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            rel_item.setTextAlignment(Qt.AlignCenter)
            font_b = QFont()
            font_b.setBold(True)
            rel_item.setFont(font_b)
            self.table.setItem(r, 1, rel_item)

            ver_item = QTableWidgetItem(f"v{item['version']}")
            ver_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 2, ver_item)

            status_item = QTableWidgetItem(item["status_text"])
            status_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            status_item.setToolTip(item["status_desc"])
            if item["status"] == "update":
                status_item.setForeground(QColor(C.ORANGE_DARK))
            elif item["status"] == "new":
                status_item.setForeground(QColor(C.ACCENT_SKY_DARK))
            else:
                status_item.setForeground(QColor(C.SUCCESS))
            self.table.setItem(r, 3, status_item)

            d_item = QTableWidgetItem(item.get("upload_date", ""))
            d_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 4, d_item)

            fn_item = QTableWidgetItem(item.get("filename", ""))
            fn_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 5, fn_item)

        self.table.blockSignals(False)

    def _select_updates_and_new(self):
        self.table.blockSignals(True)
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked if item["status"] in ("update", "new") else Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _select_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _select_min_release(self, min_rel: int):
        self.table.blockSignals(True)
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Checked if item["release_num"] >= min_rel else Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _clear_selection(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk:
                chk.setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_count()

    def _on_table_item_changed(self, item: QTableWidgetItem):
        if item.column() == 0:
            self._update_selection_count()

    def _on_cell_double_clicked(self, row: int, col: int):
        chk = self.table.item(row, 0)
        if chk:
            new_st = Qt.Unchecked if chk.checkState() == Qt.Checked else Qt.Checked
            chk.setCheckState(new_st)
            self._update_selection_count()

    def _update_selection_count(self):
        selected = 0
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk and chk.checkState() == Qt.Checked:
                selected += 1

        self.lbl_selection_summary.setText(f"Selected: {selected} release(s)")
        self.btn_ingest.setText(f"📥 Ingest Selected Releases ({selected})")
        self.btn_ingest.setEnabled(selected > 0)

    def _on_ingest_confirmed(self):
        self.selected_tasks = []
        for r, item in enumerate(self.releases_data):
            chk = self.table.item(r, 0)
            if chk and chk.checkState() == Qt.Checked:
                self.selected_tasks.append({
                    "spec_number": self.current_spec,
                    "version": item["version"],
                    "filename": item["filename"],
                    "file_url": item["url"],
                    "release_date": item["upload_date"],
                    "spec_title": self.current_spec_title or "",
                })

        if not self.selected_tasks:
            return

        self.auto_replace_older = self.chk_auto_supersede.isChecked()
        self.accept()
