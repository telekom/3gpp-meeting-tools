"""
UI Components: Version Tree Widget and Clause Content Inspector with live term highlighting.
"""

import html
import re
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAction,
    QGroupBox,
    QLabel,
    QMenu,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.spec_search.core.spec_search_db import parse_version_tuple


class SpecSearchVersionTreeWidget(QTreeWidget):
    """Tri-state checkable tree grouping indexed specification releases with release dates."""

    selection_changed = pyqtSignal()
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._updating_checks = False
        self.setHeaderHidden(True)
        self.setAlternatingRowColors(True)
        self.setIndentation(14)
        self.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                padding: 1px;
            }
            QTreeWidget::item:selected {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        self.itemChanged.connect(self._on_item_changed)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

    def populate(self, versions: List[Dict[str, Any]], saved_checked: Optional[List[Dict[str, Any]]] = None):
        self._updating_checks = True
        self.clear()

        if not versions:
            self._updating_checks = False
            self.selection_changed.emit()
            return

        saved_tuples = {(i.get("spec_number"), i.get("version")) for i in saved_checked} if saved_checked else None

        base_font = self.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        all_item = QTreeWidgetItem(self)
        all_item.setText(0, "All Indexed Specifications")
        all_item.setData(0, Qt.UserRole, {"type": "all"})
        all_item.setFlags(all_item.flags() | Qt.ItemIsUserCheckable)
        all_item.setFont(0, bold_font)
        all_item.setCheckState(0, Qt.Checked)

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        for spec_num in sorted(specs_map.keys()):
            spec_item = QTreeWidgetItem(self)
            spec_item.setText(0, f"TS {spec_num}")
            spec_item.setData(0, Qt.UserRole, {"type": "spec", "spec_number": spec_num})
            spec_item.setFlags(spec_item.flags() | Qt.ItemIsUserCheckable)
            spec_item.setFont(0, bold_font)
            spec_item.setCheckState(0, Qt.Checked)

            sorted_vers = sorted(specs_map[spec_num], key=lambda x: parse_version_tuple(x["version"]), reverse=True)
            for v in sorted_vers:
                child = QTreeWidgetItem(spec_item)
                date_label = f" ({v['release_date']})" if v.get("release_date") else ""
                child.setText(0, f"v{v['version']}{date_label}")
                child.setData(0, Qt.UserRole, {
                    "type": "version",
                    "id": v["id"],
                    "spec_number": spec_num,
                    "version": v["version"],
                    "release_date": v.get("release_date", ""),
                })
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                child.setFont(0, base_font)

                is_checked = (spec_num, v["version"]) in saved_tuples if saved_tuples else True
                child.setCheckState(0, Qt.Checked if is_checked else Qt.Unchecked)

            spec_item.setExpanded(True)

        self._update_parent_states()
        self._updating_checks = False
        self.selection_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        selected_ids = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    d = child.data(0, Qt.UserRole)
                    if d and "id" in d:
                        selected_ids.append(d["id"])
        return selected_ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        checked = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    d = child.data(0, Qt.UserRole) or {}
                    checked.append({"spec_number": d.get("spec_number"), "version": d.get("version")})
        return checked

    def _on_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._updating_checks:
            return
        self._updating_checks = True
        state = item.checkState(0)
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        if item_type == "all":
            target = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for s_idx in range(1, self.topLevelItemCount()):
                s_item = self.topLevelItem(s_idx)
                s_item.setCheckState(0, target)
                for c_idx in range(s_item.childCount()):
                    s_item.child(c_idx).setCheckState(0, target)
        elif item_type == "spec":
            target = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for c_idx in range(item.childCount()):
                item.child(c_idx).setCheckState(0, target)
            self._update_parent_states()
        elif item_type == "version":
            self._update_parent_states()

        self._updating_checks = False
        self.selection_changed.emit()

    def _update_parent_states(self):
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            c_cnt = spec_item.childCount()
            if c_cnt == 0:
                continue
            checked = sum(1 for i in range(c_cnt) if spec_item.child(i).checkState(0) == Qt.Checked)
            spec_item.setCheckState(0, Qt.Checked if checked == c_cnt else (Qt.Unchecked if checked == 0 else Qt.PartiallyChecked))

    def _on_context_menu(self, pos):
        item = self.itemAt(pos)
        if not item:
            return
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        menu = QMenu(self)
        if item_type == "version":
            s_num = data.get("spec_number", "")
            ver = data.get("version", "")
            act = QAction(f"🗑️ Delete TS {s_num} v{ver}", self)
            act.triggered.connect(lambda: self.delete_version_requested.emit(s_num, ver))
            menu.addAction(act)
            menu.exec_(self.viewport().mapToGlobal(pos))
        elif item_type == "spec":
            s_num = data.get("spec_number", "")
            c_cnt = item.childCount()
            act = QAction(f"🗑️ Delete all {c_cnt} version(s) of TS {s_num}", self)
            act.triggered.connect(lambda: self.delete_spec_requested.emit(s_num, c_cnt))
            menu.addAction(act)
            menu.exec_(self.viewport().mapToGlobal(pos))


class SpecClauseInspector(QGroupBox):
    """Renders full specification clause content with highlighted matches."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Clause Content Inspector", parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        self.title_lbl = QLabel("No clause selected")
        self.title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        layout.addWidget(self.title_lbl)

        self.browser = QTextBrowser()
        self.browser.setReadOnly(True)
        self.browser.setPlaceholderText("Select a cell in the matrix to view the full clause text with highlighted matches...")
        layout.addWidget(self.browser)

    def display_clause(
        self,
        clause_number: str,
        clause_title: str,
        spec_number: str,
        version: str,
        content: str,
        release_date: Optional[str] = None,
        search_query: str = "",
    ):
        date_str = f" ({release_date})" if release_date else ""
        self.title_lbl.setText(f"TS {spec_number} v{version}{date_str} - Clause {clause_number}: {clause_title}")

        escaped = html.escape(content)
        if search_query.strip():
            pattern = re.compile(re.escape(html.escape(search_query.strip())), re.IGNORECASE)
            escaped = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 3px; border-radius: 2px;">{m.group(0)}</mark>',
                escaped,
            )

        formatted_html = f"""
        <div style="font-family: Segoe UI, sans-serif; font-size: 11px; line-height: 1.5; color: #1E293B;">
            <h3 style="color: #0369A1; margin-bottom: 8px;">{clause_number} {html.escape(clause_title)}</h3>
            <pre style="white-space: pre-wrap; font-family: Segoe UI, sans-serif; font-size: 11px; color: #334155;">{escaped}</pre>
        </div>
        """
        self.browser.setHtml(formatted_html)

    def clear_display(self):
        self.title_lbl.setText("No clause selected")
        self.browser.clear()