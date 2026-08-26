"""
UI Components: Version Tree Widget and Rich Clause Content Inspector with match excerpts,
term navigation, and one-click citation copying.
"""

import html
import re
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QTextCursor, QTextDocument
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
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
    """Renders rich clause content with context excerpts, match navigation, and citation copying."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Clause Content Inspector", parent)
        self._current_search_query = ""
        self._last_citation_text = ""
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Top Control Bar
        header_bar = QHBoxLayout()
        self.title_lbl = QLabel("No clause selected")
        self.title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        header_bar.addWidget(self.title_lbl, stretch=1)

        self.match_badge = QLabel("")
        self.match_badge.setStyleSheet("font-weight: bold; color: #B45309; background-color: #FEF3C7; padding: 2px 6px; border-radius: 4px;")
        self.match_badge.setVisible(False)
        header_bar.addWidget(self.match_badge)

        self.btn_prev = QPushButton("◀ Prev")
        self.btn_prev.setToolTip("Jump to previous match in clause")
        self.btn_prev.clicked.connect(self._find_prev)
        self.btn_prev.setEnabled(False)
        header_bar.addWidget(self.btn_prev)

        self.btn_next = QPushButton("Next ▶")
        self.btn_next.setToolTip("Jump to next match in clause")
        self.btn_next.clicked.connect(self._find_next)
        self.btn_next.setEnabled(False)
        header_bar.addWidget(self.btn_next)

        self.btn_copy = QPushButton("📋 Copy Citation")
        self.btn_copy.setToolTip("Copy clause citation and text to clipboard")
        self.btn_copy.clicked.connect(self._copy_citation)
        self.btn_copy.setEnabled(False)
        header_bar.addWidget(self.btn_copy)

        layout.addLayout(header_bar)

        # Document Browser
        self.browser = QTextBrowser()
        self.browser.setReadOnly(True)
        self.browser.setOpenExternalLinks(True)
        self.browser.setPlaceholderText("Select a cell in the evolution matrix above to inspect the full clause content and match context...")
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
        self._current_search_query = search_query.strip()
        date_str = f" ({release_date})" if release_date else ""
        header_title = f"TS {spec_number} v{version}{date_str} - Clause {clause_number}: {clause_title}"
        self.title_lbl.setText(header_title)

        # Save clean citation string
        self._last_citation_text = f"3GPP TS {spec_number} v{version}{date_str}, Clause {clause_number} ({clause_title}):\n\n{content}"
        self.btn_copy.setEnabled(True)

        # Count match hits
        match_count = 0
        if self._current_search_query:
            match_count = len(re.findall(re.escape(self._current_search_query), content, re.IGNORECASE))

        if match_count > 0:
            self.match_badge.setText(f"🎯 {match_count} Match(es)")
            self.match_badge.setVisible(True)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)
        else:
            self.match_badge.setVisible(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)

        # Extract Key Match Excerpt
        excerpt_html = ""
        if self._current_search_query and match_count > 0:
            excerpt_html = self._generate_match_excerpt(content, self._current_search_query)

        # Format full document HTML with highlighted search terms
        escaped_content = html.escape(content)
        if self._current_search_query:
            pattern = re.compile(re.escape(html.escape(self._current_search_query)), re.IGNORECASE)
            escaped_content = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
                escaped_content,
            )

        formatted_html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; line-height: 1.6; color: #1E293B; margin: 4px;">
            {excerpt_html}
            <div style="border-bottom: 2px solid #E2E8F0; margin: 10px 0 8px 0; padding-bottom: 4px;">
                <span style="color: #0369A1; font-weight: bold; font-size: 13px;">{clause_number} {html.escape(clause_title)}</span>
            </div>
            <pre style="white-space: pre-wrap; font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; color: #334155; line-height: 1.5;">{escaped_content}</pre>
        </body>
        </html>
        """
        self.browser.setHtml(formatted_html)

        # Auto-scroll to first occurrence
        if self._current_search_query and match_count > 0:
            self._find_next()

    def _generate_match_excerpt(self, full_text: str, query: str) -> str:
        """Extracts surrounding paragraph context around the match to show at the top."""
        paragraphs = full_text.split("\n")
        matched_paras = [p for p in paragraphs if query.lower() in p.lower()]

        if not matched_paras:
            return ""

        excerpt = matched_paras[0].strip()
        escaped_excerpt = html.escape(excerpt)
        pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
        highlighted_excerpt = pattern.sub(
            lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
            escaped_excerpt,
        )

        return f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid #0284C7; border: 1px solid #E2E8F0; border-left-width: 4px; border-radius: 4px; padding: 8px 10px; margin-bottom: 12px;">
            <div style="font-weight: bold; color: #0369A1; font-size: 11px; margin-bottom: 4px;">💡 Key Match Excerpt & Context:</div>
            <div style="color: #1E293B; font-size: 11px; font-style: normal; line-height: 1.5;">{highlighted_excerpt}</div>
        </div>
        """

    def _find_next(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query)
        if not found:
            self.browser.moveCursor(QTextCursor.Start)
            self.browser.find(self._current_search_query)

    def _find_prev(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query, QTextDocument.FindBackward)
        if not found:
            self.browser.moveCursor(QTextCursor.End)
            self.browser.find(self._current_search_query, QTextDocument.FindBackward)

    def _copy_citation(self):
        if self._last_citation_text:
            QApplication.clipboard().setText(self._last_citation_text)
            self.btn_copy.setText("✅ Copied!")
            QApplication.processEvents()
            self.btn_copy.setText("📋 Copy Citation")

    def clear_display(self):
        self.title_lbl.setText("No clause selected")
        self.match_badge.setVisible(False)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self.btn_copy.setEnabled(False)
        self.browser.clear()
        self._current_search_query = ""
        self._last_citation_text = ""