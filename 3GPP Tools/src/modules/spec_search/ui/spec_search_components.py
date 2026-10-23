"""
UI Components: Version Tree Widget and Rich Clause Content Inspector with match excerpts,
term navigation, one-click citation copying, direct document opening, and persistent selection state.
"""

import html
import re
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
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

from modules.spec_search.core.spec_clause_exporter import copy_clause_to_clipboard
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
        """
        Populates the tree with indexed specifications and releases.
        Restores check states from saved_checked (if provided), or defaults to all checked.
        """
        self._updating_checks = True
        self.clear()

        if not versions:
            self._updating_checks = False
            self.selection_changed.emit()
            return

        saved_tuples = (
            {(str(i.get("spec_number")), str(i.get("version"))) for i in saved_checked if i}
            if saved_checked is not None
            else None
        )

        base_font = self.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        all_item = QTreeWidgetItem(self)
        all_item.setText(0, "All Indexed Specifications")
        all_item.setData(0, Qt.UserRole, {"type": "all"})
        all_item.setFlags(all_item.flags() | Qt.ItemIsUserCheckable)
        all_item.setFont(0, bold_font)

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        for spec_num in sorted(specs_map.keys()):
            spec_item = QTreeWidgetItem(self)
            spec_item.setText(0, f"TS {spec_num}")
            spec_item.setData(0, Qt.UserRole, {"type": "spec", "spec_number": spec_num})
            spec_item.setFlags(spec_item.flags() | Qt.ItemIsUserCheckable)
            spec_item.setFont(0, bold_font)

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

                # If saved_checked was provided, restore exact check state; otherwise default to Checked
                if saved_tuples is not None:
                    is_checked = (str(spec_num), str(v["version"])) in saved_tuples
                else:
                    is_checked = True

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
        """Returns list of all checked specification versions for state persistence."""
        checked = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    d = child.data(0, Qt.UserRole) or {}
                    checked.append({
                        "spec_number": str(d.get("spec_number", "")),
                        "version": str(d.get("version", "")),
                    })
        return checked

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        """Returns all versions (regardless of check state) under the given spec_number."""
        results = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            s_data = spec_item.data(0, Qt.UserRole) or {}
            if s_data.get("spec_number") == spec_number:
                for c_idx in range(spec_item.childCount()):
                    child = spec_item.child(c_idx)
                    d = child.data(0, Qt.UserRole) or {}
                    results.append({
                        "spec_number": str(d.get("spec_number", "")),
                        "version": str(d.get("version", "")),
                    })
        return results

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
        total_children = 0
        total_checked = 0

        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            c_cnt = spec_item.childCount()
            if c_cnt == 0:
                continue
            checked = sum(1 for i in range(c_cnt) if spec_item.child(i).checkState(0) == Qt.Checked)
            total_children += c_cnt
            total_checked += checked
            spec_item.setCheckState(
                0, Qt.Checked if checked == c_cnt else (Qt.Unchecked if checked == 0 else Qt.PartiallyChecked)
            )

        # Update root "All Indexed Specifications" item
        if self.topLevelItemCount() > 0:
            all_item = self.topLevelItem(0)
            if total_children > 0:
                all_item.setCheckState(
                    0, Qt.Checked if total_checked == total_children else (Qt.Unchecked if total_checked == 0 else Qt.PartiallyChecked)
                )

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
    """
    Renders rich clause content with context excerpts, match navigation, citation copying,
    1-click Word document launching, and deep-linking to the Specifications catalogue.
    """

    compare_clause_requested = pyqtSignal(str, str, str, str)  # spec_number, version, clause_number, clause_title
    open_document_requested = pyqtSignal(str, str, str, str)    # spec_number, version, clause_number, format ("word", "pdf", "folder")
    show_in_specs_tab_requested = pyqtSignal(str, str)          # spec_number, version
    quick_diff_requested = pyqtSignal(str, str, str, str, str)  # spec_number, base_ver, target_ver, clause_number, clause_title
    jump_to_version_requested = pyqtSignal(str, str, str)       # spec_number, target_version, clause_number
    extract_clause_requested = pyqtSignal(str, str, str, str, str, bool)  # spec_number, version, clause_number, clause_title, content, save_as
    export_markdown_requested = pyqtSignal(str, str, str, str, str)       # spec_number, version, clause_number, clause_title, content

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Clause Content Inspector", parent)
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Top Control Bar
        header_bar = QHBoxLayout()
        header_bar.setSpacing(6)

        self.title_lbl = QLabel("No clause selected")
        self.title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        header_bar.addWidget(self.title_lbl, stretch=1)

        # Clickable First Introduced Chip
        self.first_added_badge = QPushButton("")
        self.first_added_badge.setCursor(Qt.PointingHandCursor)
        self.first_added_badge.setToolTip("Click to jump directly to the version where this clause was first introduced")
        self.first_added_badge.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                font-size: 11px;
                background-color: #DCFCE7;
                color: #15803D;
                border: 1px solid #86EFAC;
                padding: 1px 8px;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #BBF7D0;
                border-color: #4ADE80;
            }
        """)
        self.first_added_badge.setVisible(False)
        self.first_added_badge.clicked.connect(self._on_first_added_clicked)
        header_bar.addWidget(self.first_added_badge)

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

        self.btn_copy = QPushButton("📋 Copy ▾")
        self.btn_copy.setToolTip("Copy clause text (Rich HTML for Email/Outlook, Markdown quote block, or Plain Text)")
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        copy_menu = QMenu(self)
        act_cp_email = copy_menu.addAction("📋 Copy for Email / Outlook (Rich Text)")
        act_cp_email.triggered.connect(lambda: self._copy_formatted("email"))
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown (Quote Block)")
        act_cp_md.triggered.connect(lambda: self._copy_formatted("markdown"))
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(lambda: self._copy_formatted("text"))
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(lambda: self._copy_formatted("citation"))
        self.btn_copy.setMenu(copy_menu)
        self.btn_copy.setEnabled(False)
        header_bar.addWidget(self.btn_copy)

        # Clause Extraction to Word / Markdown
        self.btn_extract = QPushButton("✂️ Extract ▾")
        self.btn_extract.setToolTip("Extract this clause to a clean Word (.docx) document or export to Markdown (.md)")
        self.btn_extract.setCursor(Qt.PointingHandCursor)
        self.btn_extract.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #F0FDF4;
                color: #166534;
                border: 1px solid #BBF7D0;
                padding: 2px 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #DCFCE7;
                color: #14532D;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        extract_menu = QMenu(self)
        act_ext_word = extract_menu.addAction("✂️ Extract Clause in Word (.docx)")
        act_ext_word.triggered.connect(lambda: self._on_extract_clicked(save_as=False))
        act_ext_save = extract_menu.addAction("💾 Save Clause as Word (.docx)...")
        act_ext_save.triggered.connect(lambda: self._on_extract_clicked(save_as=True))
        extract_menu.addSeparator()
        act_ext_md = extract_menu.addAction("📘 Export Clause to Markdown (.md)...")
        act_ext_md.triggered.connect(self._on_export_markdown_clicked)
        self.btn_extract.setMenu(extract_menu)
        self.btn_extract.setEnabled(False)
        header_bar.addWidget(self.btn_extract)

        # 1-Click Open Spec Document Button with Format Menu
        self.btn_open_doc = QPushButton("📖 Full Doc ▾")
        self.btn_open_doc.setToolTip("Open complete specification document in Microsoft Word or PDF")
        self.btn_open_doc.setCursor(Qt.PointingHandCursor)
        self.btn_open_doc.setStyleSheet("""
            QPushButton {
                font-weight: 500;
                background-color: #F8FAFC;
                color: #334155;
                border: 1px solid #CBD5E1;
                padding: 2px 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                color: #0F172A;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        doc_menu = QMenu(self)
        act_word = doc_menu.addAction("📝 Open Full Specification in Word (.docx / .doc)")
        act_word.triggered.connect(lambda: self._on_open_doc_clicked("word"))
        act_pdf = doc_menu.addAction("📕 Open Converted PDF (.pdf)")
        act_pdf.triggered.connect(lambda: self._on_open_doc_clicked("pdf"))
        doc_menu.addSeparator()
        act_folder = doc_menu.addAction("📂 Open Containing Folder in Explorer")
        act_folder.triggered.connect(lambda: self._on_open_doc_clicked("folder"))
        self.btn_open_doc.setMenu(doc_menu)
        self.btn_open_doc.setEnabled(False)
        header_bar.addWidget(self.btn_open_doc)

        # Deep Link: Show in Specifications Tab
        self.btn_show_in_specs = QPushButton("🔍 Specs Tab")
        self.btn_show_in_specs.setToolTip("Switch to Specifications Tab with this spec and version selected")
        self.btn_show_in_specs.setCursor(Qt.PointingHandCursor)
        self.btn_show_in_specs.setStyleSheet("""
            QPushButton {
                font-weight: 500;
                background-color: #EBF3FC;
                color: #1E5C99;
                border: 1px solid #B0D0F0;
                padding: 2px 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #DFEDFA;
                color: #004488;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        self.btn_show_in_specs.clicked.connect(self._on_show_in_specs_clicked)
        self.btn_show_in_specs.setEnabled(False)
        header_bar.addWidget(self.btn_show_in_specs)

        # Quick Diff with Preceding Release
        self.btn_quick_diff = QPushButton("🔍 Diff Prev")
        self.btn_quick_diff.setToolTip("Compare this clause with the immediately preceding release")
        self.btn_quick_diff.setCursor(Qt.PointingHandCursor)
        self.btn_quick_diff.setStyleSheet("""
            QPushButton {
                font-weight: 500;
                background-color: #FEF3C7;
                color: #92400E;
                border: 1px solid #FCD34D;
                padding: 2px 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #FDE68A;
                color: #78350F;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        self.btn_quick_diff.clicked.connect(self._on_quick_diff_clicked)
        self.btn_quick_diff.setEnabled(False)
        header_bar.addWidget(self.btn_quick_diff)

        # Dedicated LLM Diff Button
        self.btn_llm_diff = QPushButton("🤖 LLM Diff")
        self.btn_llm_diff.setToolTip("Compare this clause with another version and generate an LLM analysis prompt with hierarchical context.")
        self.btn_llm_diff.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #EDE9FE;
                color: #6D28D9;
                border: 1px solid #DDD6FE;
                padding: 2px 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #DDD6FE;
                color: #5B21B6;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        self.btn_llm_diff.clicked.connect(self._on_llm_diff_clicked)
        self.btn_llm_diff.setEnabled(False)
        header_bar.addWidget(self.btn_llm_diff)

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
        first_added_ver: Optional[str] = None,
        prev_ver: Optional[str] = None,
    ):
        self._current_spec = spec_number
        self._current_version = version
        self._current_clause_num = clause_number
        self._current_clause_title = clause_title
        self._current_content = content
        self._current_release_date = release_date or ""
        self._current_search_query = search_query.strip()
        self._first_added_ver = first_added_ver or ""
        self._prev_ver = prev_ver or ""

        date_str = f" ({release_date})" if release_date else ""
        header_title = f"TS {spec_number} v{version}{date_str} - Clause {clause_number}: {clause_title}"
        self.title_lbl.setText(header_title)

        self._last_citation_text = f"3GPP TS {spec_number} v{version}{date_str}, Clause {clause_number} ({clause_title}):\n\n{content}"
        self.btn_copy.setEnabled(True)
        self.btn_extract.setEnabled(True)
        self.btn_llm_diff.setEnabled(True)
        self.btn_open_doc.setEnabled(True)
        self.btn_show_in_specs.setEnabled(True)

        # Update Quick Diff button
        if self._prev_ver:
            self.btn_quick_diff.setEnabled(True)
            self.btn_quick_diff.setToolTip(f"Compare clause {clause_number} between preceding v{self._prev_ver} and current v{version}")
        else:
            self.btn_quick_diff.setEnabled(False)
            self.btn_quick_diff.setToolTip("No preceding release indexed to compare against")

        # Update First Introduced chip
        if self._first_added_ver:
            clean_first = self._first_added_ver.lstrip("v")
            clean_curr = version.lstrip("v")
            if clean_first == clean_curr:
                self.first_added_badge.setText(f"🟢 First Added in v{clean_first}")
            else:
                self.first_added_badge.setText(f"🟢 Introduced in v{clean_first} (Click to View)")
            self.first_added_badge.setVisible(True)
        else:
            self.first_added_badge.setVisible(False)

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

        excerpt_html = ""
        if self._current_search_query and match_count > 0:
            excerpt_html = self._generate_match_excerpt(content, self._current_search_query)

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

        if self._current_search_query and match_count > 0:
            self._find_next()

    def _on_first_added_clicked(self):
        if self._current_spec and self._first_added_ver:
            self.jump_to_version_requested.emit(
                self._current_spec,
                self._first_added_ver,
                self._current_clause_num,
            )

    def _on_open_doc_clicked(self, fmt="word"):
        if self._current_spec and self._current_version:
            self.open_document_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                fmt,
            )

    def _on_show_in_specs_clicked(self):
        if self._current_spec:
            self.show_in_specs_tab_requested.emit(
                self._current_spec,
                self._current_version,
            )

    def _on_quick_diff_clicked(self):
        if self._current_spec and self._prev_ver and self._current_version:
            self.quick_diff_requested.emit(
                self._current_spec,
                self._prev_ver,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _on_llm_diff_clicked(self):
        if self._current_spec and self._current_clause_num:
            self.compare_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _generate_match_excerpt(self, full_text: str, query: str) -> str:
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

    def _copy_formatted(self, mode: str = "email"):
        if not self._current_spec or not self._current_clause_num or not self._current_content:
            return
        status = copy_clause_to_clipboard(
            self._current_spec,
            self._current_version,
            self._current_clause_num,
            self._current_clause_title,
            self._current_content,
            self._current_release_date,
            mode=mode,
        )
        self.btn_copy.setText("✅ Copied!")
        QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 Copy ▾"))

    def _on_extract_clicked(self, save_as: bool = False):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.extract_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
                save_as,
            )

    def _on_export_markdown_clicked(self):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.export_markdown_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
            )

    def _copy_citation(self):
        self._copy_formatted("citation")

    def clear_display(self):
        self.title_lbl.setText("No clause selected")
        self.match_badge.setVisible(False)
        self.first_added_badge.setVisible(False)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self.btn_copy.setEnabled(False)
        self.btn_extract.setEnabled(False)
        self.btn_open_doc.setEnabled(False)
        self.btn_show_in_specs.setEnabled(False)
        self.btn_quick_diff.setEnabled(False)
        self.btn_llm_diff.setEnabled(False)
        self.browser.clear()
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
