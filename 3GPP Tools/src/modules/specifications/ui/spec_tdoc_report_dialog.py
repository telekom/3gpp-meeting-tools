# --- File: src/modules/specifications/ui/spec_tdoc_report_dialog.py ---
"""
Interactive Specification TDoc Metadata Report & Technical Dossier Dialog.

Provides delegates with an on-demand searchable dossier of all TDocs integrated
or referenced in a specification version. Displays corporate authorship, meeting
approval status, agenda items, and solution abstracts without downloading actual TDocs.
"""

import html
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PyQt5.QtCore import QPoint, Qt, QTimer
from PyQt5.QtGui import QBrush, QColor, QCursor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    escape_qt_mnemonic,
)
from core.utils.paths import get_project_root
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

logger = logging.getLogger(__name__)


class SpecTDocReportDialog(QDialog):
    """Modeless technical dossier dialog for inspecting all TDocs integrated in a specification."""

    def __init__(
        self,
        spec_number: str,
        version: str,
        tdoc_entries: Sequence[Dict[str, Any]],
        markdown_report: str = "",
        spec_title: str = "",
        parent: Optional[QWidget] = None,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top)
        self._parent_ref = parent
        self.spec_number = spec_number.strip()
        self.version = version.lstrip("vV").strip()
        self.all_entries = list(tdoc_entries)
        self.filtered_entries = list(tdoc_entries)
        self.markdown_report = markdown_report
        self.spec_title = spec_title

        # Top-level modeless independent window configuration
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass

        self.setWindowTitle(f"📑 TDoc Technical Dossier — TS {self.spec_number} v{self.version}")
        self.resize(1120, 740)
        self.setMinimumSize(880, 540)

        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8FAFC; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()
        self._populate_filters()
        self._refresh_table()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # -------------------------------------------------------------
        # 1. HEADER RIBBON & KPI SUMMARY CARD
        # -------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        spec_pill = QLabel(f"<b>TS {self.spec_number}</b>")
        spec_pill.setStyleSheet(BADGE_STYLE_PRIMARY + " font-size: 14px; padding: 4px 10px;")

        ver_pill = QLabel(f"<b>v{self.version}</b>")
        ver_pill.setStyleSheet("""
            background-color: #ECFDF5;
            color: #065F46;
            border: 1px solid #A7F3D0;
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
            padding: 3px 8px;
        """)

        title_text = self.spec_title or "Integrated Technical Contributions Dossier"
        self.title_lbl = QLabel(title_text)
        self.title_lbl.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B;")
        self.title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top_row.addWidget(spec_pill)
        top_row.addWidget(ver_pill)
        top_row.addWidget(self.title_lbl, stretch=1)

        # KPI Badges
        total_cnt = len(self.all_entries)
        sources = {r.get("source") for r in self.all_entries if r.get("source")}
        approved_cnt = sum(1 for r in self.all_entries if "approv" in r.get("status", "") or "agree" in r.get("status", ""))

        self.kpi_lbl = QLabel(
            f"<span style='color: #475569;'>Total:</span> <b>{total_cnt}</b> TDocs &nbsp;|&nbsp; "
            f"<span style='color: #475569;'>Sources:</span> <b>{len(sources)}</b> Companies &nbsp;|&nbsp; "
            f"<span style='color: #166534;'>Approved/Agreed:</span> <b>{approved_cnt}</b>"
        )
        self.kpi_lbl.setStyleSheet("font-size: 12px; background-color: #F1F5F9; padding: 4px 10px; border-radius: 4px;")
        top_row.addWidget(self.kpi_lbl)

        header_layout.addLayout(top_row)

        # Filter bar
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Filter by TDoc, title, company, agenda item, keyword...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFixedHeight(28)
        self.search_input.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border-color: #0284C7;
            }
        """)
        self.search_input.textChanged.connect(self._apply_filters)
        filter_row.addWidget(self.search_input, stretch=2)

        filter_row.addWidget(QLabel("<b>Company:</b>"))
        self.company_combo = QComboBox()
        self.company_combo.setFixedHeight(28)
        self.company_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.company_combo, stretch=1)

        filter_row.addWidget(QLabel("<b>Type:</b>"))
        self.type_combo = QComboBox()
        self.type_combo.setFixedHeight(28)
        self.type_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.type_combo)

        filter_row.addWidget(QLabel("<b>Status:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.setFixedHeight(28)
        self.status_combo.currentIndexChanged.connect(self._apply_filters)
        filter_row.addWidget(self.status_combo)

        header_layout.addLayout(filter_row)
        main_layout.addWidget(header_card)

        # -------------------------------------------------------------
        # 2. MAIN TABBED VIEWER: [Table View] | [Markdown Dossier]
        # -------------------------------------------------------------
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabBar::tab {
                font-weight: bold;
                padding: 6px 16px;
            }
        """)

        # Tab 1: Interactive Table
        tab_table = QWidget()
        tab_table_layout = QVBoxLayout(tab_table)
        tab_table_layout.setContentsMargins(0, 8, 0, 0)
        tab_table_layout.setSpacing(8)

        # Splitter between Table and Detail Drawer
        self.splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "TDoc", "Type", "Status", "Source (Company)", "AI", "Title", "Technical Abstract / Remarks"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                gridline-color: #F1F5F9;
                font-size: 12px;
            }
            QTableWidget::item:selected {
                background-color: #E0F2FE;
                color: #0F172A;
            }
        """)

        h_hdr = self.table.horizontalHeader()
        h_hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(3, QHeaderView.Interactive)
        h_hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        h_hdr.setSectionResizeMode(5, QHeaderView.Stretch)
        h_hdr.setSectionResizeMode(6, QHeaderView.Stretch)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)

        self.splitter.addWidget(self.table)

        # Bottom Abstract / Solution Drawer
        self.drawer_frame = QFrame()
        self.drawer_frame.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
            }
        """)
        drawer_layout = QVBoxLayout(self.drawer_frame)
        drawer_layout.setContentsMargins(10, 8, 10, 8)
        drawer_layout.setSpacing(4)

        drawer_hdr = QHBoxLayout()
        self.drawer_title_lbl = QLabel("<b>Technical Details && Proposed Solution</b>")
        self.drawer_title_lbl.setStyleSheet("font-size: 12px; color: #1E293B;")
        drawer_hdr.addWidget(self.drawer_title_lbl)
        drawer_hdr.addStretch()

        self.btn_open_tdoc_drawer = QPushButton("📂 Open TDoc (.docx)")
        self.btn_open_tdoc_drawer.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_tdoc_drawer.setCursor(Qt.PointingHandCursor)
        self.btn_open_tdoc_drawer.clicked.connect(self._open_selected_tdoc)
        drawer_hdr.addWidget(self.btn_open_tdoc_drawer)

        drawer_layout.addLayout(drawer_hdr)

        self.drawer_text = QTextEdit()
        self.drawer_text.setReadOnly(True)
        self.drawer_text.setPlaceholderText("Select any contribution above to inspect its technical abstract, proposed solution, and secretary remarks...")
        self.drawer_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #F8FAFC;
                font-size: 12px;
                color: #334155;
                line-height: 1.4;
            }
        """)
        drawer_layout.addWidget(self.drawer_text)

        self.splitter.addWidget(self.drawer_frame)
        self.splitter.setSizes([450, 150])

        tab_table_layout.addWidget(self.splitter)
        self.tab_widget.addTab(tab_table, "📊 Interactive Contribution Table")

        # Tab 2: Markdown Dossier
        tab_md = QWidget()
        tab_md_layout = QVBoxLayout(tab_md)
        tab_md_layout.setContentsMargins(8, 8, 8, 8)
        tab_md_layout.setSpacing(8)

        self.md_browser = QTextEdit()
        self.md_browser.setReadOnly(True)
        self.md_browser.setFont(QFont("Consolas", 10))
        self.md_browser.setPlainText(self.markdown_report)
        self.md_browser.setStyleSheet("""
            QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                color: #1E293B;
            }
        """)
        tab_md_layout.addWidget(self.md_browser)
        self.tab_widget.addTab(tab_md, "📝 Technical Dossier (Markdown)")

        main_layout.addWidget(self.tab_widget, stretch=1)

        # -------------------------------------------------------------
        # 3. FOOTER ACTIONS
        # -------------------------------------------------------------
        footer_layout = QHBoxLayout()

        self.count_lbl = QLabel(f"Showing {len(self.filtered_entries)} of {len(self.all_entries)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 12px; color: #64748B;")
        footer_layout.addWidget(self.count_lbl)
        footer_layout.addStretch()

        self.btn_copy_md = QPushButton("📋 Copy Markdown Dossier")
        self.btn_copy_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_md.setCursor(Qt.PointingHandCursor)
        self.btn_copy_md.clicked.connect(self._copy_markdown)
        footer_layout.addWidget(self.btn_copy_md)

        self.btn_export_md = QPushButton("💾 Save Markdown (.md)")
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.setCursor(Qt.PointingHandCursor)
        self.btn_export_md.clicked.connect(self._export_markdown_file)
        footer_layout.addWidget(self.btn_export_md)

        self.btn_export_csv = QPushButton("📊 Export CSV")
        self.btn_export_csv.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_csv.setCursor(Qt.PointingHandCursor)
        self.btn_export_csv.clicked.connect(self._export_csv_file)
        footer_layout.addWidget(self.btn_export_csv)

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.close)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)

    def _populate_filters(self):
        """Populates the Company, Type, and Status dropdowns."""
        self.company_combo.blockSignals(True)
        self.type_combo.blockSignals(True)
        self.status_combo.blockSignals(True)

        self.company_combo.clear()
        self.type_combo.clear()
        self.status_combo.clear()

        # Companies
        companies = sorted({r.get("source", "").strip() for r in self.all_entries if r.get("source", "").strip()})
        self.company_combo.addItem("All Companies", "")
        for comp in companies:
            cnt = sum(1 for r in self.all_entries if r.get("source", "").strip() == comp)
            self.company_combo.addItem(f"{comp} ({cnt})", comp)

        # Types
        types = sorted({r.get("type", "").strip() for r in self.all_entries if r.get("type", "").strip()})
        self.type_combo.addItem("All Types", "")
        for t in types:
            cnt = sum(1 for r in self.all_entries if r.get("type", "").strip() == t)
            self.type_combo.addItem(f"{t} ({cnt})", t)

        # Statuses
        statuses = sorted({r.get("status", "").strip().lower() for r in self.all_entries if r.get("status", "").strip()})
        self.status_combo.addItem("All Statuses", "")
        for s in statuses:
            cnt = sum(1 for r in self.all_entries if r.get("status", "").strip().lower() == s)
            self.status_combo.addItem(f"{s.capitalize()} ({cnt})", s)

        self.company_combo.blockSignals(False)
        self.type_combo.blockSignals(False)
        self.status_combo.blockSignals(False)

    def _apply_filters(self):
        """Filters the entries based on search input and filter combos."""
        q = self.search_input.text().strip().lower()
        sel_comp = self.company_combo.currentData() or ""
        sel_type = self.type_combo.currentData() or ""
        sel_status = self.status_combo.currentData() or ""

        filtered = []
        for r in self.all_entries:
            if sel_comp and r.get("source", "").strip() != sel_comp:
                continue
            if sel_type and r.get("type", "").strip() != sel_type:
                continue
            if sel_status and r.get("status", "").strip().lower() != sel_status:
                continue

            if q:
                tokens = [
                    r.get("tdoc", ""),
                    r.get("title", ""),
                    r.get("source", ""),
                    r.get("agenda_item", ""),
                    r.get("ai_desc", ""),
                    r.get("abstract", ""),
                    r.get("wi", ""),
                    r.get("remarks", ""),
                ]
                match_text = " ".join(t.lower() for t in tokens if t)
                if not all(term in match_text for term in q.split()):
                    continue

            filtered.append(r)

        self.filtered_entries = filtered
        self._refresh_table()

    def _refresh_table(self):
        """Re-populates the table widget with filtered records."""
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)

        for row_idx, r in enumerate(self.filtered_entries):
            self.table.insertRow(row_idx)

            # 0. TDoc ID
            tdoc_str = r.get("tdoc", "")
            tdoc_item = QTableWidgetItem(tdoc_str)
            tdoc_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
            tdoc_item.setForeground(QBrush(QColor("#0369A1")))
            tdoc_item.setData(Qt.UserRole, r)
            self.table.setItem(row_idx, 0, tdoc_item)

            # 1. Type
            type_str = r.get("type") or "pCR"
            type_item = QTableWidgetItem(type_str)
            type_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 1, type_item)

            # 2. Status
            status_str = r.get("status") or "—"
            status_item = QTableWidgetItem(status_str.capitalize())
            status_item.setTextAlignment(Qt.AlignCenter)
            if "approv" in status_str or "agree" in status_str:
                status_item.setForeground(QBrush(QColor("#166534")))
            elif "noted" in status_str:
                status_item.setForeground(QBrush(QColor("#0369A1")))
            elif "postpone" in status_str or "revise" in status_str:
                status_item.setForeground(QBrush(QColor("#B45309")))
            self.table.setItem(row_idx, 2, status_item)

            # 3. Source (Company)
            source_item = QTableWidgetItem(r.get("source", ""))
            self.table.setItem(row_idx, 3, source_item)

            # 4. Agenda Item
            ai_item = QTableWidgetItem(r.get("agenda_item", ""))
            ai_item.setTextAlignment(Qt.AlignCenter)
            if r.get("ai_desc"):
                ai_item.setToolTip(f"AI {r.get('agenda_item')}: {r.get('ai_desc')}")
            self.table.setItem(row_idx, 4, ai_item)

            # 5. Title
            title_item = QTableWidgetItem(r.get("title", ""))
            self.table.setItem(row_idx, 5, title_item)

            # 6. Abstract / Remarks Preview
            abs_text = r.get("abstract") or r.get("remarks") or ""
            clean_abs = " ".join(abs_text.split())
            abs_item = QTableWidgetItem(clean_abs)
            abs_item.setToolTip(abs_text)
            self.table.setItem(row_idx, 6, abs_item)

        self.table.setSortingEnabled(True)
        self.count_lbl.setText(f"Showing {len(self.filtered_entries)} of {len(self.all_entries)} TDocs")

    def _on_selection_changed(self):
        """Updates the bottom drawer when a table row is selected."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            self.drawer_text.clear()
            self.drawer_title_lbl.setText("<b>Technical Details && Proposed Solution</b>")
            return

        row_idx = selected_rows[0].row()
        item = self.table.item(row_idx, 0)
        if not item:
            return

        r = item.data(Qt.UserRole)
        if not r:
            return

        tdoc = r.get("tdoc", "")
        title = r.get("title", "No Title")
        source = r.get("source", "Unknown")
        ai = r.get("agenda_item", "")
        ai_desc = r.get("ai_desc", "")
        wi = r.get("wi", "")
        status = r.get("status", "")
        meeting = r.get("meeting", "")
        abstract = r.get("abstract", "")
        remarks = r.get("remarks", "")
        contact = r.get("contact", "")

        self.drawer_title_lbl.setText(f"<b>{tdoc}: {html.escape(title)}</b>")

        html_parts = []
        meta_spans = []
        if source:
            meta_spans.append(f"<b>Source:</b> {html.escape(source)}")
        if meeting:
            meta_spans.append(f"<b>Meeting:</b> {html.escape(meeting)}")
        if ai:
            meta_spans.append(f"<b>Agenda Item:</b> {html.escape(ai)}{f' ({html.escape(ai_desc)})' if ai_desc else ''}")
        if status:
            meta_spans.append(f"<b>Status:</b> {html.escape(status.capitalize())}")
        if wi:
            meta_spans.append(f"<b>Work Item:</b> {html.escape(wi)}")
        if contact:
            meta_spans.append(f"<b>Contact:</b> {html.escape(contact)}")

        html_parts.append(f"<div style='margin-bottom: 8px; color: #475569;'>{' &nbsp;|&nbsp; '.join(meta_spans)}</div>")

        if abstract:
            html_parts.append("<div style='font-weight: bold; margin-bottom: 4px; color: #0F172A;'>Contribution Abstract / Technical Scope:</div>")
            html_parts.append(f"<div style='background-color: #F1F5F9; border-left: 3px solid #0284C7; padding: 6px 10px; margin-bottom: 8px; white-space: pre-wrap;'>{html.escape(abstract)}</div>")

        if remarks:
            html_parts.append("<div style='font-weight: bold; margin-bottom: 4px; color: #0F172A;'>Secretary Remarks / Outcome:</div>")
            html_parts.append(f"<div style='background-color: #FEF3C7; border-left: 3px solid #D97706; padding: 6px 10px; color: #92400E;'>{html.escape(remarks)}</div>")

        if not abstract and not remarks:
            html_parts.append("<div style='color: #94A3B8; font-style: italic;'>No extended abstract or secretary remarks in the meeting spreadsheet for this document.</div>")

        self.drawer_text.setHtml("".join(html_parts))

    def _on_table_double_clicked(self, item: QTableWidgetItem):
        """Double clicking any cell in a row opens the TDoc in Word."""
        row_idx = item.row()
        tdoc_item = self.table.item(row_idx, 0)
        if tdoc_item:
            r = tdoc_item.data(Qt.UserRole)
            if r and r.get("tdoc"):
                logger.info(f"📂 [TDoc Report] Double-click launch for TDoc {r['tdoc']}")
                SpecTDocOpener.open_tdoc(r["tdoc"], meeting_hint=r.get("meeting", ""), parent=self)

    def _open_selected_tdoc(self):
        """Opens the currently selected TDoc document."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return
        row_idx = selected_rows[0].row()
        item = self.table.item(row_idx, 0)
        if item:
            r = item.data(Qt.UserRole)
            if r and r.get("tdoc"):
                SpecTDocOpener.open_tdoc(r["tdoc"], meeting_hint=r.get("meeting", ""), parent=self)

    def _show_context_menu(self, pos: QPoint):
        """Right-click context menu for table rows."""
        item = self.table.itemAt(pos)
        if not item:
            return

        row_idx = item.row()
        tdoc_item = self.table.item(row_idx, 0)
        if not tdoc_item:
            return

        r = tdoc_item.data(Qt.UserRole)
        if not r or not r.get("tdoc"):
            return

        tdoc = r["tdoc"]
        meeting = r.get("meeting", "")
        title = r.get("title", "")

        menu = QMenu(self)

        act_open = menu.addAction(f"📂 Open TDoc Document ({tdoc}.docx)")
        act_open.triggered.connect(lambda: SpecTDocOpener.open_tdoc(tdoc, meeting_hint=meeting, parent=self))

        act_portal = menu.addAction(f"🌐 Open {tdoc} in 3GU Web Portal")
        base_id = tdoc.split("r")[0].split("R")[0]
        act_portal.triggered.connect(lambda: SpecTDocOpener.open_tdoc_in_portal(base_id))

        act_jump = menu.addAction(f"🚀 Jump to {tdoc} in Meetings Tab")
        act_jump.triggered.connect(lambda: SpecTDocOpener.jump_to_meetings_tab(tdoc, main_window=self._parent_ref))

        menu.addSeparator()

        act_citation = menu.addAction(f"📋 Copy Citation for {tdoc}")
        act_citation.triggered.connect(lambda: SpecTDocOpener.copy_tdoc_citation(
            tdoc_id=tdoc,
            spec_number=self.spec_number,
            version=self.version,
            meeting=meeting,
            subject=title,
            parent=self
        ))

        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def _copy_markdown(self):
        """Copies the Markdown dossier report to clipboard."""
        if not self.markdown_report:
            self.markdown_report = SpecTDocReportGenerator.generate_markdown_report(
                self.spec_number, self.version, self.all_entries, spec_title=self.spec_title
            )
        QApplication.clipboard().setText(self.markdown_report)
        self.btn_copy_md.setText("✅ Copied Dossier!")
        logger.info(f"📋 [TDoc Report] Copied Markdown dossier for TS {self.spec_number} v{self.version} to clipboard.")
        QTimer.singleShot(1500, lambda: self.btn_copy_md.setText("📋 Copy Markdown Dossier"))

    def _export_markdown_file(self):
        """Saves the Markdown dossier report to disk."""
        if not self.markdown_report:
            self.markdown_report = SpecTDocReportGenerator.generate_markdown_report(
                self.spec_number, self.version, self.all_entries, spec_title=self.spec_title
            )

        default_name = f"TS_{self.spec_number}_v{self.version}_TDoc_Dossier.md"
        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc Dossier Report",
            default_name,
            "Markdown Files (*.md);;All Files (*)"
        )
        if dest_path:
            try:
                Path(dest_path).write_text(self.markdown_report, encoding="utf-8")
                logger.info(f"💾 [TDoc Report] Saved Markdown dossier to: {dest_path}")
                QMessageBox.information(self, "Export Complete", f"Successfully exported dossier to:\n{dest_path}")
            except Exception as e:
                logger.error(f"Error saving markdown dossier: {e}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"Failed to save dossier:\n{e}")

    def _export_csv_file(self):
        """Exports the entries to a CSV file."""
        csv_text = SpecTDocReportGenerator.generate_csv_report(self.filtered_entries)
        default_name = f"TS_{self.spec_number}_v{self.version}_TDocs.csv"
        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc CSV Report",
            default_name,
            "CSV Files (*.csv);;All Files (*)"
        )
        if dest_path:
            try:
                Path(dest_path).write_text(csv_text, encoding="utf-8")
                logger.info(f"📊 [TDoc Report] Saved CSV report to: {dest_path}")
                QMessageBox.information(self, "Export Complete", f"Successfully exported CSV to:\n{dest_path}")
            except Exception as e:
                logger.error(f"Error saving CSV: {e}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"Failed to save CSV:\n{e}")
