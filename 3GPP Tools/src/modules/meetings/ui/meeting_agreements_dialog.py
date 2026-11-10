# --- File: src/modules/meetings/ui/meeting_agreements_dialog.py ---
"""
Meeting Agreements & Conclusions Dialog (`MeetingAgreementsDialog`).

Dedicated viewer and management console for Chairman's Notes agreements,
working assumptions, and conclusions for a single 3GPP meeting (e.g. RAN1/RAN2/RAN3/SA2).
Provides real-time search, category filtering, split-view reading pane,
1-click table localization, clipboard citation copying, and Markdown export.
"""
import html
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableView,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)

logger = logging.getLogger(__name__)


def _clean_agreement_preview(text: str, max_chars: int = 140) -> str:
    """Produces a clean, single-line preview from multi-line agreement text."""
    if not text:
        return ""
    # Strip leading tags and header labels
    t = re.sub(r'^(?:🟢|🟡|🔵)?\s*\[\s*(?:AGREEMENT|WORKING ASSUMPTION|CONCLUSION|DECISION|FFS)\s*\]\s*:?', '', text, flags=re.IGNORECASE)
    t = re.sub(r'^(?:Agreements?|Working\s+Assumptions?)\s*[:\n]?', '', t, flags=re.IGNORECASE)
    lines = [line.strip().lstrip('•-*1234567890. ') for line in t.splitlines() if line.strip()]
    joined = " ".join(lines)
    if len(joined) > max_chars:
        return joined[:max_chars].rstrip() + "..."
    return joined


class AgreementsTableModel(QAbstractTableModel):
    """Table model representing meeting agreements for the viewer dialog."""
    HEADERS = ["Kind", "TDoc", "Agenda Item", "Work Item", "Agreement Summary"]

    def __init__(self, agreements: List[dict], parent=None):
        super().__init__(parent)
        self._all_agreements: List[dict] = list(agreements or [])
        self._filtered_agreements: List[dict] = list(self._all_agreements)

    def rowCount(self, parent=QModelIndex()):
        return len(self._filtered_agreements)

    def columnCount(self, parent=QModelIndex()):
        return len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEADERS[section]
        return None

    def get_row(self, row: int) -> Optional[dict]:
        if 0 <= row < len(self._filtered_agreements):
            return self._filtered_agreements[row]
        return None

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid() or index.row() >= len(self._filtered_agreements):
            return None

        ag = self._filtered_agreements[index.row()]
        col = index.column()
        kind = str(ag.get("kind", "Agreement") or "Agreement").strip()
        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()

        if role == Qt.DisplayRole:
            if col == 0:
                if is_wa:
                    return "🟡 Working Assumption"
                elif is_ffs:
                    return "🔵 FFS / Open Issue"
                return "🟢 Agreement"
            elif col == 1:
                return ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
            elif col == 2:
                ai = str(ag.get("ai") or "").strip()
                return ai if ai else "-"
            elif col == 3:
                wi = str(ag.get("wi") or "").strip()
                return wi if wi and wi.lower() != "general" else "-"
            elif col == 4:
                return _clean_agreement_preview(ag.get("text", ""))

        elif role == Qt.BackgroundRole:
            if is_wa:
                return QColor("#FEFCE8")  # Amber 50
            elif is_ffs:
                return QColor("#EFF6FF")  # Blue 50
            return QColor("#F0FDF4")      # Emerald 50

        elif role == Qt.ForegroundRole:
            if col == 0:
                if is_wa:
                    return QColor("#B45309")
                elif is_ffs:
                    return QColor("#1D4ED8")
                return QColor("#15803D")
            elif col == 1:
                return QColor("#0F172A")

        elif role == Qt.TextAlignmentRole:
            if col in (0, 1, 2, 3):
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.ToolTipRole:
            txt = ag.get("text", "")
            return f"<div style='max-width: 450px; white-space: pre-wrap;'>{html.escape(txt)}</div>"

        return None

    def apply_filter(self, search_text: str = "", kind_filter: str = "ALL", ai_filter: str = "ALL"):
        """Filters agreements by search query, agreement kind, and agenda item."""
        self.beginResetModel()
        query = (search_text or "").lower().strip()
        kind_req = (kind_filter or "ALL").upper()
        ai_req = (ai_filter or "ALL").strip()

        filtered = []
        for ag in self._all_agreements:
            # 1. Kind filter
            k = str(ag.get("kind", "")).lower()
            if kind_req == "AGREEMENT" and ("working assumption" in k or "ffs" in k):
                continue
            elif kind_req == "WA" and "working assumption" not in k:
                continue
            elif kind_req == "FFS" and ("ffs" not in k and "open issue" not in k):
                continue

            # 2. AI filter
            if ai_req != "ALL":
                ai_val = str(ag.get("ai", "")).strip()
                if ai_val != ai_req:
                    continue

            # 3. Text search
            if query:
                searchable = " ".join([
                    str(ag.get("tdoc", "")),
                    str(ag.get("ai", "")),
                    str(ag.get("wi", "")),
                    str(ag.get("title", "")),
                    str(ag.get("company", "")),
                    str(ag.get("text", "")),
                ]).lower()
                if query not in searchable:
                    continue

            filtered.append(ag)

        self._filtered_agreements = filtered
        self.endResetModel()


class MeetingAgreementsDialog(QDialog):
    """
    Dedicated dialog providing structured viewing, searching, inspection,
    and export of all agreements in the current meeting.
    """
    locate_tdoc_requested = pyqtSignal(str)

    def __init__(self, agreements: List[dict], meeting_name: str = "", meeting_dir: Optional[Path] = None, parent=None):
        super().__init__(parent)
        self.raw_agreements = list(agreements or [])
        self.meeting_name = meeting_name or "3GPP Meeting"
        self.meeting_dir = Path(meeting_dir) if meeting_dir else None

        self.setWindowTitle(f"⚖️ Meeting Agreements & Conclusions — {self.meeting_name}")
        self.resize(1020, 680)
        self.setMinimumSize(800, 500)

        self.model = AgreementsTableModel(self.raw_agreements, parent=self)
        self._init_ui()
        self._update_kpi_counters()

        # Pre-select first item if available
        if self.model.rowCount() > 0:
            self.table.selectRow(0)
            self._on_table_row_selected(self.model.index(0, 0))

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # 1. Header Ribbon with Title & KPI Chips
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 6px 12px;
            }
        """)
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(4, 4, 4, 4)

        title_layout = QVBoxLayout()
        title_lbl = QLabel(f"<b>⚖️ Agreements & Working Assumptions</b> — {self.meeting_name}")
        title_lbl.setStyleSheet("font-size: 15px; color: #0F172A; font-weight: bold;")
        sub_lbl = QLabel("Chairman's Notes extracted agreements, decisions, and working assumptions")
        sub_lbl.setStyleSheet("font-size: 11px; color: #64748B;")
        title_layout.addWidget(title_lbl)
        title_layout.addWidget(sub_lbl)
        header_layout.addLayout(title_layout)

        header_layout.addStretch()

        # KPI Badges
        self.kpi_layout = QHBoxLayout()
        self.kpi_layout.setSpacing(8)
        header_layout.addLayout(self.kpi_layout)

        main_layout.addWidget(header_frame)

        # 2. Search & Filter Bar
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search agreements text, TDoc ID, AI, WI, or company...")
        self.search_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.search_input, stretch=3)

        # Kind Filter Pills
        self.combo_kind = QComboBox()
        self.combo_kind.addItem("All Kinds", "ALL")
        self.combo_kind.addItem("🟢 Agreements Only", "AGREEMENT")
        self.combo_kind.addItem("🟡 Working Assumptions Only", "WA")
        self.combo_kind.addItem("🔵 FFS / Open Issues Only", "FFS")
        self.combo_kind.setStyleSheet("""
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
                font-weight: 500;
                color: #1E293B;
            }
        """)
        self.combo_kind.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_kind)

        # AI Filter Combo
        self.combo_ai = QComboBox()
        self.combo_ai.addItem("All Agenda Items", "ALL")
        unique_ais = sorted({str(ag.get("ai", "")).strip() for ag in self.raw_agreements if ag.get("ai")})
        for ai in unique_ais:
            self.combo_ai.addItem(f"AI {ai}", ai)
        self.combo_ai.setStyleSheet(self.combo_kind.styleSheet())
        self.combo_ai.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_ai)

        main_layout.addLayout(filter_layout)

        # 3. Master-Detail Splitter
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)

        # Master: Agreements Table
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.setAlternatingRowColors(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableView {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                gridline-color: #F1F5F9;
                selection-background-color: #DBEAFE;
                selection-color: #1E3A8A;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #E2E8F0;
            }
        """)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 150)
        header.resizeSection(1, 110)
        header.resizeSection(2, 90)
        header.resizeSection(3, 130)
        header.setSectionResizeMode(4, QHeaderView.Stretch)

        self.table.selectionModel().selectionChanged.connect(self._on_selection_changed)
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        splitter.addWidget(self.table)

        # Detail: Rich Viewer Pane
        detail_container = QFrame()
        detail_container.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(10, 8, 10, 8)
        detail_layout.setSpacing(6)

        detail_bar = QHBoxLayout()
        self.detail_title_lbl = QLabel("<b>Agreement Details</b>")
        self.detail_title_lbl.setStyleSheet("font-size: 13px; color: #0F172A;")
        detail_bar.addWidget(self.detail_title_lbl)
        detail_bar.addStretch()

        self.btn_copy_single = QPushButton("📋 Copy Agreement")
        self.btn_copy_single.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_single.setToolTip("Copy this agreement text with formal citation to clipboard.")
        self.btn_copy_single.clicked.connect(self._copy_current_agreement)
        detail_bar.addWidget(self.btn_copy_single)

        self.btn_locate_table = QPushButton("🎯 Show in Meeting Table")
        self.btn_locate_table.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_locate_table.setToolTip("Scroll to and highlight this TDoc in the main meeting window.")
        self.btn_locate_table.clicked.connect(self._locate_current_tdoc)
        detail_bar.addWidget(self.btn_locate_table)

        detail_layout.addLayout(detail_bar)

        self.viewer = QTextBrowser()
        self.viewer.setOpenExternalLinks(True)
        self.viewer.setStyleSheet("""
            QTextBrowser {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                padding: 10px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                color: #0F172A;
            }
        """)
        detail_layout.addWidget(self.viewer)

        splitter.addWidget(detail_container)
        splitter.setSizes([260, 260])
        main_layout.addWidget(splitter, stretch=1)

        # 4. Bottom Action Bar
        bottom_layout = QHBoxLayout()

        self.btn_copy_all = QPushButton("📋 Copy All (Markdown)")
        self.btn_copy_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_all.setToolTip("Copy all currently visible agreements formatted as Markdown.")
        self.btn_copy_all.clicked.connect(self._copy_all_markdown)
        bottom_layout.addWidget(self.btn_copy_all)

        self.btn_export_md = QPushButton("💾 Export to Agreements.md...")
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.setToolTip("Save agreements to a standalone Markdown (.md) document.")
        self.btn_export_md.clicked.connect(self._export_to_markdown_file)
        bottom_layout.addWidget(self.btn_export_md)

        bottom_layout.addStretch()

        self.status_count_lbl = QLabel("")
        self.status_count_lbl.setStyleSheet("font-size: 11px; color: #64748B; font-weight: bold;")
        bottom_layout.addWidget(self.status_count_lbl)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_close.clicked.connect(self.accept)
        bottom_layout.addWidget(btn_close)

        main_layout.addLayout(bottom_layout)
        self._update_status_count()

    def _update_kpi_counters(self):
        """Builds visually appealing KPI chip badges at the top of the dialog."""
        # Clear existing
        while self.kpi_layout.count():
            item = self.kpi_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        tot = len(self.raw_agreements)
        ag_cnt = sum(1 for a in self.raw_agreements if "working assumption" not in str(a.get("kind", "")).lower() and "ffs" not in str(a.get("kind", "")).lower())
        wa_cnt = sum(1 for a in self.raw_agreements if "working assumption" in str(a.get("kind", "")).lower())
        ffs_cnt = sum(1 for a in self.raw_agreements if "ffs" in str(a.get("kind", "")).lower() or "open issue" in str(a.get("kind", "")).lower())
        ai_cnt = len({str(a.get("ai", "")).strip() for a in self.raw_agreements if a.get("ai")})

        def _make_chip(label: str, count: int, bg: str, fg: str, border: str) -> QLabel:
            lbl = QLabel(f"<b>{count}</b> {label}")
            lbl.setStyleSheet(f"""
                QLabel {{
                    background-color: {bg};
                    color: {fg};
                    border: 1px solid {border};
                    border-radius: 4px;
                    padding: 3px 8px;
                    font-size: 11px;
                }}
            """)
            return lbl

        self.kpi_layout.addWidget(_make_chip("Total", tot, "#F1F5F9", "#334155", "#CBD5E1"))
        if ag_cnt > 0:
            self.kpi_layout.addWidget(_make_chip("Agreements", ag_cnt, "#DCFCE7", "#166534", "#86EFAC"))
        if wa_cnt > 0:
            self.kpi_layout.addWidget(_make_chip("Working Assumptions", wa_cnt, "#FEF3C7", "#92400E", "#FCD34D"))
        if ffs_cnt > 0:
            self.kpi_layout.addWidget(_make_chip("FFS / Open", ffs_cnt, "#DBEAFE", "#1E40AF", "#93C5FD"))
        if ai_cnt > 0:
            self.kpi_layout.addWidget(_make_chip("AIs Covered", ai_cnt, "#F3E8FF", "#6B21A8", "#D8B4FE"))

    def _on_filter_changed(self):
        query = self.search_input.text()
        kind = self.combo_kind.currentData()
        ai = self.combo_ai.currentData()
        self.model.apply_filter(search_text=query, kind_filter=kind, ai_filter=ai)
        self._update_status_count()
        if self.model.rowCount() > 0:
            self.table.selectRow(0)
            self._on_table_row_selected(self.model.index(0, 0))
        else:
            self.viewer.setHtml("<p style='color: #64748B;'><i>No agreements match active search filters.</i></p>")

    def _update_status_count(self):
        curr = self.model.rowCount()
        tot = len(self.raw_agreements)
        self.status_count_lbl.setText(f"Showing {curr} of {tot} agreements")

    def _on_selection_changed(self, selected, deselected):
        indexes = self.table.selectionModel().selectedRows()
        if indexes:
            self._on_table_row_selected(indexes[0])

    def _on_table_row_selected(self, index: QModelIndex):
        ag = self.model.get_row(index.row())
        if not ag:
            return

        tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
        ai = ag.get("ai") or "-"
        wi = ag.get("wi") or "-"
        title = ag.get("title") or ""
        company = ag.get("company") or ""
        kind = str(ag.get("kind", "Agreement") or "Agreement").strip()
        file_src = ag.get("file") or ""
        text = ag.get("text", "")

        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()

        if is_wa:
            tag_bg, tag_fg, tag_border = "#FEF3C7", "#92400E", "#FCD34D"
            tag_label = "🟡 WORKING ASSUMPTION"
        elif is_ffs:
            tag_bg, tag_fg, tag_border = "#DBEAFE", "#1E40AF", "#93C5FD"
            tag_label = "🔵 FFS / OPEN ISSUE"
        else:
            tag_bg, tag_fg, tag_border = "#DCFCE7", "#166534", "#86EFAC"
            tag_label = "🟢 3GPP TECHNICAL AGREEMENT"

        self.detail_title_lbl.setText(f"<b>Agreement Details:</b> <span style='color: #2563EB;'>{html.escape(tdoc)}</span> (AI {html.escape(ai)})")

        meta_rows = []
        if tdoc and tdoc != "AI Level":
            meta_rows.append(f"<b>TDoc:</b> <span style='font-family: monospace; font-size: 13px; font-weight: bold;'>{html.escape(tdoc)}</span>")
        if ai and ai != "-":
            meta_rows.append(f"<b>Agenda Item:</b> {html.escape(ai)}")
        if wi and wi != "-" and wi.lower() != "general":
            meta_rows.append(f"<b>Work Item:</b> <code>{html.escape(wi)}</code>")
        if title:
            meta_rows.append(f"<b>Contribution Title:</b> {html.escape(title)}")
        if company:
            meta_rows.append(f"<b>Source Company:</b> {html.escape(company)}")
        if file_src:
            meta_rows.append(f"<b>Notes Document:</b> {html.escape(Path(file_src).name)}")

        meta_html = " &nbsp;•&nbsp; ".join(meta_rows)

        # Format agreement text nicely
        escaped_txt = html.escape(text)

        html_content = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.5; color: #0F172A;">
            <div style="margin-bottom: 8px;">
                <span style="background-color: {tag_bg}; color: {tag_fg}; border: 1px solid {tag_border}; border-radius: 4px; padding: 2px 8px; font-weight: bold; font-size: 11px;">
                    {tag_label}
                </span>
            </div>
            <div style="background-color: #F1F5F9; border-radius: 4px; padding: 6px 10px; margin-bottom: 12px; font-size: 11px; color: #334155;">
                {meta_html}
            </div>
            <div style="background-color: #FFFFFF; border: 1px solid #CBD5E1; border-left: 4px solid {tag_fg}; border-radius: 4px; padding: 12px 14px; font-size: 13px; white-space: pre-wrap; font-family: 'Consolas', 'Segoe UI', monospace;">{escaped_txt}</div>
        </body>
        </html>
        """
        self.viewer.setHtml(html_content)

    def _on_table_double_clicked(self, index: QModelIndex):
        """Double clicking a row locates it in the parent meeting table."""
        self._locate_current_tdoc()

    def _locate_current_tdoc(self):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return
        ag = self.model.get_row(indexes[0].row())
        if not ag:
            return
        tdoc = ag.get("tdoc") or ag.get("source_tdoc")
        if tdoc and tdoc != "AI Level":
            self.locate_tdoc_requested.emit(tdoc)
            parent = self.parent()
            if parent and hasattr(parent, "locate_tdoc"):
                parent.locate_tdoc(tdoc)

    def _copy_current_agreement(self):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return
        ag = self.model.get_row(indexes[0].row())
        if not ag:
            return

        tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
        ai = ag.get("ai") or ""
        wi = ag.get("wi") or ""
        kind = ag.get("kind") or "Agreement"
        text = ag.get("text", "")

        citation_header = f"[{kind}] {self.meeting_name}"
        if ai:
            citation_header += f" | AI {ai}"
        if wi and wi.lower() != "general":
            citation_header += f" ({wi})"
        if tdoc and tdoc != "AI Level":
            citation_header += f" | {tdoc}"

        full_copy = f"{citation_header}:\n{text}"
        QApplication.clipboard().setText(full_copy)
        self.btn_copy_single.setText("✓ Copied!")
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(2000, lambda: self.btn_copy_single.setText("📋 Copy Agreement"))

    def _copy_all_markdown(self):
        md_text = self._build_markdown_content()
        if not md_text:
            QMessageBox.information(self, "No Agreements", "No agreements to copy.")
            return
        QApplication.clipboard().setText(md_text)
        self.btn_copy_all.setText("✓ Copied All!")
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(2500, lambda: self.btn_copy_all.setText("📋 Copy All (Markdown)"))

    def _build_markdown_content(self) -> str:
        """Builds a structured GFM Markdown document containing all currently filtered agreements."""
        rows = self.model._filtered_agreements
        if not rows:
            return ""

        lines = [
            f"# {self.meeting_name} — Agreements & Working Assumptions",
            f"*Generated by 3GPP Tools • Total Agreements: {len(rows)}*",
            "",
            "---",
            "",
        ]

        # Group by Agenda Item
        by_ai: Dict[str, List[dict]] = {}
        for ag in rows:
            ai = str(ag.get("ai") or "General").strip()
            by_ai.setdefault(ai, []).append(ag)

        for ai, ag_list in by_ai.items():
            lines.append(f"## Agenda Item {ai}")
            lines.append("")

            for ag in ag_list:
                kind = ag.get("kind", "Agreement")
                tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
                wi = ag.get("wi", "")
                title = ag.get("title", "")
                company = ag.get("company", "")
                text = ag.get("text", "").strip()

                meta_parts = []
                if tdoc and tdoc != "AI Level":
                    meta_parts.append(f"**TDoc:** `{tdoc}`")
                if wi and wi.lower() != "general":
                    meta_parts.append(f"**WI:** `{wi}`")
                if company:
                    meta_parts.append(f"**Source:** {company}")

                badge = "🟡 WORKING ASSUMPTION" if "working assumption" in str(kind).lower() else ("🔵 FFS" if "ffs" in str(kind).lower() else "🟢 AGREEMENT")
                lines.append(f"### {badge}: {tdoc}")
                if meta_parts:
                    lines.append(f"> {' • '.join(meta_parts)}")
                    lines.append(">")
                if title:
                    lines.append(f"> *{title}*")
                    lines.append("")

                # Format agreement block
                for t_line in text.splitlines():
                    lines.append(f"    {t_line}")
                lines.append("")

            lines.append("---")
            lines.append("")

        return "\n".join(lines).strip()

    def _export_to_markdown_file(self):
        md_text = self._build_markdown_content()
        if not md_text:
            QMessageBox.information(self, "No Agreements", "No agreements to export.")
            return

        default_dir = str(self.meeting_dir / "Export") if self.meeting_dir else ""
        default_file = f"{self.meeting_name.replace(' ', '_')}_Agreements.md"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Meeting Agreements to Markdown",
            str(Path(default_dir) / default_file) if default_dir else default_file,
            "Markdown Files (*.md);;All Files (*.*)",
        )
        if save_path:
            try:
                Path(save_path).parent.mkdir(parents=True, exist_ok=True)
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(md_text)
                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Successfully exported {self.model.rowCount()} agreement(s) to:\n{save_path}",
                )
            except Exception as e:
                logger.error(f"Failed to export agreements markdown: {e}")
                QMessageBox.warning(self, "Export Failed", f"Could not save file:\n{e}")

    def focus_tdoc(self, tdoc_id: str):
        """Highlights and selects the agreement corresponding to a given TDoc ID."""
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()
        for row_idx, ag in enumerate(self.model._filtered_agreements):
            t = str(ag.get("tdoc") or ag.get("source_tdoc") or "").strip().upper()
            if t == tdoc_clean:
                self.table.selectRow(row_idx)
                self.table.scrollTo(self.model.index(row_idx, 0))
                self._on_table_row_selected(self.model.index(row_idx, 0))
                return
