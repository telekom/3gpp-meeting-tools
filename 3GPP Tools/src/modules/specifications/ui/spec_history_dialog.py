# --- File: src/modules/specifications/ui/spec_history_dialog.py ---
"""
Specification Change History & Integrated TDocs Viewer Dialog.

Displays chronological specification revisions, formal TSG CRs, pre-release p-CRs,
and rapporteur notes. Allows direct 1-click opening of referenced TDoc documents
(.docx / .zip), cross-referencing to meetings, and batch clipboard export.
"""

import html
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QPoint, Qt, QUrl
from PyQt5.QtGui import QBrush, QColor, QCursor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    COMBOBOX_STYLE_TOOLBAR,
    SEARCH_INPUT_STYLE,
    TABLE_STYLE_CLEAN,
)
from core.utils.paths import get_project_root
from modules.meetings.ui.tdocs_menus import _copy_to_clipboard
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.core.spec_history_extractor import (
    RE_TDOC_TOKEN,
    SpecHistoryExtractThread,
    SpecHistoryExtractor,
)
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
from modules.specifications.core.spec_tdoc_report import (
    SpecTDocReportGenerator,
    SpecTDocReportThread,
)
from modules.specifications.ui.spec_tdoc_report_dialog import SpecTDocReportDialog

logger = logging.getLogger(__name__)


class NumericSortItem(QTableWidgetItem):
    """Sorts table rows numerically when possible, falling back to natural string sorting."""

    def __lt__(self, other):
        try:
            return int(self.text().strip()) < int(other.text().strip())
        except (ValueError, TypeError):
            return super().__lt__(other)


class SpecVersionHistoryDialog(QDialog):
    """Modern modal dialog inspecting all Change History and integrated TDocs for a specification."""

    def __init__(
        self,
        spec_number: str,
        initial_version: Optional[str] = None,
        parent: Optional[QWidget] = None,
        history_db: Optional[SpecHistoryDatabase] = None,
        specs_db: Optional[Any] = None,
        initial_search: str = "",
        initial_clause: str = "",
        sort_by_origin_relevance: bool = False,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER (prevents dialog from being forced always-on-top)
        self._parent_ref = parent
        self.spec_number = spec_number.strip()
        self.initial_version = initial_version.lstrip("vV").strip() if initial_version else ""
        self.initial_search = initial_search.strip()
        self.initial_clause = initial_clause.strip()
        self.sort_by_origin_relevance = sort_by_origin_relevance
        self.history_db = history_db or SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
        self.specs_db = specs_db
        self._attribution_engine = None

        # Modeless top-level independent window configuration
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass

        self.setWindowTitle(f"Change History & Integrated TDocs — {self.spec_number}")
        self.resize(1060, 680)
        self.setMinimumSize(850, 500)

        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        self.all_records: List[Dict[str, Any]] = []
        self.filtered_records: List[Dict[str, Any]] = []
        self._active_report_thread = None
        self._report_dialog = None

        logger.info(f"📜 [Change History] Opened Change History & TDocs browser for TS {self.spec_number}{f' (v{self.initial_version})' if self.initial_version else ''}.")
        self._setup_ui()
        self._load_data(auto_extract_if_empty=True)

    def load_spec(
        self,
        spec_number: str,
        version: Optional[str] = None,
        initial_search: str = "",
        initial_clause: str = "",
        sort_by_origin_relevance: bool = False
    ):
        """Switches the active specification and reloads change history records."""
        self.spec_number = spec_number.strip()
        self.initial_version = version.lstrip("vV").strip() if version else ""
        if initial_search:
            self.initial_search = initial_search.strip()
        if initial_clause:
            self.initial_clause = initial_clause.strip()
        if sort_by_origin_relevance:
            self.sort_by_origin_relevance = sort_by_origin_relevance
        self.setWindowTitle(f"Change History & Integrated TDocs — {self.spec_number}")

        spec_title = "Specification Change History"
        if self.specs_db and hasattr(self.specs_db, "get_spec_details"):
            try:
                details = self.specs_db.get_spec_details(self.spec_number)
                if details and details.get("title"):
                    spec_title = details["title"]
            except Exception:
                pass
        self.title_lbl.setText(spec_title)
        self.spec_badge.setText(f"<b>{self.spec_number}</b>")
        logger.info(f"📜 [Change History] Reloading browser for TS {self.spec_number}{f' (v{self.initial_version})' if self.initial_version else ''}.")
        self._load_data(auto_extract_if_empty=True)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # -------------------------------------------------------------
        # 1. HEADER CARD (Spec badge, Title, Version Filter, Search Box)
        # -------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        self.spec_badge = QLabel(f"<b>{self.spec_number}</b>")
        self.spec_badge.setStyleSheet(BADGE_STYLE_PRIMARY + " font-size: 14px; padding: 3px 10px;")

        # Fetch formal title if available
        spec_title = "Specification Change History"
        if self.specs_db and hasattr(self.specs_db, "get_spec_details"):
            try:
                details = self.specs_db.get_spec_details(self.spec_number)
                if details and details.get("title"):
                    spec_title = details["title"]
            except Exception:
                pass

        self.title_lbl = QLabel(spec_title)
        self.title_lbl.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B;")
        self.title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top_row.addWidget(self.spec_badge)
        top_row.addWidget(self.title_lbl)
        top_row.addStretch()

        # TDoc Metadata Report Action Button
        self.btn_tdoc_report = QPushButton("📑 TDoc Dossier Report ▾")
        self.btn_tdoc_report.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_tdoc_report.setCursor(Qt.PointingHandCursor)
        self.btn_tdoc_report.setToolTip("Generate an on-demand technical report showing what all referenced TDocs are about, their authors, and abstracts.")
        report_menu = QMenu(self)
        report_menu.setToolTipsVisible(True)

        act_ver = report_menu.addAction("📑 Report for Active Version")
        act_ver.setToolTip("Extract metadata and generate report for all TDocs in the currently selected version.")
        act_ver.triggered.connect(lambda: self._generate_tdoc_report(scope="version"))

        act_sel = report_menu.addAction("📑 Report for Selected Rows")
        act_sel.setToolTip("Extract metadata and generate report for TDocs in highlighted table rows.")
        act_sel.triggered.connect(lambda: self._generate_tdoc_report(scope="selection"))

        act_all = report_menu.addAction("📑 Full Specification Dossier (All Versions)")
        act_all.setToolTip("Extract metadata and generate report for all TDocs across all indexed releases of this specification.")
        act_all.triggered.connect(lambda: self._generate_tdoc_report(scope="all"))

        self.btn_tdoc_report.setMenu(report_menu)
        top_row.addWidget(self.btn_tdoc_report)

        self.rescan_btn = QPushButton("🔄 Re-scan from Local Doc")
        self.rescan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.rescan_btn.setCursor(Qt.PointingHandCursor)
        self.rescan_btn.setToolTip("Re-extracts Change History from the latest local Word document.")
        self.rescan_btn.clicked.connect(self._on_rescan_clicked)
        top_row.addWidget(self.rescan_btn)

        header_layout.addLayout(top_row)

        # Controls Row: Version selector + Live search filter
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        ver_lbl = QLabel("<b>Version:</b>")
        ver_lbl.setStyleSheet("color: #475569; font-size: 12px;")

        self.version_combo = QComboBox()
        self.version_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.version_combo.setMinimumWidth(160)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_input.setPlaceholderText("🔍 Filter by TDoc (e.g. S2-2606625), CR #, meeting, or text...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._apply_filters)

        self.stats_lbl = QLabel("0 record(s)")
        self.stats_lbl.setStyleSheet("color: #64748B; font-size: 12px; font-weight: bold;")

        filter_row.addWidget(ver_lbl)
        filter_row.addWidget(self.version_combo)
        filter_row.addWidget(self.search_input, stretch=1)
        filter_row.addWidget(self.stats_lbl)

        header_layout.addLayout(filter_row)

        # Controls Row 2: Affected Clause Filter + Origin CR Relevance Sorting
        clause_row = QHBoxLayout()
        clause_row.setSpacing(8)

        self.chk_clause_filter = QCheckBox("🎯 Affected Clause:")
        self.chk_clause_filter.setStyleSheet("font-size: 11px; font-weight: 600; color: #334155;")
        self.chk_clause_filter.toggled.connect(self._apply_filters)

        self.clause_input = QLineEdit()
        self.clause_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.clause_input.setPlaceholderText("Clause # (e.g. 5.15.11)...")
        self.clause_input.setFixedWidth(160)
        self.clause_input.setClearButtonEnabled(True)
        self.clause_input.textChanged.connect(self._apply_filters)

        self.clause_mode_combo = QComboBox()
        self.clause_mode_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.clause_mode_combo.addItems([
            "Rank by Clause (Additive)",
            "Strict Clause Filter"
        ])
        self.clause_mode_combo.setToolTip(
            "Rank by Clause (Additive): Boosts matching CRs to the top while preserving CRs without cover sheets.\n"
            "Strict Clause Filter: Hides CRs that do not explicitly list this clause."
        )
        self.clause_mode_combo.currentIndexChanged.connect(self._apply_filters)

        self.chk_origin_ranking = QCheckBox("⭐ Origin CR Ranking")
        self.chk_origin_ranking.setStyleSheet("font-size: 11px; font-weight: bold; color: #0284C7;")
        self.chk_origin_ranking.setToolTip("Ranks candidate CRs by origin relevance using Category B, Title/Intro keywords, Affected Clauses, and chronological release primacy.")
        self.chk_origin_ranking.toggled.connect(self._apply_filters)

        clause_row.addWidget(self.chk_clause_filter)
        clause_row.addWidget(self.clause_input)
        clause_row.addWidget(self.clause_mode_combo)
        clause_row.addSpacing(12)
        clause_row.addWidget(self.chk_origin_ranking)
        clause_row.addStretch()

        header_layout.addLayout(clause_row)

        main_layout.addWidget(header_card)

        # -------------------------------------------------------------
        # 2. SPLITTER: TABLE (TOP) + DETAILS / CHIPS DRAWER (BOTTOM)
        # -------------------------------------------------------------
        splitter = QSplitter(Qt.Vertical)
        splitter.setHandleWidth(6)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #E2E8F0;
                border-radius: 2px;
            }
            QSplitter::handle:hover {
                background-color: #0284C7;
            }
        """)

        # Table Container Card
        table_card = QFrame()
        table_card.setObjectName("cardFrame")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(2, 2, 2, 2)

        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table.verticalHeader().setVisible(False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)
        self.table.itemDoubleClicked.connect(self._on_row_double_clicked)

        table_layout.addWidget(self.table)
        splitter.addWidget(table_card)

        # Details / Rapporteur Notes / Interactive TDoc View Card
        self.details_card = QFrame()
        self.details_card.setObjectName("cardFrame")
        details_layout = QVBoxLayout(self.details_card)
        details_layout.setContentsMargins(12, 10, 12, 10)
        details_layout.setSpacing(6)

        details_header = QHBoxLayout()
        self.drawer_title_lbl = QLabel("<b>Selected Entry Details:</b>")
        self.drawer_title_lbl.setStyleSheet("font-size: 12px; color: #334155;")

        self.btn_copy_entry_tdocs = QPushButton("📋 Copy Entry TDocs")
        self.btn_copy_entry_tdocs.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_entry_tdocs.setCursor(Qt.PointingHandCursor)
        self.btn_copy_entry_tdocs.setToolTip("Copies all TDoc identifiers for the selected entry to the clipboard.")
        self.btn_copy_entry_tdocs.clicked.connect(self._copy_selected_entry_tdocs)
        self.btn_copy_entry_tdocs.setEnabled(False)

        self.btn_open_primary = QPushButton("📄 Open TDoc")
        self.btn_open_primary.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_primary.setCursor(Qt.PointingHandCursor)
        self.btn_open_primary.setToolTip("Open the primary TDoc Word document or ZIP archive.")
        self.btn_open_primary.clicked.connect(self._open_selected_tdoc)
        self.btn_open_primary.setEnabled(False)

        details_header.addWidget(self.drawer_title_lbl)
        details_header.addStretch()
        details_header.addWidget(self.btn_copy_entry_tdocs)
        details_header.addWidget(self.btn_open_primary)
        details_layout.addLayout(details_header)

        # Unified Interactive Notes & TDoc Browser
        self.notes_browser = QTextBrowser()
        self.notes_browser.setReadOnly(True)
        self.notes_browser.setOpenLinks(False)
        self.notes_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                color: #1E293B;
                padding: 6px 8px;
            }
        """)
        self.notes_browser.anchorClicked.connect(self._on_notes_anchor_clicked)
        self.notes_browser.setContextMenuPolicy(Qt.CustomContextMenu)
        self.notes_browser.customContextMenuRequested.connect(self._show_notes_context_menu)
        details_layout.addWidget(self.notes_browser)

        splitter.addWidget(self.details_card)
        splitter.setSizes([400, 200])
        main_layout.addWidget(splitter, stretch=1)

        # -------------------------------------------------------------
        # 3. ACTION BUTTONS FOOTER
        # -------------------------------------------------------------
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(8)

        self.btn_copy_list = QPushButton("📋 Copy TDocs List")
        self.btn_copy_list.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_list.setCursor(Qt.PointingHandCursor)
        self.btn_copy_list.setToolTip("Copies all TDoc identifiers for the active version filter as a comma-separated list.")
        self.btn_copy_list.clicked.connect(self._copy_tdocs_list)

        self.btn_open_doc = QPushButton("📂 Open Spec Document")
        self.btn_open_doc.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_doc.setCursor(Qt.PointingHandCursor)
        self.btn_open_doc.setToolTip("Opens the baseline specification Word document.")
        self.btn_open_doc.clicked.connect(self._open_spec_document)

        self.btn_footer_report = QPushButton("📑 TDoc Dossier Report...")
        self.btn_footer_report.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_footer_report.setCursor(Qt.PointingHandCursor)
        self.btn_footer_report.setToolTip("Generate technical dossier report for referenced TDocs.")
        self.btn_footer_report.clicked.connect(lambda: self._generate_tdoc_report(scope="version"))

        footer_layout.addWidget(self.btn_copy_list)
        footer_layout.addWidget(self.btn_open_doc)
        footer_layout.addWidget(self.btn_footer_report)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)

        # Preset values now that all widgets are fully constructed
        if self.initial_search:
            self.search_input.setText(self.initial_search)
        if self.initial_clause:
            self.clause_input.setText(self.initial_clause)
            self.chk_clause_filter.setChecked(True)
        if self.sort_by_origin_relevance:
            self.chk_origin_ranking.setChecked(True)

    def _load_data(self, auto_extract_if_empty: bool = True):
        """Loads change history from 3gpp_spec_history.db, auto-extracting from local docx if empty."""
        records = self.history_db.get_spec_history(self.spec_number)

        if not records and auto_extract_if_empty:
            # Check if local docx exists and trigger fast extraction
            doc_file = SpecHistoryExtractor.resolve_spec_document(self.spec_number)
            if doc_file and doc_file.exists():
                extracted = SpecHistoryExtractor.extract_from_docx(doc_file, self.spec_number)
                if extracted:
                    self.history_db.upsert_history_records(extracted)
                    records = self.history_db.get_spec_history(self.spec_number)

        self.all_records = records

        # Populate version combo box
        self.version_combo.blockSignals(True)
        self.version_combo.clear()
        self.version_combo.addItem("All Versions", userData="")

        version_counts: Dict[str, int] = {}
        for r in self.all_records:
            v = r.get("version", "")
            if v:
                version_counts[v] = version_counts.get(v, 0) + 1

        # Natural descending sort of version strings
        def ver_key(v):
            return [(0, int(x)) if x.isdigit() else (1, x) for x in v.split(".")]

        sorted_vers = sorted(version_counts.keys(), key=ver_key, reverse=True)
        target_idx = 0
        for idx, ver in enumerate(sorted_vers):
            cnt = version_counts[ver]
            self.version_combo.addItem(f"v{ver} ({cnt} entry)", userData=ver)
            if self.initial_version and ver == self.initial_version:
                target_idx = idx + 1

        if self.sort_by_origin_relevance and not self.initial_version:
            target_idx = 0
        self.version_combo.setCurrentIndex(target_idx)
        self.version_combo.blockSignals(False)

        self._apply_filters()

    def _on_version_changed(self):
        self._apply_filters()

    def _apply_filters(self):
        """Applies version dropdown filter, clause filter, and origin ranking."""
        if not hasattr(self, "table") or not hasattr(self, "version_combo") or not hasattr(self, "search_input"):
            return
        target_ver = self.version_combo.currentData() or ""
        search_txt = self.search_input.text().strip()
        search_lower = search_txt.lower()
        clause_val = self.clause_input.text().strip() if hasattr(self, "clause_input") and self.chk_clause_filter.isChecked() else ""
        is_strict_clause = bool(clause_val) and hasattr(self, "clause_mode_combo") and self.clause_mode_combo.currentText() == "Strict Clause Filter"
        is_origin_ranking = hasattr(self, "chk_origin_ranking") and self.chk_origin_ranking.isChecked()

        if is_origin_ranking and (search_txt or clause_val):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            if not self._attribution_engine:
                self._attribution_engine = SpecClauseAttributionEngine(history_db=self.history_db)

            hint_clauses = [clause_val] if clause_val else None
            ranked = self._attribution_engine.rank_origin_crs_for_query(
                spec_number=self.spec_number,
                query=search_txt or clause_val,
                target_version=target_ver or None,
                affected_clauses_hint=hint_clauses,
            )

            # Apply version filter if explicitly selected
            if target_ver:
                ranked = [r for r in ranked if r.get("version") == target_ver]

            # Apply strict clause filter if enabled
            if is_strict_clause:
                import re as _re
                ranked = [
                    r for r in ranked
                    if clause_val in r.get("affected_clauses", "") or _re.search(rf"\b{_re.escape(clause_val)}\b", r.get("subject", ""))
                ]

            filtered = ranked
        else:
            filtered = []
            import re as _re
            for r in self.all_records:
                # 1. Version filter
                if target_ver and r.get("version") != target_ver:
                    continue

                # 2. Strict Clause filter
                if is_strict_clause:
                    aff = r.get("affected_clauses", "")
                    subj = r.get("subject", "")
                    if not (clause_val in aff or _re.search(rf"\b{_re.escape(clause_val)}\b", subj)):
                        continue

                # 3. Substring filter
                if search_lower:
                    row_str = " ".join([
                        r.get("meeting", ""),
                        r.get("source_tdoc", ""),
                        r.get("cr_number", ""),
                        r.get("subject", ""),
                        r.get("category", ""),
                        r.get("affected_clauses", ""),
                        " ".join(t["raw_tdoc"] for t in r.get("tdocs", []))
                    ]).lower()
                    if search_lower not in row_str:
                        continue

                # Additive clause boost if additive mode is active
                if clause_val and not is_strict_clause:
                    aff = r.get("affected_clauses", "")
                    subj = r.get("subject", "")
                    if clause_val in aff or _re.search(rf"\b{_re.escape(clause_val)}\b", subj):
                        r = dict(r)
                        r["score"] = r.get("score", 0) + 30
                        r.setdefault("evidence_badges", []).append(f"🎯 Clause Match: {clause_val}")

                filtered.append(r)

            if clause_val and not is_strict_clause:
                filtered.sort(key=lambda x: x.get("score", 0), reverse=True)

        self.filtered_records = filtered
        self._populate_table()

    def _populate_table(self):
        """Renders the filtered records into QTableWidget."""
        self.table.setRowCount(0)
        self.table.setSortingEnabled(False)

        is_ranking_active = hasattr(self, "chk_origin_ranking") and self.chk_origin_ranking.isChecked()

        if is_ranking_active:
            self.table.setColumnCount(9)
            self.table.setHorizontalHeaderLabels([
                "Relevance", "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
            ])
        else:
            self.table.setColumnCount(8)
            self.table.setHorizontalHeaderLabels([
                "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
            ])

        total_rows = len(self.filtered_records)
        self.stats_lbl.setText(f"{total_rows} record(s)")

        for row_idx, r in enumerate(self.filtered_records):
            self.table.insertRow(row_idx)
            col_offset = 1 if is_ranking_active else 0

            if is_ranking_active:
                score = r.get("score", 0)
                role = r.get("role", "")
                role_icon = role.split()[0] if role else "⭐"
                rel_item = NumericSortItem(str(score))
                rel_item.setText(f"{role_icon} {score}")
                rel_item.setTextAlignment(Qt.AlignCenter)
                f_rel = rel_item.font()
                f_rel.setBold(True)
                rel_item.setFont(f_rel)
                if score >= 70:
                    rel_item.setForeground(QBrush(QColor("#0284C7")))
                elif score >= 50:
                    rel_item.setForeground(QBrush(QColor("#0D9488")))
                else:
                    rel_item.setForeground(QBrush(QColor("#64748B")))
                badges = r.get("evidence_badges", [])
                tip = f"<b>{role}</b> (Score: {score}/100)<br/>" + "<br/>".join(f"• {b}" for b in badges)
                rel_item.setToolTip(tip)
                self.table.setItem(row_idx, 0, rel_item)

            # Date
            date_item = QTableWidgetItem(r.get("date", ""))
            date_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 0, date_item)

            # Meeting
            meeting_item = QTableWidgetItem(r.get("meeting", ""))
            meeting_item.setTextAlignment(Qt.AlignCenter)
            meeting_item.setForeground(QBrush(QColor("#0369A1")))
            self.table.setItem(row_idx, col_offset + 1, meeting_item)

            # Version
            ver_text = f"v{r.get('version', '')}" if r.get('version') else "—"
            ver_item = QTableWidgetItem(ver_text)
            ver_item.setTextAlignment(Qt.AlignCenter)
            f = ver_item.font()
            f.setBold(True)
            ver_item.setFont(f)
            self.table.setItem(row_idx, col_offset + 2, ver_item)

            # Type
            cr_num = r.get("cr_number", "")
            pcr_cnt = len(r.get("tdocs", []))
            if cr_num and cr_num != "-":
                type_str = "CR"
            elif pcr_cnt > 0:
                type_str = f"p-CR ({pcr_cnt})"
            else:
                type_str = "Editorial"
            type_item = QTableWidgetItem(type_str)
            type_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 3, type_item)

            # CR #
            cr_item = NumericSortItem(cr_num if cr_num != "-" else "—")
            cr_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 4, cr_item)

            # Cat
            cat_val = r.get("category", "")
            cat_item = QTableWidgetItem(cat_val if cat_val != "-" else "—")
            cat_item.setTextAlignment(Qt.AlignCenter)
            if cat_val == "B":
                cat_item.setForeground(QBrush(QColor("#0284C7")))
                f_cat = cat_item.font()
                f_cat.setBold(True)
                cat_item.setFont(f_cat)
            self.table.setItem(row_idx, col_offset + 5, cat_item)

            # Primary TDoc
            source_tdoc = r.get("source_tdoc", "")
            if not source_tdoc and r.get("tdocs"):
                source_tdoc = r["tdocs"][0]["raw_tdoc"]
            tdoc_item = QTableWidgetItem(source_tdoc if source_tdoc else "—")
            tdoc_item.setTextAlignment(Qt.AlignCenter)
            if source_tdoc and source_tdoc != "—":
                tdoc_item.setForeground(QBrush(QColor("#0284C7")))
                f_tdoc = tdoc_item.font()
                f_tdoc.setBold(True)
                tdoc_item.setFont(f_tdoc)
            self.table.setItem(row_idx, col_offset + 6, tdoc_item)

            # Subject / Comment
            subj_item = QTableWidgetItem(r.get("subject", "").replace("\n", " "))
            self.table.setItem(row_idx, col_offset + 7, subj_item)

        # Set sensible initial column widths
        if is_ranking_active:
            self.table.setColumnWidth(0, 85)   # Relevance
            self.table.setColumnWidth(1, 85)   # Date
            self.table.setColumnWidth(2, 115)  # Meeting
            self.table.setColumnWidth(3, 80)   # Version
            self.table.setColumnWidth(4, 75)   # Type
            self.table.setColumnWidth(5, 65)   # CR #
            self.table.setColumnWidth(6, 50)   # Cat
            self.table.setColumnWidth(7, 115)  # Primary TDoc
        else:
            self.table.setColumnWidth(0, 85)   # Date
            self.table.setColumnWidth(1, 115)  # Meeting
            self.table.setColumnWidth(2, 80)   # Version
            self.table.setColumnWidth(3, 85)   # Type
            self.table.setColumnWidth(4, 65)   # CR #
            self.table.setColumnWidth(5, 50)   # Cat
            self.table.setColumnWidth(6, 115)  # Primary TDoc

        if total_rows > 0:
            self.table.selectRow(0)

    def _format_record_to_html(self, rec: Dict[str, Any]) -> str:
        """Formats the selected entry's metadata and rapporteur notes into interactive HTML with clickable TDoc badges."""
        if not rec:
            return (
                "<html><body style=\"font-family: 'Segoe UI', -apple-system, sans-serif; "
                "font-size: 12px; color: #94A3B8; background-color: #F8FAFC; margin: 10px;\">"
                "<i>Select an entry above to view details, rapporteur notes, and clickable TDoc links.</i>"
                "</body></html>"
            )

        subject_text = rec.get("subject", "").strip()
        role = rec.get("role", "")
        score = rec.get("score", 0)
        evidence_badges = rec.get("evidence_badges", [])
        aff_clauses = rec.get("affected_clauses", "")
        primary = rec.get("source_tdoc", "").strip()
        if primary == "-":
            primary = ""
        meeting = rec.get("meeting", "—")
        version = rec.get("version", "—")
        cr_num = rec.get("cr_number", "—")
        cat = rec.get("category", "—")

        all_tdocs = rec.get("tdocs", [])
        raw_tokens: List[str] = []
        if primary:
            raw_tokens.append(primary)
        for t in all_tdocs:
            tok = t.get("raw_tdoc", "")
            if tok and tok not in raw_tokens:
                raw_tokens.append(tok)

        def _make_badge(token: str, trailing_nbsp: bool = True) -> str:
            t_sp = "&nbsp;" if trailing_nbsp else ""
            return (
                f'<a href="tdoc:{token}" title="Click to open {token} (Right-click for options)" '
                f'style="color: #0369A1; background-color: #E0F2FE; text-decoration: none; '
                f'font-weight: bold; font-family: Consolas, monospace;">'
                f'&nbsp;{token}{t_sp}</a>'
            )

        if subject_text:
            escaped = html.escape(subject_text)

            def _repl_tdoc(match):
                token = match.group(1)
                next_char = escaped[match.end():match.end()+1]
                trailing = next_char not in (",", ";", ":", ".", "!", "?", ")", "]", '"')
                return _make_badge(token, trailing_nbsp=trailing)

            html_body = RE_TDOC_TOKEN.sub(_repl_tdoc, escaped)
            lines = html_body.split("\n")
            formatted_lines = []
            for line in lines:
                l_str = line.strip()
                if not l_str:
                    formatted_lines.append("<div style='height: 6px;'></div>")
                elif l_str.startswith(("- ", "* ", "• ")):
                    clean_b = l_str[2:].strip()
                    formatted_lines.append(f"<div style='margin-left: 14px; margin-bottom: 4px;'>• {clean_b}</div>")
                else:
                    formatted_lines.append(f"<div style='margin-bottom: 5px; line-height: 1.5;'>{line}</div>")
            notes_html = "\n".join(formatted_lines)
        else:
            notes_html = "<div style='color: #64748B;'><i>(No rapporteur notes recorded for this entry)</i></div>"

        # Top metadata row
        primary_html = ""
        if primary and primary not in subject_text:
            primary_html = f"&nbsp;&nbsp;|&nbsp;&nbsp;<b>Primary TDoc:</b> {_make_badge(primary, False)}"

        count_html = ""
        if len(raw_tokens) > 0:
            count_html = f"&nbsp;&nbsp;|&nbsp;&nbsp;<b style='color: #0284C7;'>{len(raw_tokens)} TDoc(s)</b>"

        meta_row = (
            f"<div style='padding-bottom: 8px; margin-bottom: 10px; border-bottom: 1px solid #E2E8F0; "
            f"font-size: 11px; color: #475569;'>"
            f"<b>Meeting:</b> {meeting} &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>Version:</b> v{version} &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>CR:</b> {cr_num} (Cat {cat})"
            f"{primary_html}"
            f"{count_html}"
            f"</div>"
        )

        role_banner = ""
        if role or score:
            badges_html = " ".join(
                f'<span style="background-color: #E0F2FE; color: #0369A1; border: 1px solid #BAE6FD; border-radius: 3px; padding: 2px 6px; font-size: 11px; margin-right: 4px;">{html.escape(b)}</span>'
                for b in evidence_badges
            )
            aff_html = f'<br/><span style="color: #475569; font-size: 11px;"><b>Affected Clauses:</b> <code>{html.escape(aff_clauses)}</code></span>' if aff_clauses else ""
            role_banner = f'''
            <div style="background-color: #F0F9FF; border: 1px solid #B9E6FE; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-size: 13px; font-weight: bold; color: #0284C7; margin-bottom: 4px;">
                    {html.escape(role)} <span style="font-size: 11px; font-weight: normal; color: #64748B;">(Relevance Score: <b>{score}</b>/100)</span>
                </div>
                <div>{badges_html}</div>
                {aff_html}
            </div>
            '''

        full_html = f"""
        <html>
        <head>
        <style>
          body {{
            font-family: 'Segoe UI', -apple-system, sans-serif;
            font-size: 12px;
            color: #1E293B;
            background-color: #F8FAFC;
            margin: 6px 8px;
          }}
        </style>
        </head>
        <body>
          {role_banner}
          {meta_row}
          {notes_html}
        </body>
        </html>
        """
        return full_html

    def _on_table_selection_changed(self):
        """Populates the bottom drawer with the selected row's interactive notes and clickable TDocs."""
        selected = self.table.selectedItems()
        if not selected:
            self.notes_browser.setHtml(self._format_record_to_html({}))
            self.drawer_title_lbl.setText("<b>Selected Entry Details:</b>")
            self.btn_open_primary.setText("📄 Open TDoc")
            self.btn_open_primary.setEnabled(False)
            self.btn_copy_entry_tdocs.setEnabled(False)
            return

        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return

        rec = self.filtered_records[row_idx]

        # Collect unique raw tokens
        all_tdocs = rec.get("tdocs", [])
        primary = rec.get("source_tdoc", "")
        raw_tokens: List[str] = []
        if primary and primary != "-":
            raw_tokens.append(primary)
        for t in all_tdocs:
            tok = t.get("raw_tdoc", "")
            if tok and tok not in raw_tokens:
                raw_tokens.append(tok)

        self.notes_browser.setHtml(self._format_record_to_html(rec))

        # Update drawer header title
        m_name = rec.get("meeting", "")
        v_name = f"v{rec.get('version', '')}" if rec.get("version") else ""
        cr_text = f"CR {rec.get('cr_number')}" if rec.get("cr_number") and rec.get("cr_number") != "-" else "p-CRs / Draft"
        header_text = f"<b>Selected Entry:</b> {f'{v_name} · ' if v_name else ''}{f'{m_name} · ' if m_name else ''}{cr_text}"
        if raw_tokens:
            header_text += f" <span style='color: #64748B;'>({len(raw_tokens)} TDoc{'s' if len(raw_tokens) > 1 else ''})</span>"
        self.drawer_title_lbl.setText(header_text)

        self.btn_open_primary.setEnabled(bool(raw_tokens))
        if raw_tokens:
            self.btn_open_primary.setText(f"📄 Open {raw_tokens[0]}")
        else:
            self.btn_open_primary.setText("📄 Open TDoc")

        self.btn_copy_entry_tdocs.setEnabled(bool(raw_tokens))

    def _on_notes_anchor_clicked(self, url: QUrl):
        """Dispatches left-click navigation on interactive links in the notes browser."""
        url_str = url.toString()
        if url_str.startswith("tdoc:"):
            tdoc_id = url_str.split(":", 1)[1]
            row_idx = self.table.currentRow()
            meeting = ""
            if 0 <= row_idx < len(self.filtered_records):
                meeting = self.filtered_records[row_idx].get("meeting", "")
            self._open_tdoc(tdoc_id, meeting)
        elif url.scheme() in ("http", "https"):
            from PyQt5.QtGui import QDesktopServices
            QDesktopServices.openUrl(url)

    def _show_notes_context_menu(self, pos: QPoint):
        """Shows context menu on right-click inside the notes browser, targeting TDoc links or selection."""
        anchor = self.notes_browser.anchorAt(pos)
        row_idx = self.table.currentRow()
        rec = self.filtered_records[row_idx] if 0 <= row_idx < len(self.filtered_records) else {}
        meeting_name = rec.get("meeting", "")

        menu = QMenu(self)

        if anchor and anchor.startswith("tdoc:"):
            tdoc_id = anchor.split(":", 1)[1]
            act_open = menu.addAction(f"📄 Open {tdoc_id} Document")
            act_open.triggered.connect(lambda: self._open_tdoc(tdoc_id, meeting_name))

            act_jump = menu.addAction(f"🚀 Jump to {tdoc_id} in Meetings Tab")
            act_jump.triggered.connect(lambda: SpecTDocOpener.jump_to_meetings_tab(tdoc_id))

            act_portal = menu.addAction(f"🌐 Open {tdoc_id} on 3GPP Portal (3GU)")
            act_portal.triggered.connect(lambda: SpecTDocOpener.open_tdoc_in_browser(tdoc_id))

            menu.addSeparator()
            act_copy_cit = menu.addAction(f"📋 Copy Citation ({tdoc_id})")
            act_copy_cit.triggered.connect(
                lambda: SpecTDocOpener.copy_tdoc_citation(
                    tdoc_id, self.spec_number, rec.get("version", ""), meeting_name, rec.get("subject", ""), parent=self
                )
            )

            act_copy_id = menu.addAction(f"📋 Copy Identifier ({tdoc_id})")
            act_copy_id.triggered.connect(
                lambda: _copy_to_clipboard(tdoc_id, label=tdoc_id, parent=self)
            )
        else:
            if self.notes_browser.textCursor().hasSelection():
                act_copy_sel = menu.addAction("📋 Copy Selection")
                act_copy_sel.triggered.connect(self.notes_browser.copy)

            act_select_all = menu.addAction("Select All")
            act_select_all.triggered.connect(self.notes_browser.selectAll)

            if rec:
                menu.addSeparator()
                act_copy_all = menu.addAction("📋 Copy All TDocs for this Entry")
                act_copy_all.triggered.connect(lambda: self._copy_entry_tdocs(rec))

        menu.exec_(self.notes_browser.viewport().mapToGlobal(pos))

    def _copy_selected_entry_tdocs(self):
        """Copies all TDocs referenced in the currently selected row."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return
        self._copy_entry_tdocs(self.filtered_records[row_idx])

    def _open_selected_tdoc(self):
        """Opens the primary or first TDoc of the currently selected row."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return
        rec = self.filtered_records[row_idx]
        primary = rec.get("source_tdoc", "")
        if (not primary or primary == "-") and rec.get("tdocs"):
            primary = rec["tdocs"][0]["raw_tdoc"]
        if primary and primary != "-":
            self._open_tdoc(primary, rec.get("meeting", ""))

    def _on_row_double_clicked(self, item: QTableWidgetItem):
        self._open_selected_tdoc()

    def _open_tdoc(self, tdoc_id: str, meeting_name: str = ""):
        """Dispatches direct opening of the target TDoc Word/ZIP file."""
        SpecTDocOpener.open_tdoc(tdoc_id, meeting_name, parent=self)

    def _show_context_menu(self, pos: QPoint):
        """Context menu on table rows."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return

        rec = self.filtered_records[row_idx]
        primary_tdoc = rec.get("source_tdoc", "")
        if not primary_tdoc and rec.get("tdocs"):
            primary_tdoc = rec["tdocs"][0]["raw_tdoc"]

        menu = QMenu(self)
        if primary_tdoc:
            act_open = menu.addAction(f"📄 Open {primary_tdoc}")
            act_open.triggered.connect(lambda: self._open_tdoc(primary_tdoc, rec.get("meeting", "")))

            act_jump = menu.addAction(f"🚀 Jump to {primary_tdoc} in Meetings Tab")
            act_jump.triggered.connect(lambda: SpecTDocOpener.jump_to_meetings_tab(primary_tdoc))

            menu.addSeparator()
            act_copy = menu.addAction(f"📋 Copy Citation ({primary_tdoc})")
            act_copy.triggered.connect(
                lambda: SpecTDocOpener.copy_tdoc_citation(
                    primary_tdoc, self.spec_number, rec.get("version", ""), rec.get("meeting", ""), rec.get("subject", ""), parent=self
                )
            )

        act_copy_all = menu.addAction("📋 Copy All TDocs for this Entry")
        act_copy_all.triggered.connect(lambda: self._copy_entry_tdocs(rec))

        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def _copy_entry_tdocs(self, rec: Dict[str, Any]):
        """Copies all TDoc tokens from a single entry."""
        from modules.meetings.ui.tdocs_menus import _copy_to_clipboard

        tdocs = [t["raw_tdoc"] for t in rec.get("tdocs", [])]
        primary = rec.get("source_tdoc", "")
        if primary and primary not in tdocs:
            tdocs.insert(0, primary)

        plain = ", ".join(tdocs)
        _copy_to_clipboard(plain, label=f"{len(tdocs)} TDoc(s)", parent=self)

    def _copy_tdocs_list(self):
        """Copies all TDocs for the active version filter as a comma-separated list."""
        from modules.meetings.ui.tdocs_menus import _copy_to_clipboard

        all_tokens = []
        for r in self.filtered_records:
            primary = r.get("source_tdoc", "")
            if primary and primary != "-" and primary not in all_tokens:
                all_tokens.append(primary)
            for t in r.get("tdocs", []):
                tok = t.get("raw_tdoc", "")
                if tok and tok not in all_tokens:
                    all_tokens.append(tok)

        if not all_tokens:
            QToolTip.showText(QCursor.pos(), "No TDocs to copy.", self)
            return

        plain = ", ".join(all_tokens)
        _copy_to_clipboard(plain, label=f"{len(all_tokens)} TDoc(s)", parent=self)

    def _open_spec_document(self):
        """Locates and opens the local specification document."""
        from modules.word_tools.core.document_opener import open_document

        target = SpecHistoryExtractor.resolve_spec_document(
            self.spec_number, self.version_combo.currentData() or ""
        )
        if target and target.exists():
            open_document(target)
        else:
            QToolTip.showText(QCursor.pos(), f"No local document found for {self.spec_number}", self)

    def _on_rescan_clicked(self):
        """Executes background re-extraction from local document."""
        self.rescan_btn.setEnabled(False)
        self.rescan_btn.setText("⏳ Scanning...")

        self.extract_thread = SpecHistoryExtractThread(
            spec_number=self.spec_number,
            db=self.history_db,
            version=self.version_combo.currentData() or ""
        )

        def _on_success(s_num, count):
            self.rescan_btn.setEnabled(True)
            self.rescan_btn.setText("🔄 Re-scan from Local Doc")
            self._load_data(auto_extract_if_empty=False)
            QToolTip.showText(QCursor.pos(), f"✅ Updated: {count} records indexed.", self)

        def _on_error(err):
            self.rescan_btn.setEnabled(True)
            self.rescan_btn.setText("🔄 Re-scan from Local Doc")
            QToolTip.showText(QCursor.pos(), f"❌ Scan failed: {err}", self)

        self.extract_thread.finished_success.connect(_on_success)
        self.extract_thread.error.connect(_on_error)
        self.extract_thread.start()

    def _generate_tdoc_report(self, scope: str = "version"):
        """
        Gathers TDocs referenced in the specified scope, resolves their technical metadata
        from meeting spreadsheets in background, and presents the interactive dossier dialog.
        """
        tdoc_set = set()
        default_meeting = ""

        if scope == "selection":
            selected_rows = self.table.selectionModel().selectedRows()
            if not selected_rows:
                QMessageBox.information(self, "No Rows Selected", "Please select one or more rows in the Change History table first.")
                return
            for idx in selected_rows:
                row_idx = idx.row()
                item = self.table.item(row_idx, 0)
                if item:
                    r = item.data(Qt.UserRole)
                    if r:
                        if r.get("source_tdoc"):
                            tdoc_set.add(r["source_tdoc"])
                        for t in r.get("tdocs", []):
                            tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                        if not default_meeting and r.get("meeting"):
                            default_meeting = r["meeting"]
        elif scope == "all":
            for r in self.all_records:
                if r.get("source_tdoc"):
                    tdoc_set.add(r["source_tdoc"])
                for t in r.get("tdocs", []):
                    tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                if not default_meeting and r.get("meeting"):
                    default_meeting = r["meeting"]
        else:  # "version"
            target_ver = self.version_combo.currentData() or ""
            target_records = [r for r in self.all_records if not target_ver or r.get("version") == target_ver]
            for r in target_records:
                if r.get("source_tdoc"):
                    tdoc_set.add(r["source_tdoc"])
                for t in r.get("tdocs", []):
                    tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                if not default_meeting and r.get("meeting"):
                    default_meeting = r["meeting"]

        clean_tdocs = sorted([t.strip() for t in tdoc_set if t and t.strip() != "-"])
        if not clean_tdocs:
            QMessageBox.information(
                self,
                "No TDocs Found",
                f"No TDocs or p-CRs are referenced in the selected scope for TS {self.spec_number}."
            )
            return

        active_ver = self.version_combo.currentData() or (self.filtered_records[0].get("version", "") if self.filtered_records else "")
        ver_label = active_ver if scope != "all" else "all"

        logger.info(f"📑 [TDoc Report] Starting dossier generation for TS {self.spec_number} v{ver_label} ({len(clean_tdocs)} TDocs)...")

        # Progress Dialog
        progress_dlg = QProgressDialog(
            f"Resolving metadata for {len(clean_tdocs)} TDocs from meeting agendas...",
            "Cancel",
            0,
            100,
            self
        )
        progress_dlg.setWindowTitle(f"Generating Technical Dossier: TS {self.spec_number}")
        progress_dlg.setMinimumDuration(150)
        progress_dlg.setValue(5)

        thread = SpecTDocReportThread(
            spec_number=self.spec_number,
            version=ver_label,
            tdoc_list=clean_tdocs,
            spec_title=self.title_lbl.text(),
            default_meeting=default_meeting,
            parent=self
        )

        def _on_prog(pct: int, msg: str):
            if not progress_dlg.wasCanceled():
                progress_dlg.setValue(pct)
                progress_dlg.setLabelText(msg)

        def _on_finish(results: list, md_report: str):
            progress_dlg.close()
            logger.info(f"✅ [TDoc Report] Generated dossier with {len(results)} TDocs for TS {self.spec_number} v{ver_label}.")
            report_dlg = SpecTDocReportDialog(
                spec_number=self.spec_number,
                version=ver_label,
                tdoc_entries=results,
                markdown_report=md_report,
                spec_title=self.title_lbl.text(),
                parent=self._parent_ref or self
            )
            self._report_dialog = report_dlg
            report_dlg.show()
            report_dlg.raise_()
            report_dlg.activateWindow()

        def _on_err(err_msg: str):
            progress_dlg.close()
            logger.error(f"❌ [TDoc Report] Error generating dossier: {err_msg}")
            QMessageBox.critical(self, "Report Generation Error", f"Failed to generate TDoc report:\n{err_msg}")

        thread.progress.connect(_on_prog)
        thread.finished.connect(_on_finish)
        thread.error.connect(_on_err)
        progress_dlg.canceled.connect(thread.requestInterruption)

        self._active_report_thread = thread
        thread.start()
