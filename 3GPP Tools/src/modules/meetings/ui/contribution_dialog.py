# --- File: src/modules/meetings/ui/contribution_dialog.py ---
import os
import re
import logging
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QAbstractTableModel, QModelIndex, QPoint, QStringListModel, QTimer
from PyQt5.QtGui import QColor, QBrush, QFont
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QLabel, QLineEdit, QPushButton, QCheckBox, QRadioButton,
    QButtonGroup, QDateEdit, QListWidget, QListWidgetItem,
    QTableView, QHeaderView, QTextEdit, QProgressBar,
    QMessageBox, QFileDialog, QFrame, QCompleter, QMenu, QApplication
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
    SEARCH_INPUT_STYLE
)
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.stats.contribution_stats import ContributionStatsExporterThread
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import ReadOnlyViewerDialog, StatisticsSettingsDialog
from modules.work_items.core.wi_database import WorkItemsDatabase

logger = logging.getLogger(__name__)



# Imported from dedicated submodules
from modules.meetings.ui.contribution_widgets import (
    KPICard,
    FilterTokenChip,
    WITokenInputWidget,
    AITokenInputWidget,
    CompanySelectorWidget
)
from modules.meetings.core.contribution_worker import ContributionSearchWorker



# Re-exported for backward compatibility
from modules.meetings.ui.contribution_models import ContributionResultsTableModel
from modules.meetings.ui.contribution_export_controller import (
    export_to_excel,
    on_export_finished,
    generate_statistics,
    on_stats_finished,
)


class ContributionReportDialog(QDialog):
    """Independent modeless cross-meeting contribution report window."""
    open_meeting_requested = pyqtSignal(str, str)  # (wg_name, meeting_number)

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        # Pass None as parent so the OS does not pin this window on top of the main window.
        # Adding MinMaxButtonsHint provides standard minimize/maximize title bar buttons.
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)

        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.worker = None
        self.exporter_thread = None
        self.stats_thread = None
        self.active_tdoc_threads = {}

        self.setWindowTitle("3GPP Company Contribution Search & Report")
        self.resize(1340, 820)
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal)

        # Left: Filter Sidebar
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(6)

        title_lbl = QLabel("<b>Search & Audit Filters</b>")
        title_lbl.setStyleSheet("font-size: 13px; color: #1E293B;")
        left_layout.addWidget(title_lbl)

        left_layout.addWidget(QLabel("Working Groups:"))
        self.wg_filter = CheckableComboBox("Working Groups")
        wgs = self.meetings_db.get_working_groups()
        self.wg_filter.updateItems(wgs)
        left_layout.addWidget(self.wg_filter)

        left_layout.addWidget(QLabel("Date Range:"))
        date_box = QHBoxLayout()
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())

        date_box.addWidget(self.date_from)
        date_box.addWidget(QLabel("to"))
        date_box.addWidget(self.date_to)
        left_layout.addLayout(date_box)

        left_layout.addWidget(QLabel("Work Items (Acronym / Code):"))
        self.wi_token_widget = WITokenInputWidget(self.wi_db, self)
        left_layout.addWidget(self.wi_token_widget)

        left_layout.addWidget(QLabel("Agenda Items:"))
        self.ai_token_widget = AITokenInputWidget(self)
        left_layout.addWidget(self.ai_token_widget)

        self.chk_wi_ai_or = QCheckBox("Match WI OR Agenda Item")
        self.chk_wi_ai_or.setChecked(False)
        self.chk_wi_ai_or.setToolTip(
            "When enabled, a contribution may match either a selected WI or a selected Agenda Item. "
            "When disabled, it must match both when both filters are populated."
        )
        left_layout.addWidget(self.chk_wi_ai_or)

        left_layout.addWidget(QLabel("Companies (from Sanitizer):"))
        self.company_widget = CompanySelectorWidget(self)
        left_layout.addWidget(self.company_widget)

        self.chk_bypass_cache = QCheckBox("Bypass local cache (re-fetch from 3GPP)")
        self.chk_bypass_cache.setToolTip("Force downloading and parsing the latest TDoc list sheets from the 3GPP portal.")
        left_layout.addWidget(self.chk_bypass_cache)

        btn_box = QHBoxLayout()
        self.btn_search = QPushButton("🔍 Search Contributions")
        self.btn_search.setObjectName("primaryBtn")
        self.btn_search.clicked.connect(self._start_search)

        self.btn_cancel = QPushButton("🛑 Stop")
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_search)

        btn_box.addWidget(self.btn_search)
        btn_box.addWidget(self.btn_cancel)
        left_layout.addLayout(btn_box)

        splitter.addWidget(left_widget)

        # Right: Results, KPIs & Logs
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)

        # In-App KPI Summary Bar
        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(8)
        self.kpi_total = KPICard("Total Contributions", "0 TDocs", "Filtered dataset")
        self.kpi_agree = KPICard("Agreed / Approved", "0 TDocs", "0.0% of total")
        self.kpi_joint = KPICard("Collaboration", "0% Joint", "Co-authored")
        self.kpi_partner = KPICard("Top Co-signer", "None", "Most frequent ally")
        self.kpi_wi = KPICard("Top Work Item", "None", "Highest volume WI")

        kpi_layout.addWidget(self.kpi_total)
        kpi_layout.addWidget(self.kpi_agree)
        kpi_layout.addWidget(self.kpi_joint)
        kpi_layout.addWidget(self.kpi_partner)
        kpi_layout.addWidget(self.kpi_wi)
        right_layout.addLayout(kpi_layout)

        # Persistent summary of the criteria that produced the current result set.
        self.filter_summary_lbl = QLabel("No search run yet.")
        self.filter_summary_lbl.setWordWrap(True)
        self.filter_summary_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.filter_summary_lbl.setStyleSheet(
            "QLabel { background-color: #F8FAFC; color: #475569; "
            "border: 1px solid #E2E8F0; border-radius: 4px; "
            "padding: 5px 8px; font-size: 11px; }"
        )
        right_layout.addWidget(self.filter_summary_lbl)

        # Progress Bar & Log Viewer
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setTextVisible(False)
        right_layout.addWidget(self.progress_bar)

        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMaximumHeight(80)
        self.log_viewer.setStyleSheet("background-color: #0F172A; color: #38BDF8; font-family: monospace; font-size: 11px;")
        right_layout.addWidget(self.log_viewer)

        # Table View
        self.table_model = ContributionResultsTableModel()
        self.table_view = QTableView()
        self.table_view.setModel(self.table_model)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setStyleSheet(
            "QTableView { border: 1px solid #CBD5E1; gridline-color: #F1F5F9; background-color: #FFFFFF; } "
            "QTableView::item:selected { background-color: #EBF3FC; color: #1E293B; }"
        )
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.horizontalHeader().setStretchLastSection(True)

        self.table_view.doubleClicked.connect(self._on_table_double_clicked)
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_table_context_menu)

        self._apply_column_widths()
        right_layout.addWidget(self.table_view)

        # Footer Actions
        footer_layout = QHBoxLayout()
        self.lbl_count = QLabel("0 contributions found.")
        self.lbl_count.setStyleSheet("color: #64748B; font-weight: bold;")

        self.btn_stats = QPushButton("📊 Statistics Dashboard...")
        self.btn_stats.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_stats.setEnabled(False)
        self.btn_stats.setToolTip("Compile an interactive HTML analytics report with graphs and timelines.")
        self.btn_stats.clicked.connect(self._generate_statistics)

        self.btn_export = QPushButton("📥 Export to Excel...")
        self.btn_export.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_to_excel)

        footer_layout.addWidget(self.lbl_count)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_stats)
        footer_layout.addWidget(self.btn_export)
        right_layout.addLayout(footer_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([380, 960])
        main_layout.addWidget(splitter)

    def _apply_column_widths(self):
        header = self.table_view.horizontalHeader()
        widths = {
            0: 65,   # WG
            1: 95,   # Meeting
            2: 95,   # TDoc
            3: 280,  # Title
            4: 160,  # Source
            5: 130,  # Matched Company
            6: 60,   # Type
            7: 75,   # For
            8: 85,   # Agenda Item
            9: 95,   # TDoc Status
            10: 140, # Related WIs
            11: 250  # Abstract
        }
        for col_idx, width in widths.items():
            if col_idx < len(self.table_model.COLUMNS):
                header.resizeSection(col_idx, width)

    def set_active_wgs(self, wgs: list):
        for i in range(1, self.wg_filter.model().rowCount()):
            item = self.wg_filter.model().item(i)
            if item:
                item.setCheckState(Qt.Checked if item.text() in wgs else Qt.Unchecked)
        self.wg_filter.update()

    def _cleanup_threads(self):
        """Cleanly halts all worker threads to prevent QThread destroyed while running."""
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.quit()
            self.worker.wait(300)
        if self.exporter_thread and self.exporter_thread.isRunning():
            self.exporter_thread.quit()
            self.exporter_thread.wait(200)
        if self.stats_thread and self.stats_thread.isRunning():
            self.stats_thread.quit()
            self.stats_thread.wait(200)
        for th in list(self.active_tdoc_threads.values()):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(200)
        self.active_tdoc_threads.clear()

    def closeEvent(self, event):
        self._cleanup_threads()
        super().closeEvent(event)

    def reject(self):
        self._cleanup_threads()
        super().reject()

    def _start_search(self):
        selected_wgs = self.wg_filter.getCheckedItems()
        target_companies = self.company_widget.get_selected_companies()
        primary_only = self.company_widget.is_primary_only()
        target_wis = self.wi_token_widget.get_tokens()
        target_ais = self.ai_token_widget.get_tokens()
        wi_ai_or = self.chk_wi_ai_or.isChecked()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")
        bypass = self.chk_bypass_cache.isChecked()

        summary_parts = []
        if selected_wgs:
            summary_parts.append("WG: " + ", ".join(selected_wgs))
        summary_parts.append(f"{date_from} → {date_to}")
        if target_companies:
            company_rule = "primary" if primary_only else "any co-author"
            summary_parts.append("Companies: " + ", ".join(sorted(target_companies)) + f" ({company_rule})")
        wi_text = ", ".join(sorted(target_wis)) if target_wis else ""
        ai_text = ", ".join(sorted(target_ais)) if target_ais else ""
        if wi_text and ai_text:
            joiner = " OR " if wi_ai_or else " AND "
            summary_parts.append(f"WI: {wi_text}{joiner}AI: {ai_text}")
        elif wi_text:
            summary_parts.append(f"WI: {wi_text}")
        elif ai_text:
            summary_parts.append(f"AI: {ai_text}")
        self.filter_summary_lbl.setText("  •  ".join(summary_parts) if summary_parts else "All contributions")

        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export.setEnabled(False)
        self.btn_stats.setEnabled(False)
        self.log_viewer.clear()
        self.progress_bar.setValue(0)

        self.worker = ContributionSearchWorker(
            meetings_db=self.meetings_db,
            wgs=selected_wgs,
            date_from=date_from,
            date_to=date_to,
            target_companies=target_companies,
            primary_only=primary_only,
            target_wis=target_wis,
            target_ais=target_ais,
            wi_ai_or=wi_ai_or,
            bypass_cache=bypass,
            cache_dir=self.settings.cache_dir,
            parent=self
        )

        self.worker.stage_progress.connect(self._on_stage_progress)
        self.worker.log_msg.connect(self._append_log)
        self.worker.results_ready.connect(self._on_results_ready)
        self.worker.finished.connect(self._on_search_finished)
        self.worker.start()

    def _cancel_search(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.btn_cancel.setEnabled(False)

    def _on_stage_progress(self, desc: str, current: int, total: int):
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(current)
        self.setWindowTitle(f"[{current}/{total}] {desc} - 3GPP Contributions")

    def _append_log(self, msg: str):
        self.log_viewer.append(msg)
        self.log_viewer.verticalScrollBar().setValue(self.log_viewer.verticalScrollBar().maximum())

    def _on_results_ready(self, results: list):
        self.table_model.update_data(results)
        total = len(results)
        self.lbl_count.setText(f"Found {total} matching contribution(s).")
        self._apply_column_widths()

        if total == 0:
            self.kpi_total.set_value("0 TDocs", "No matches")
            self.kpi_agree.set_value("0 TDocs", "0.0% of total")
            self.kpi_joint.set_value("0% Joint", "0 joint")
            self.kpi_partner.set_value("None", "No allies")
            self.kpi_wi.set_value("None", "No WIs")
            self.btn_export.setEnabled(False)
            self.btn_stats.setEnabled(False)
            return

        # 1. Total KPI & WG Coverage
        wgs_count = len(set(r.get("WG", "") for r in results if r.get("WG")))
        self.kpi_total.set_value(f"{total} TDocs", f"Across {wgs_count} WG(s)")

        # 2. Agreed / Approved KPI (Standard 3GPP Metric)
        agreed_count = sum(
            1 for r in results if any(w in str(r.get("TDoc Status", "")).lower() for w in ["agreed", "approved"])
        )
        agree_pct = round((agreed_count / total) * 100, 1) if total else 0.0
        self.kpi_agree.set_value(f"{agreed_count} TDocs", f"{agree_pct}% of {total} total")

        # 3. Collaboration KPI & 4. Top Co-signer
        joint_count = 0
        partner_counts = {}
        target_companies = self.company_widget.get_selected_companies()

        for r in results:
            src = str(r.get("Source", ""))
            contribs = CompanySanitizer.get_matching_contributors(src)
            if len(contribs) > 1:
                joint_count += 1

            if target_companies:
                targets_in_row = [c for c in contribs if c in target_companies]
                if targets_in_row:
                    for c in contribs:
                        if c not in target_companies:
                            partner_counts[c] = partner_counts.get(c, 0) + 1
            else:
                for c in contribs:
                    partner_counts[c] = partner_counts.get(c, 0) + 1

        joint_pct = round((joint_count / total) * 100, 1)
        solo_count = total - joint_count
        self.kpi_joint.set_value(f"{joint_pct}% Joint", f"{solo_count} solo, {joint_count} joint")

        if partner_counts:
            top_partner, top_count = max(partner_counts.items(), key=lambda x: x[1])
            self.kpi_partner.set_value(top_partner, f"{top_count} joint TDocs")
        else:
            self.kpi_partner.set_value("N/A", "No co-signers")

        # 5. Top Work Item KPI (splits multi-entry Related WIs, filters out placeholders)[cite: 32]
        wi_counts = {}
        for r in results:
            wi_str = str(r.get("Related WIs", "")).strip()
            if wi_str and wi_str.lower() not in ["", "none", "unknown", "-", "dummy"]:
                for item in re.split(r'[,;]+', wi_str):
                    item_clean = item.strip()
                    if item_clean and item_clean.lower() != 'dummy':
                        wi_counts[item_clean] = wi_counts.get(item_clean, 0) + 1

        if wi_counts:
            top_wi, wi_vol = max(wi_counts.items(), key=lambda x: x[1])
            disp_wi = top_wi if len(top_wi) <= 18 else top_wi[:16] + ".."
            self.kpi_wi.set_value(disp_wi, f"{wi_vol} TDocs")
        else:
            self.kpi_wi.set_value("Unspecified", "No WI tagged")

        self.btn_export.setEnabled(True)
        self.btn_stats.setEnabled(True)

    def _on_search_finished(self, success: bool, msg: str):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("3GPP Company Contribution Search & Report")
        if not success:
            QMessageBox.warning(self, "Search Notice", msg)

    # --- ROW ACTIONS: OPEN MEETING, OPEN DOCS FOLDER, OPEN TDOC ---

    def _request_open_meeting(self, row_data: dict):
        """Requests opening the full meeting window in the main application."""
        wg = str(row_data.get("WG", "")).strip()
        mtg_raw = str(row_data.get("Meeting", "")).strip()
        mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
        if wg and mtg_num:
            self.open_meeting_requested.emit(wg, mtg_num)
        else:
            QMessageBox.warning(self, "Missing Info", "Cannot identify the Working Group or meeting number for this row.")

    def _open_docs_folder(self, row_data: dict):
        """Opens the Docs folder of the meeting for this TDoc in the default web browser."""
        docs_url = str(row_data.get("docs_folder_url", "")).strip()
        if not docs_url:
            wg = row_data.get("WG", "")
            mtg_num = row_data.get("meeting_number", "")
            mtgs = self.meetings_db.search_meetings(wg_name=[wg] if wg else None, search_term=mtg_num)
            if mtgs:
                docs_url = mtgs[0].get("docs_folder_url", "")

        if docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            webbrowser.open(full_url)
        else:
            QMessageBox.warning(self, "Missing URL", "No documents folder URL available for this meeting.")

    def _get_tdoc_download_url(self, row_data: dict) -> str:
        """Constructs the fully qualified 3GPP FTP/HTTP download URL for the TDoc."""
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return ""

        docs_url = str(row_data.get("docs_folder_url", "")).strip()
        if not docs_url:
            wg = row_data.get("WG", "")
            mtg_num = row_data.get("meeting_number", "")
            mtgs = self.meetings_db.search_meetings(wg_name=[wg] if wg else None, search_term=mtg_num)
            if mtgs:
                docs_url = mtgs[0].get("docs_folder_url", "")

        if docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            return f"{full_url.rstrip('/')}/{tdoc_id}.zip"
        return ""

    def _open_tdoc(self, row_data: dict):
        """Downloads (if missing) and opens the actual document using TDocActionThread."""
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return

        match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
        target_filename = tdoc_id.upper()

        if target_filename in self.active_tdoc_threads:
            th = self.active_tdoc_threads[target_filename]
            if th and th.isRunning():
                return

        wg_name = row_data.get("WG") or row_data.get("wg_name", "")
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        raw_url = row_data.get("url_key", "")
        docs_url = row_data.get("docs_folder_url", "")

        main_ftp = raw_url if raw_url.startswith("http") else (f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}" if raw_url else "")

        is_active_sync = self.meetings_db.is_active_sync_meeting(
            wg_name,
            row_data.get("start_date", ""),
            row_data.get("end_date", ""),
            int(row_data.get("is_electronic", 0) or 0)
        )

        candidate_urls = URLRouter.build_priority_url_list(
            wg_name,
            folder_name,
            main_ftp,
            is_active_sync,
            target_filename=target_filename
        )

        if docs_url:
            clean_docs = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            clean_docs = clean_docs.rstrip('/')
            if clean_docs not in candidate_urls:
                candidate_urls.append(clean_docs)

        if not candidate_urls:
            QMessageBox.warning(self, "Missing URL", f"No candidate download routes found for {tdoc_id}.")
            return

        meeting_dir = Path(self.settings.cache_dir) / folder_name
        meeting_dir.mkdir(parents=True, exist_ok=True)

        thread = TDocActionThread(base_tdoc, target_filename, candidate_urls, meeting_dir, open_file=True)
        self.active_tdoc_threads[target_filename] = thread
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if thread.target_filename in self.active_tdoc_threads:
            del self.active_tdoc_threads[thread.target_filename]
        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _on_table_double_clicked(self, index: QModelIndex):
        """Handles double-click navigation based on the clicked column."""
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        col_name = self.table_model.COLUMNS[index.column()]

        # Double-clicking WG or Meeting opens the meeting window
        if col_name in ["WG", "Meeting"]:
            self._request_open_meeting(row_data)
            return

        # Double-clicking Abstract opens the popup viewer
        if col_name == "Abstract":
            abstract_text = str(row_data.get("Abstract", "")).strip()
            tdoc_id = str(row_data.get("TDoc", "")).strip()
            if abstract_text:
                ReadOnlyViewerDialog(self, f"📄 Abstract: {tdoc_id}", abstract_text).exec_()
                return

        # Double-clicking any other cell opens the actual TDoc document
        self._open_tdoc(row_data)

    def _on_table_context_menu(self, pos: QPoint):
        """Row-level right-click menu with explicit actions."""
        index = self.table_view.indexAt(pos)
        if not index.isValid():
            return

        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        menu = QMenu(self)
        mtg = row_data.get("Meeting", "")
        tdoc = row_data.get("TDoc", "")

        # 1. Open Meeting Table
        act_open_mtg = menu.addAction(f"🗓️ Open Meeting Table ({mtg})")
        act_open_mtg.triggered.connect(lambda: self._request_open_meeting(row_data))

        # 2. Open Docs Folder
        act_open_docs = menu.addAction("📂 Open Docs Folder")
        act_open_docs.triggered.connect(lambda: self._open_docs_folder(row_data))

        menu.addSeparator()

        # 3. Open / View TDoc
        if tdoc:
            act_open_tdoc = menu.addAction(f"📄 Open TDoc Document ({tdoc})")
            act_open_tdoc.triggered.connect(lambda: self._open_tdoc(row_data))

            act_portal = menu.addAction(f"🌐 View {tdoc} on 3GU Portal")
            act_portal.triggered.connect(lambda: webbrowser.open(f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={tdoc}"))

            menu.addSeparator()

            # Copy options
            act_copy_tdoc = menu.addAction(f"📋 Copy TDoc Number ({tdoc})")
            act_copy_tdoc.triggered.connect(lambda: QApplication.clipboard().setText(tdoc))

            tdoc_url = self._get_tdoc_download_url(row_data)
            if tdoc_url:
                act_copy_url = menu.addAction("🔗 Copy URL")
                act_copy_url.triggered.connect(lambda: QApplication.clipboard().setText(tdoc_url))

        title = str(row_data.get("Title", "")).strip()
        if title:
            act_copy_title = menu.addAction("📋 Copy Title")
            act_copy_title.triggered.connect(lambda: QApplication.clipboard().setText(title))

        abstract = str(row_data.get("Abstract", "")).strip()
        if abstract:
            menu.addSeparator()
            act_view_abs = menu.addAction("📝 View Abstract...")
            act_view_abs.triggered.connect(lambda: ReadOnlyViewerDialog(self, f"📄 Abstract: {tdoc}", abstract).exec_())

        menu.exec_(self.table_view.viewport().mapToGlobal(pos))

    # --- EXPORTS ---

    def _export_to_excel(self):
        export_to_excel(self)

    def _on_export_finished(self, success: bool, msg: str):
        on_export_finished(self, success, msg)

    def _generate_statistics(self):
        generate_statistics(self)

    def _on_stats_finished(self, success: bool, msg: str):
        on_stats_finished(self, success, msg)
