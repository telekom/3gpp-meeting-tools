import os
import re
import logging
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QColor, QBrush, QFont
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QLabel, QLineEdit, QPushButton, QCheckBox, QRadioButton,
    QButtonGroup, QDateEdit, QListWidget, QListWidgetItem,
    QTableView, QHeaderView, QTextEdit, QProgressBar,
    QMessageBox, QFileDialog, QFrame, QCompleter
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER
)
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase


class WITokenChip(QFrame):
    """Visual removable chip/tag for selected Work Items."""
    removed = pyqtSignal(str)

    def __init__(self, token_text: str, parent=None):
        super().__init__(parent)
        self.token_text = token_text
        self.setStyleSheet("""
            QFrame {
                background-color: #EBF3FC;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 1px 6px;
            }
            QLabel {
                color: #1E5C99;
                font-size: 11px;
                font-weight: bold;
                border: none;
            }
            QPushButton {
                background: transparent;
                border: none;
                color: #64748B;
                font-weight: bold;
                font-size: 11px;
                padding: 0px 2px;
            }
            QPushButton:hover {
                color: #DC2626;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(token_text)
        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.clicked.connect(lambda: self.removed.emit(self.token_text))

        layout.addWidget(lbl)
        layout.addWidget(btn_close)


class WITokenInputWidget(QWidget):
    """Token/tag box with autocompletion from the Work Items database."""
    tokens_changed = pyqtSignal()

    def __init__(self, wi_db: WorkItemsDatabase, parent=None):
        super().__init__(parent)
        self.wi_db = wi_db
        self.selected_tokens = set()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(4)

        self.chips_widget = QWidget()
        self.chips_layout = QHBoxLayout(self.chips_widget)
        self.chips_layout.setContentsMargins(0, 0, 0, 0)
        self.chips_layout.setSpacing(4)
        self.chips_layout.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(self.chips_widget)

        input_layout = QHBoxLayout()
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Search WI acronym or code (e.g., 5G_eURLLC)...")
        self.input_edit.returnPressed.connect(self._add_current_input)

        self._setup_completer()

        btn_add = QPushButton("Add")
        btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_add.setFixedWidth(50)
        btn_add.clicked.connect(self._add_current_input)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(btn_add)
        main_layout.addLayout(input_layout)

    def _setup_completer(self):
        try:
            items = self.wi_db.search_work_items(limit=1500)
            acronyms = sorted(list({item.get('acronym') for item in items if item.get('acronym')}))
            completer = QCompleter(acronyms, self)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setFilterMode(Qt.MatchContains)
            self.input_edit.setCompleter(completer)
        except Exception as e:
            logging.warning(f"Could not initialize WI completer: {e}")

    def _add_current_input(self):
        text = self.input_edit.text().strip()
        if text and text not in self.selected_tokens:
            self.selected_tokens.add(text)
            chip = WITokenChip(text, self)
            chip.removed.connect(self._remove_token)
            self.chips_layout.addWidget(chip)
            self.input_edit.clear()
            self.tokens_changed.emit()

    def _remove_token(self, token: str):
        if token in self.selected_tokens:
            self.selected_tokens.remove(token)
            for i in range(self.chips_layout.count()):
                widget = self.chips_layout.itemAt(i).widget()
                if isinstance(widget, WITokenChip) and widget.token_text == token:
                    widget.deleteLater()
                    break
            self.tokens_changed.emit()

    def get_tokens(self) -> list:
        return list(self.selected_tokens)


class CompanySelectorWidget(QWidget):
    """Searchable multi-select checkbox list populated from CompanySanitizer."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter companies list...")
        self.search_input.textChanged.connect(self._filter_list)
        layout.addWidget(self.search_input)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("QListWidget::item { padding: 3px; }")
        layout.addWidget(self.list_widget)

        actions_layout = QHBoxLayout()
        btn_select_all = QPushButton("Select All")
        btn_select_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_select_all.clicked.connect(lambda: self._set_all_checked(True))

        btn_clear_all = QPushButton("Clear All")
        btn_clear_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_clear_all.clicked.connect(lambda: self._set_all_checked(False))

        actions_layout.addWidget(btn_select_all)
        actions_layout.addWidget(btn_clear_all)
        layout.addLayout(actions_layout)

        match_frame = QFrame()
        match_frame.setStyleSheet("QFrame { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px; padding: 4px; }")
        match_layout = QVBoxLayout(match_frame)
        match_layout.setContentsMargins(4, 4, 4, 4)
        match_layout.setSpacing(2)

        match_layout.addWidget(QLabel("<b>Author Match Rule:</b>"))
        self.rb_any = QRadioButton("Any co-author (joint contribution)")
        self.rb_primary = QRadioButton("Primary source only (first listed)")
        self.rb_any.setChecked(True)

        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.rb_any)
        self.btn_group.addButton(self.rb_primary)

        match_layout.addWidget(self.rb_any)
        match_layout.addWidget(self.rb_primary)
        layout.addWidget(match_frame)

        self._populate_companies()

    def _populate_companies(self):
        companies = sorted(list(CompanySanitizer.SIGNATURE_SYNONYMS_REGEX.keys()), key=lambda s: s.lower())
        for comp in companies:
            item = QListWidgetItem(comp, self.list_widget)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)

    def _filter_list(self, text: str):
        query = text.strip().lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setHidden(query not in item.text().lower())

    def _set_all_checked(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(state)

    def get_selected_companies(self) -> set:
        return {
            self.list_widget.item(i).text()
            for i in range(self.list_widget.count())
            if self.list_widget.item(i).checkState() == Qt.Checked
        }

    def is_primary_only(self) -> bool:
        return self.rb_primary.isChecked()


class ContributionSearchWorker(QThread):
    """Background worker executing parallelized retrieval using TDocsParser JSON caching."""
    stage_progress = pyqtSignal(str, int, int)
    log_msg = pyqtSignal(str)
    results_ready = pyqtSignal(list)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        wgs: list,
        date_from: str,
        date_to: str,
        target_companies: set,
        primary_only: bool,
        target_wis: list,
        bypass_cache: bool,
        cache_dir: str,
        parent=None
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.wgs = wgs
        self.date_from = date_from
        self.date_to = date_to
        self.target_companies = target_companies
        self.primary_only = primary_only
        self.target_wis = [w.strip().lower() for w in target_wis if w.strip()]
        self.bypass_cache = bypass_cache
        self.cache_dir = Path(cache_dir)
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            # === STAGE 1: Resolve matching meetings ===
            self.stage_progress.emit("Stage 1/3: Resolving matching meetings...", 0, 100)
            self.log_msg.emit("🔍 Querying meetings database for target criteria...")

            meetings = self.meetings_db.search_meetings(
                wg_name=self.wgs if self.wgs else None,
                date_from=self.date_from,
                date_to=self.date_to
            )

            total_meetings = len(meetings)
            if total_meetings == 0:
                self.log_msg.emit("⚠️ No meetings found matching the selected criteria.")
                self.results_ready.emit([])
                self.finished.emit(True, "No meetings found.")
                return

            self.log_msg.emit(f"✅ Found {total_meetings} meeting(s) to process.")

            all_matched_tdocs = []
            processed_count = 0

            # === STAGE 2: Bounded Concurrency (Laptop-Safe: max 3 workers) ===
            max_workers = min(3, os.cpu_count() or 2)
            self.log_msg.emit(f"⚡ Starting worker pool ({max_workers} threads)...")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_meeting = {
                    executor.submit(self._process_single_meeting, mtg): mtg
                    for mtg in meetings
                }

                for future in as_completed(future_to_meeting):
                    if self._is_cancelled:
                        self.log_msg.emit("🛑 Operation cancelled by user.")
                        executor.shutdown(wait=False, cancel_futures=True)
                        break

                    processed_count += 1
                    status_desc = f"Stage 2/3: Processing ({processed_count}/{total_meetings})..."
                    self.stage_progress.emit(status_desc, processed_count, total_meetings)

                    try:
                        matches = future.result()
                        all_matched_tdocs.extend(matches)
                    except Exception as e:
                        mtg_ref = future_to_meeting[future]
                        self.log_msg.emit(f"❌ Error in meeting {mtg_ref.get('meeting_number')}: {e}")

            # === STAGE 3: Final Aggregation & Sort ===
            self.stage_progress.emit("Stage 3/3: Aggregating results...", 100, 100)
            all_matched_tdocs.sort(key=lambda r: (r.get("Meeting", ""), r.get("TDoc", "")))

            summary = f"Audit complete! Found {len(all_matched_tdocs)} contributions across {total_meetings} meeting(s)."
            self.log_msg.emit(f"🏁 {summary}")
            self.results_ready.emit(all_matched_tdocs)
            self.finished.emit(True, summary)

        except Exception as e:
            err = f"Error during contribution search: {e}"
            logging.error(err, exc_info=True)
            self.log_msg.emit(f"❌ {err}")
            self.finished.emit(False, str(e))

    def _process_single_meeting(self, mtg: dict) -> list:
        mtg_num = mtg.get("meeting_number", "Unknown")
        wg_name = mtg.get("wg_name", "")
        mtg_id = mtg.get("mtg_id")
        tag = f"{wg_name} #{mtg_num}"
        folder_name = mtg.get("folder_name") or mtg_num

        local_mtg_dir = self.cache_dir / folder_name
        agenda_dir = local_mtg_dir / "Agenda"
        excel_file = None

        if agenda_dir.exists():
            excel_file = next(
                (f for f in agenda_dir.iterdir()
                 if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                 and f.name.endswith(".xlsx")),
                None
            )
            if not excel_file and mtg_id:
                fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                if fallback.exists():
                    excel_file = fallback

        # If bypass requested or file is absent, download from 3GPP
        if self.bypass_cache or not excel_file:
            if not mtg_id:
                self.log_msg.emit(f"⚠️ [{tag}] Missing 3GPP portal ID; skipping download.")
                return []

            self.log_msg.emit(f"📥 [{tag}] Downloading TDocs workbook from 3GPP...")
            dl = TDocsDownloaderThread(mtg_id, local_mtg_dir)
            dl.run()

            if agenda_dir.exists():
                excel_file = next(
                    (f for f in agenda_dir.iterdir()
                     if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                     and f.name.endswith(".xlsx")),
                    None
                )
                if not excel_file:
                    fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                    if fallback.exists():
                        excel_file = fallback

        if not excel_file or not excel_file.exists():
            self.log_msg.emit(f"⚠️ [{tag}] No TDocs list available.")
            return []

        # TDocsParser automatically checks <excel_file>.json cache first
        json_cache = str(excel_file) + ".json"
        if not self.bypass_cache and os.path.exists(json_cache):
            self.log_msg.emit(f"⚡ [{tag}] Loaded via JSON cache ({excel_file.name}).")
        else:
            self.log_msg.emit(f"⚙️ [{tag}] Ingesting spreadsheet ({excel_file.name})...")

        tdocs_data = TDocsParser.parse_tdocs_excel(str(excel_file))

        matched = []
        for row in tdocs_data:
            if self._matches_filter(row):
                row_copy = dict(row)
                row_copy["Meeting"] = tag
                row_copy["docs_folder_url"] = mtg.get("docs_folder_url", "")
                matched.append(row_copy)

        self.log_msg.emit(f"   ↳ [{tag}] {len(matched)} matching contribution(s).")
        return matched

    def _matches_filter(self, row: dict) -> bool:
        source_raw = str(row.get("Source", "")).strip()

        # Company filter
        if self.target_companies:
            if not source_raw:
                return False

            if self.primary_only:
                parts = re.split(r'[,;]|\band\b', source_raw, maxsplit=1)
                first_author = parts[0] if parts else source_raw
                matched_companies = set(CompanySanitizer.get_matching_contributors(first_author))
            else:
                matched_companies = set(CompanySanitizer.get_matching_contributors(source_raw))

            intersect = matched_companies & self.target_companies
            if not intersect:
                return False

            row["Matched Company"] = ", ".join(sorted(list(intersect)))
        else:
            all_contributors = CompanySanitizer.get_matching_contributors(source_raw)
            row["Matched Company"] = ", ".join(all_contributors) if all_contributors else source_raw

        # Work Item filter
        if self.target_wis:
            wi_val = ""
            for key in row.keys():
                if "WORK ITEM" in key.upper() or key.upper() in ["WI", "WID"]:
                    wi_val = str(row.get(key, "")).strip().lower()
                    break

            title_val = str(row.get("Title", "")).strip().lower()

            wi_hit = any((token in wi_val or token in title_val) for token in self.target_wis)
            if not wi_hit:
                return False

        return True


class ContributionResultsTableModel(QAbstractTableModel):
    COLUMNS = [
        "Meeting", "TDoc", "Title", "Source", "Matched Company",
        "Type", "For", "Agenda Item", "TDoc Status", "Work Item"
    ]

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self.COLUMNS)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.COLUMNS[section]
        return None

    def data(self, index, role):
        if not index.isValid() or not (0 <= index.row() < len(self._data)):
            return None

        row_dict = self._data[index.row()]
        col_name = self.COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            if col_name == "Work Item":
                for k in ["Work Item", "WI", "WID", "Work Item / Study Item"]:
                    if k in row_dict and row_dict[k]:
                        return str(row_dict[k])
                return ""
            return str(row_dict.get(col_name, ""))

        if role == Qt.TextAlignmentRole:
            if col_name in ["Meeting", "TDoc", "Type", "For", "Agenda Item", "TDoc Status"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.ForegroundRole:
            if col_name == "TDoc":
                return QBrush(QColor("#E20074"))
            if col_name == "Matched Company":
                return QBrush(QColor("#0F766E"))

        if role == Qt.FontRole:
            if col_name in ["TDoc", "Matched Company"]:
                f = QFont()
                f.setBold(True)
                return f

        if role == Qt.UserRole:
            return row_dict

        return None

    def update_data(self, new_data: list):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()


class ContributionReportDialog(QDialog):
    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.worker = None
        self.exporter_thread = None

        self.setWindowTitle("3GPP Company Contribution Search & Report")
        self.resize(1240, 780)
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal)

        # Left Filter Sidebar
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

        # Right Panel: Results & Logs
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.setTextVisible(False)
        right_layout.addWidget(self.progress_bar)

        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMaximumHeight(90)
        self.log_viewer.setStyleSheet("background-color: #0F172A; color: #38BDF8; font-family: monospace; font-size: 11px;")
        right_layout.addWidget(self.log_viewer)

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
        self.table_view.doubleClicked.connect(self._open_selected_tdoc)
        right_layout.addWidget(self.table_view)

        footer_layout = QHBoxLayout()
        self.lbl_count = QLabel("0 contributions found.")
        self.lbl_count.setStyleSheet("color: #64748B; font-weight: bold;")

        self.btn_export = QPushButton("📥 Export to Excel...")
        self.btn_export.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_to_excel)

        footer_layout.addWidget(self.lbl_count)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_export)
        right_layout.addLayout(footer_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([390, 850])
        main_layout.addWidget(splitter)

    def set_active_wgs(self, wgs: list):
        """Pre-seeds the Working Group filter with active selections."""
        for i in range(1, self.wg_filter.model().rowCount()):
            item = self.wg_filter.model().item(i)
            if item:
                item.setCheckState(Qt.Checked if item.text() in wgs else Qt.Unchecked)
        self.wg_filter.update()

    def _start_search(self):
        selected_wgs = self.wg_filter.getCheckedItems()
        target_companies = self.company_widget.get_selected_companies()
        primary_only = self.company_widget.is_primary_only()
        target_wis = self.wi_token_widget.get_tokens()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")
        bypass = self.chk_bypass_cache.isChecked()

        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export.setEnabled(False)
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
        self.lbl_count.setText(f"Found {len(results)} matching contribution(s).")
        self.btn_export.setEnabled(len(results) > 0)
        self.table_view.resizeColumnsToContents()

    def _on_search_finished(self, success: bool, msg: str):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("3GPP Company Contribution Search & Report")
        if not success:
            QMessageBox.warning(self, "Search Notice", msg)

    def _open_selected_tdoc(self, index):
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        tdoc_id = row_data.get("TDoc")
        docs_url = row_data.get("docs_folder_url")
        if tdoc_id and docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            webbrowser.open(f"{full_url.rstrip('/')}/{tdoc_id}.zip")

    def _export_to_excel(self):
        data = self.table_model._data
        if not data:
            return

        default_name = f"Contributions_Report_{QDate.currentDate().toString('yyyy-MM-dd')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Contributions Report",
            str(Path.home() / "Desktop" / default_name),
            "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        self.btn_export.setEnabled(False)
        self.btn_export.setText("⏳ Exporting...")

        mtg_info = {"wg_name": "Contributions", "meeting_number": "Audit"}
        selected_cols = ContributionResultsTableModel.COLUMNS

        self.exporter_thread = ExcelExporterThread(
            output_path=Path(save_path),
            rows_data=data,
            selected_columns=selected_cols,
            mtg_info=mtg_info,
            docs_ftp_url="",
            auto_open=True
        )

        self.exporter_thread.finished.connect(self._on_export_finished)
        self.exporter_thread.start()

    def _on_export_finished(self, success: bool, msg: str):
        self.btn_export.setEnabled(True)
        self.btn_export.setText("📥 Export to Excel...")

        if success:
            reply = QMessageBox.question(
                self, "Export Complete",
                f"Report saved successfully:\n{msg}\n\nOpen spreadsheet now?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                try:
                    os.startfile(msg) if hasattr(os, 'startfile') else webbrowser.open(f"file:///{msg}")
                except Exception as e:
                    logging.error(f"Failed to open Excel report: {e}")
        else:
            QMessageBox.critical(self, "Export Failed", f"Could not generate Excel report:\n{msg}")