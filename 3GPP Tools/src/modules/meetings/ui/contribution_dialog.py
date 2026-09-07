# --- File: src/modules/meetings/ui/contribution_dialog.py ---
import os
import re
import logging
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QAbstractTableModel, QModelIndex, QPoint
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QLabel, QLineEdit, QPushButton, QCheckBox, QRadioButton,
    QButtonGroup, QDateEdit, QListWidget, QListWidgetItem,
    QTableView, QHeaderView, QTextEdit, QProgressBar,
    QMessageBox, QFileDialog, QFrame, QCompleter, QMenu, QApplication
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER
)
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.stats.contribution_stats import ContributionStatsExporterThread
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import ReadOnlyViewerDialog, StatisticsSettingsDialog
from modules.work_items.core.wi_database import WorkItemsDatabase


class KPICard(QFrame):
    """Lightweight summary metric card."""

    def __init__(self, title: str, default_val: str = "-", subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
            QLabel#kpiTitle {
                color: #64748B;
                font-size: 10px;
                font-weight: bold;
                border: none;
            }
            QLabel#kpiValue {
                color: #1E5C99;
                font-size: 16px;
                font-weight: bold;
                border: none;
            }
            QLabel#kpiSub {
                color: #94A3B8;
                font-size: 10px;
                border: none;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(1)

        self.lbl_title = QLabel(title.upper())
        self.lbl_title.setObjectName("kpiTitle")
        self.lbl_value = QLabel(default_val)
        self.lbl_value.setObjectName("kpiValue")
        self.lbl_sub = QLabel(subtitle)
        self.lbl_sub.setObjectName("kpiSub")

        layout.addWidget(self.lbl_title)
        layout.addWidget(self.lbl_value)
        layout.addWidget(self.lbl_sub)

    def set_value(self, val: str, sub: str = ""):
        self.lbl_value.setText(val)
        if sub:
            self.lbl_sub.setText(sub)


class WITokenChip(QFrame):
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
        self.input_edit.setPlaceholderText("Search WI acronym or code (e.g., FS_6G_ARC)...")
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
        self._executor = None

    def cancel(self):
        self._is_cancelled = True
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass

    def run(self):
        try:
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

            max_workers = min(3, os.cpu_count() or 2)
            self.log_msg.emit(f"⚡ Starting worker pool ({max_workers} threads)...")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                self._executor = executor
                future_to_meeting = {
                    executor.submit(self._process_single_meeting, mtg): mtg
                    for mtg in meetings
                }

                for future in as_completed(future_to_meeting):
                    if self._is_cancelled:
                        self.log_msg.emit("🛑 Operation cancelled by user.")
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

            self._executor = None
            if self._is_cancelled:
                self.finished.emit(False, "Search cancelled.")
                return

            self.stage_progress.emit("Stage 3/3: Aggregating results...", 100, 100)
            all_matched_tdocs.sort(key=lambda r: (r.get("WG", ""), r.get("Meeting", ""), r.get("TDoc", "")))

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
        if self._is_cancelled:
            return []

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

        json_cache = str(excel_file) + ".json"
        if not self.bypass_cache and os.path.exists(json_cache):
            self.log_msg.emit(f"⚡ [{tag}] Loaded via JSON cache ({excel_file.name}).")
        else:
            self.log_msg.emit(f"⚙️ [{tag}] Ingesting spreadsheet ({excel_file.name})...")

        tdocs_data = TDocsParser.parse_tdocs_excel(str(excel_file))

        matched = []
        for row in tdocs_data:
            resolved_wi = self._extract_related_wis(row)
            row["Related WIs"] = resolved_wi
            row["WG"] = wg_name

            if self._matches_filter(row):
                row_copy = dict(row)
                row_copy["WG"] = wg_name
                row_copy["Meeting"] = tag
                row_copy["docs_folder_url"] = mtg.get("docs_folder_url", "")
                row_copy["end_date"] = mtg.get("end_date", "")
                row_copy["start_date"] = mtg.get("start_date", "")
                row_copy["wg_name"] = wg_name
                row_copy["Related WIs"] = resolved_wi
                matched.append(row_copy)

        self.log_msg.emit(f"   ↳ [{tag}] {len(matched)} matching contribution(s).")
        return matched

    def _extract_related_wis(self, row: dict) -> str:
        # 1. Primary Check: Official 3GPP 'Related WIs' field[cite: 32]
        for key in ["Related WIs", "Related WI", "Related WI(s)", "Work Item", "Work Items", "WI", "WID"]:
            val = str(row.get(key, "")).strip()
            if val and val.lower() not in ["", "none", "unknown", "-"]:
                items = [w.strip() for w in re.split(r'[,;]+', val) if w.strip()]
                if items:
                    return ", ".join(items)

        # 2. Instant Fallback: Check 'Agenda item description' for parenthesized acronym[cite: 32]
        desc = str(row.get("Agenda item description", "")).strip()
        if desc:
            match = re.search(r'\(([A-Za-z0-9_ -]{3,35})\)', desc)
            if match:
                candidate = match.group(1).strip()
                if '_' in candidate or any(c.isdigit() for c in candidate):
                    return candidate

        # 3. Instant Fallback: Check Title for bracketed tags [WI_TAG][cite: 6]
        title = str(row.get("Title", "")).strip()
        if title:
            match = re.search(r'\[([A-Za-z0-9_ -]{3,35})\]', title)
            if match:
                candidate = match.group(1).strip()
                if candidate.lower() not in ["draft", "reply", "update", "revision", "discussion", "pcr", "cr", "ls", "none"]:
                    if '_' in candidate or any(c.isdigit() for c in candidate):
                        return candidate

        return ""

    def _matches_filter(self, row: dict) -> bool:
        source_raw = str(row.get("Source", "")).strip()

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

        if self.target_wis:
            wi_val = str(row.get("Related WIs", "")).strip().lower()
            title_val = str(row.get("Title", "")).strip().lower()
            desc_val = str(row.get("Agenda item description", "")).strip().lower()

            wi_hit = any((token in wi_val or token in title_val or token in desc_val) for token in self.target_wis)
            if not wi_hit:
                return False

        return True


class ContributionResultsTableModel(QAbstractTableModel):
    """Clean, neutral table model without color tints, matching application style."""
    COLUMNS = [
        "WG", "Meeting", "TDoc", "Title", "Source", "Matched Company",
        "Type", "For", "Agenda Item", "TDoc Status", "Related WIs", "Abstract"
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
            return str(row_dict.get(col_name, ""))

        if role == Qt.TextAlignmentRole:
            if col_name in ["WG", "Meeting", "TDoc", "Type", "For", "Agenda Item", "TDoc Status", "Related WIs"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.ToolTipRole:
            val = str(row_dict.get(col_name, "")).strip()
            if col_name == "Abstract" and val:
                return f"<div style='width: 450px; white-space: pre-wrap;'>{val}</div>"
            elif col_name in ["Title", "Source"] and len(val) > 40:
                return f"<div style='width: 400px; white-space: pre-wrap;'>{val}</div>"
            return None

        if role == Qt.UserRole:
            return row_dict

        return None

    def update_data(self, new_data: list):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()


class ContributionReportDialog(QDialog):
    """Modeless cross-meeting contribution report dialog."""
    open_meeting_requested = pyqtSignal(str, str)  # (wg_name, meeting_number)

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        # Pass Qt.Window flag so it acts as an independent modeless desktop window
        super().__init__(parent, Qt.Window)
        self.setModal(False)

        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.worker = None
        self.exporter_thread = None
        self.stats_thread = None

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
        self.kpi_agree = KPICard("Win / Agreement Rate", "0.0%", "Agreed / Approved")
        self.kpi_joint = KPICard("Collaboration", "0% Joint", "Co-authored")
        self.kpi_partner = KPICard("Top Partner Vendor", "None", "Most frequent ally")
        self.kpi_wi = KPICard("Top Work Item", "None", "Highest volume WI")

        kpi_layout.addWidget(self.kpi_total)
        kpi_layout.addWidget(self.kpi_agree)
        kpi_layout.addWidget(self.kpi_joint)
        kpi_layout.addWidget(self.kpi_partner)
        kpi_layout.addWidget(self.kpi_wi)
        right_layout.addLayout(kpi_layout)

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

        # Connect double-click and custom context menu
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
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")
        bypass = self.chk_bypass_cache.isChecked()

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
            self.kpi_agree.set_value("0.0%", "0 agreed")
            self.kpi_joint.set_value("0% Joint", "0 joint")
            self.kpi_partner.set_value("None", "No allies")
            self.kpi_wi.set_value("None", "No WIs")
            self.btn_export.setEnabled(False)
            self.btn_stats.setEnabled(False)
            return

        # 1. Total KPI & WG Coverage
        wgs_count = len(set(r.get("WG", "") for r in results if r.get("WG")))
        self.kpi_total.set_value(f"{total} TDocs", f"Across {wgs_count} WG(s)")

        # 2. Agreement & Consensus Rate KPI
        agreed_count = sum(1 for r in results if any(w in str(r.get("TDoc Status", "")).lower() for w in ["agreed", "approved"]))
        gross_agree_pct = round((agreed_count / total) * 100, 1)

        decided_success = sum(1 for r in results if any(w in str(r.get("TDoc Status", "")).lower() for w in ["agreed", "approved", "merged", "endorsed"]))
        total_decided = sum(1 for r in results if any(w in str(r.get("TDoc Status", "")).lower() for w in ["agreed", "approved", "merged", "endorsed", "noted", "postponed"]))
        consensus_pct = round((decided_success / total_decided) * 100, 1) if total_decided else 0

        self.kpi_agree.set_value(f"{consensus_pct}% Consensus", f"{gross_agree_pct}% gross ({agreed_count}/{total})")

        # 3. Collaboration KPI & 4. Top Partner Vendor
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

        # 5. Top Work Item KPI (splits multi-entry Related WIs)
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

    # --- ROW ACTIONS: OPEN MEETING & OPEN TDOC ---

    def _request_open_meeting(self, row_data: dict):
        """Requests opening the full meeting window in the main application."""
        wg = str(row_data.get("WG", "")).strip()
        mtg_raw = str(row_data.get("Meeting", "")).strip()
        mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
        if wg and mtg_num:
            self.open_meeting_requested.emit(wg, mtg_num)
        else:
            QMessageBox.warning(self, "Missing Info", "Cannot identify the Working Group or meeting number for this row.")

    def _open_tdoc(self, row_data: dict):
        """Downloads and opens the document ZIP file from the meeting folder."""
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return

        docs_url = str(row_data.get("docs_folder_url", "")).strip()
        if not docs_url:
            wg = row_data.get("WG", "")
            mtg_raw = row_data.get("Meeting", "")
            mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
            mtgs = self.meetings_db.search_meetings(wg_name=[wg] if wg else None, search_term=mtg_num)
            if mtgs:
                docs_url = mtgs[0].get("docs_folder_url", "")

        if docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            target_file_url = f"{full_url.rstrip('/')}/{tdoc_id}.zip"
            webbrowser.open(target_file_url)
        else:
            QMessageBox.warning(self, "Missing URL", f"No documents folder URL available for {tdoc_id}.")

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

        # Double-clicking any other cell opens the TDoc document
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
        wg = row_data.get("WG", "")
        mtg = row_data.get("Meeting", "")
        tdoc = row_data.get("TDoc", "")

        # 1. Open Meeting Table
        act_open_mtg = menu.addAction(f"🗓️ Open Meeting Table ({mtg})")
        act_open_mtg.triggered.connect(lambda: self._request_open_meeting(row_data))

        menu.addSeparator()

        # 2. Open / View TDoc
        if tdoc:
            act_open_tdoc = menu.addAction(f"📄 Open TDoc Document ({tdoc})")
            act_open_tdoc.triggered.connect(lambda: self._open_tdoc(row_data))

            act_portal = menu.addAction(f"🌐 View {tdoc} on 3GU Portal")
            act_portal.triggered.connect(lambda: webbrowser.open(f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={tdoc}"))

            menu.addSeparator()
            act_copy_tdoc = menu.addAction(f"📋 Copy TDoc Number ({tdoc})")
            act_copy_tdoc.triggered.connect(lambda: QApplication.clipboard().setText(tdoc))

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

    def _generate_statistics(self):
        data = self.table_model._data
        if not data:
            return

        self.btn_stats.setEnabled(False)
        self.btn_stats.setText("⏳ Generating...")

        config = StatisticsSettingsDialog().load_config()
        current_cache = self.settings.cache_dir
        export_dir = Path(current_cache) / "Contributions_Audit" / "Export"

        self.stats_thread = ContributionStatsExporterThread(
            export_dir=export_dir,
            tdocs_data=data,
            target_companies=self.company_widget.get_selected_companies(),
            target_wis=self.wi_token_widget.get_tokens(),
            config=config,
            parent=self
        )
        self.stats_thread.finished.connect(self._on_stats_finished)
        self.stats_thread.start()

    def _on_stats_finished(self, success: bool, msg: str):
        self.btn_stats.setEnabled(True)
        self.btn_stats.setText("📊 Statistics Dashboard...")

        if success:
            QMessageBox.information(self, "Dashboard Ready", f"Analytics dashboard generated successfully!\n\nSaved to:\n{msg}")
            try:
                os.startfile(msg) if hasattr(os, 'startfile') else webbrowser.open(f"file:///{msg}")
            except Exception as e:
                logging.error(f"Could not open dashboard: {e}")
        else:
            QMessageBox.warning(self, "Generation Failed", f"Could not generate statistics dashboard:\n{msg}")