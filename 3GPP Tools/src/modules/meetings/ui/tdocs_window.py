# --- File: src/modules/meetings/ui/tdocs_window.py ---
import datetime
import json
import logging
import os
import re
import urllib.parse
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QFrame,
                             QPushButton, QMessageBox, QMenu, QApplication,
                             QToolTip, QCheckBox, QDialog, QTextEdit, QFileDialog)

from core.network.network_state import NetworkState
from modules.emails.ui.email_window import EmailManagerWindow
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.llm_exporter import LLMExporterThread
from modules.meetings.core.markdown_exporter import MarkdownExporterThread
from modules.meetings.core.stats.exporter_thread import StatisticsExporterThread
from modules.meetings.core.tdocs_db import TDocsDatabase
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_threads import (
    TDocsRevisionsFetcherThread,
    TDocActionThread,
    TdocsByAgendaThread,
    WordAgendaImporterThread,
)
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdoc_delegates import HtmlDelegate, TDocActionDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import (
    ReadOnlyViewerDialog,
    InteractiveNotesDialog,
    StatisticsSettingsDialog, ExcelExportDialog,
    TDocInfoDialog
)
from modules.meetings.ui.tdocs_menus import build_action_menu, build_related_menu, build_row_context_menu
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel, natural_sort_key
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.ui.general_email_dialog import (
    GeneralEmailFoldersDialog, GeneralEmailSyncDialog, TDocEmailsDialog, load_wg_email_config
)


class DropOverlayWidget(QWidget):
    """Semi-transparent visual overlay displayed when hovering files over the window."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.hide()

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: rgba(240, 248, 255, 0.94);
                border: 3px dashed #005A9E;
                border-radius: 16px;
                padding: 30px;
            }
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignCenter)
        card_layout.setSpacing(10)

        icon_lbl = QLabel("📥")
        icon_lbl.setStyleSheet("font-size: 48px; border: none; background: transparent;")
        icon_lbl.setAlignment(Qt.AlignCenter)

        text_lbl = QLabel("<b>Drop Chairman's Notes / Agenda file(s) to import</b><br>"
                          "<span style='font-size: 12px; color: #555;'>"
                          "Supports single or multiple <code>.docx</code>, <code>.doc</code>, and <code>.htm</code> files</span>")
        text_lbl.setStyleSheet("font-size: 16px; color: #005A9E; border: none; background: transparent;")
        text_lbl.setAlignment(Qt.AlignCenter)

        card_layout.addWidget(icon_lbl)
        card_layout.addWidget(text_lbl)
        layout.addWidget(card)


def _open_folder(p: Path):
    if p.exists():
        os.startfile(str(p)) if hasattr(os, 'startfile') else webbrowser.open(f"file:///{p}")
    else:
        QMessageBox.warning(None, "Not Found", "Target file/folder does not exist yet.")


class TDocsWindow(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}
        self._email_dialogs = {}

        # Initialize database handle and email thread tracking
        self.general_email_db = GeneralEmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.general_email_sync_thread = None

        # Import Queue State
        self._import_queue = []
        self._is_importing_agenda = False
        self._total_import_batch_count = 0
        self._import_accumulated_merged = 0
        self._import_success_files = []
        self._import_failed_files = []

        # Enable Drag & Drop
        self.setAcceptDrops(True)

        self.db = TDocsDatabase(self.meeting_dir / "Agenda" / "user_tdocs.db")
        user_data = self.db.get_all()

        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        self.is_sa2 = ('SA2' in wg_name)
        is_electronic = bool(self.mtg_info.get('is_electronic', 0))
        self.is_sa2_electronic = self.is_sa2 and is_electronic

        main_ftp = self.mtg_info.get("url_key", "")
        if main_ftp and not main_ftp.startswith("http"):
            main_ftp = "https://www.3gpp.org/ftp/" + main_ftp.lstrip('/')
        self.main_ftp_url = main_ftp

        docs_ftp = self.mtg_info.get("docs_folder_url", "")
        if docs_ftp and not docs_ftp.startswith("http"):
            docs_ftp = "https://www.3gpp.org/ftp/" + docs_ftp.lstrip('/')
        self.docs_ftp_url = docs_ftp

        self.revisions_url = self.main_ftp_url.rstrip('/') + '/INBOX/Revisions/' if (
                self.is_sa2_electronic and self.main_ftp_url) else ""

        mtg_icon = "💻" if is_electronic else "🤝"
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')} {mtg_icon}"
        self.setWindowTitle(title)
        self.resize(1400, 750)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        self._setup_header(main_layout, title, is_electronic, len(tdocs_data))
        self._setup_filters(main_layout)
        self._setup_table(main_layout, tdocs_data, user_data)
        self._setup_cache()

        # Visual Drop Overlay Setup
        self.drop_overlay = DropOverlayWidget(self)

    def _setup_header(self, layout, title, is_electronic, count):
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 18px; color: #333;")
        title_lbl.setToolTip("Electronic Meeting (eMeeting)" if is_electronic else "In-Person Meeting (Face-to-Face)")
        title_lbl.setCursor(Qt.WhatsThisCursor)

        self.routing_indicator = QLabel("⚪ Routing...")
        self.routing_indicator.setStyleSheet("font-weight: bold; padding: 2px 6px; border-radius: 4px;")

        self.routing_timer = QTimer(self)
        self.routing_timer.timeout.connect(self._update_routing_indicator)
        self.routing_timer.start(2000)
        self._update_routing_indicator()

        self.last_mod_lbl = QLabel(self._get_mod_date_str())
        self.last_mod_lbl.setStyleSheet("font-size: 11px; color: #999999; margin-right: 15px; font-style: italic;")

        def style_btn():
            return """
            QPushButton { 
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold; 
                border-radius: 6px; padding: 6px 12px; 
                color: #333333; background-color: #FFFFFF; border: 1px solid #CCCCCC; 
            }
            QPushButton:hover, QPushButton::menu-indicator { 
                background-color: #F0F4F8; border: 1px solid #005A9E; color: #005A9E;
            }
            """

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(style_btn())
        self.refresh_btn.setToolTip("Reload TDocs or fetch the latest revisions from the FTP.")

        refresh_menu = QMenu(self)
        refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        if self.is_sa2:
            refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)
            refresh_menu.addAction("📝 Import Word Document (.docx / .doc)...", self._import_word_agenda_dialog)
            refresh_menu.addAction("📥 Download Chairman's Notes", self._download_chair_notes)
        if self.is_sa2_electronic:
            refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
        self.refresh_btn.setMenu(refresh_menu)

        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setStyleSheet(style_btn())
        self.folder_btn.setToolTip(
            "Access local cache folders, on-site server pages, export reports, and remote FTP directories.")

        self.folder_menu = QMenu(self)
        self.folder_menu.aboutToShow.connect(self._populate_resources_menu)
        self.folder_btn.setMenu(self.folder_menu)

        self.excel_btn = QPushButton("📗 Original Excel")
        self.excel_btn.setStyleSheet(style_btn())
        self.excel_btn.setToolTip("Open the raw, underlying 3GPP Excel TDocs list.")
        self.excel_btn.clicked.connect(self._open_excel)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(style_btn())
        self.export_btn.setToolTip("Export the TDocs table to Excel, LLM Markdown corpus, or summary reports.")

        export_menu = QMenu(self)
        export_menu.addAction("📊 Export to Excel (.xlsx)...", self._open_excel_export_dialog)
        export_menu.addAction("🤖 Export Visible to LLM (Corpus)", self._export_llm_visible)
        export_menu.addAction("📝 Export Markdown Summary Reports", self._export_reports)
        self.export_btn.setMenu(export_menu)

        self.stats_btn = QPushButton("📊 Statistics")
        self.stats_btn.setStyleSheet(style_btn())
        self.stats_btn.setToolTip("Generate an interactive HTML statistics dashboard for this meeting.")
        self.stats_btn.clicked.connect(self._generate_statistics)

        self.stats_cfg_btn = QPushButton("⚙️")
        self.stats_cfg_btn.setStyleSheet(style_btn())
        self.stats_cfg_btn.setFixedWidth(35)
        self.stats_cfg_btn.setToolTip("Configure Statistics Parameters")
        self.stats_cfg_btn.clicked.connect(self._open_stats_config)

        # 📧 Emails Dropdown Menu
        self.email_btn = QPushButton("📧 Emails ▾")
        self.email_btn.setStyleSheet(style_btn())
        self.email_btn.setToolTip("Sync and inspect Outlook emails linked to TDocs and revisions.")

        email_menu = QMenu(self)
        email_menu.addAction("🔄 Sync Related Emails...", self._on_sync_general_emails)
        email_menu.addAction("⚙️ Configure Outlook Folders...", self._on_configure_email_folders)
        email_menu.addAction("✔️ Mark All Emails as Read", self._on_mark_all_emails_read)
        email_menu.addSeparator()
        email_menu.addAction("🗑️ Wipe Generic Emails Database...", self._on_wipe_generic_emails)
        if self.is_sa2_electronic:
            email_menu.addSeparator()
            email_menu.addAction("📊 Open eMeeting Email Manager (Dashboard)", self._open_email_manager)
        self.email_btn.setMenu(email_menu)

        self.count_lbl = QLabel(f"Showing {count} of {count} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        self.instant_fetch_input = QLineEdit()
        self.instant_fetch_input.setPlaceholderText("Instant Fetch...")
        self.instant_fetch_input.setToolTip(
            "Instantly fetch a TDoc, bypassing the table entirely. Press Enter to launch.")
        self.instant_fetch_input.setText(self._get_tdoc_prefix())
        self.instant_fetch_input.setFixedWidth(120)
        self.instant_fetch_input.returnPressed.connect(self._on_instant_fetch)

        self.instant_fetch_input.setFocus()
        self.instant_fetch_input.setCursorPosition(len(self.instant_fetch_input.text()))

        header_layout.addWidget(title_lbl)
        header_layout.addWidget(self.routing_indicator)
        header_layout.addStretch()

        header_layout.addWidget(QLabel("🚀"))
        header_layout.addWidget(self.instant_fetch_input)
        header_layout.addSpacing(15)

        header_layout.addWidget(self.last_mod_lbl)
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addWidget(self.export_btn)
        header_layout.addWidget(self.stats_btn)
        header_layout.addWidget(self.stats_cfg_btn)
        header_layout.addWidget(self.email_btn)
        header_layout.addSpacing(15)
        header_layout.addWidget(self.count_lbl)
        layout.addLayout(header_layout)

    def _populate_resources_menu(self):
        self.folder_menu.clear()

        # 1. Local Hard Drive Cache Folders
        self.folder_menu.addAction("📁 Local: Meeting Folder", self._open_meeting_folder)
        if self.is_sa2:
            self.folder_menu.addAction("📄 Local: TdocsByAgenda.htm", self._open_agenda_file)

        self.folder_menu.addSeparator()

        # 2. Diagnostics & Export Utilities
        self.folder_menu.addAction("⚠️ View Unmatched Companies", self._show_unmatched_companies)
        self.folder_menu.addAction("📝 Export Markdown Reports", self._export_reports)

        # 3. Local On-Site Server (10.10.10.10)
        if NetworkState.get_instance().is_local_active():
            self.folder_menu.addSeparator()
            wg_name = self.mtg_info.get("wg_name", "").upper()
            local_base = URLRouter._get_local_server_base(wg_name)

            self.folder_menu.addAction("🟢 Local Server: Main WG Folder", lambda u=local_base: webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Docs Folder", lambda u=f"{local_base}/Docs": webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Inbox Folder",
                                       lambda u=f"{local_base}/Inbox": webbrowser.open(u))
            if self.is_sa2:
                self.folder_menu.addAction("🟢 Local Server: Revisions Folder",
                                           lambda u=f"{local_base}/Inbox/Revisions": webbrowser.open(u))
                self.folder_menu.addAction("🟢 Local Server: TdocsByAgenda.htm",
                                           lambda u=f"{local_base}/TdocsByAgenda.htm": webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Home (10.10.10.10)",
                                       lambda: webbrowser.open("http://10.10.10.10/"))

        # 4. Remote Web & FTP Archive Paths
        self.folder_menu.addSeparator()
        if self.main_ftp_url:
            self.folder_menu.addAction("🌐 Web FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
        if self.docs_ftp_url:
            self.folder_menu.addAction("🌐 Web FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
        if self.revisions_url:
            self.folder_menu.addAction("🌐 Web FTP: Revisions Folder", lambda: webbrowser.open(self.revisions_url))

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, legacy_url: str = None,
                                 is_silent_compare: bool = False):
        self.model.set_loading(base_tdoc, True)

        is_active = self.mtg_info.get("is_active_sync", False)

        url_list = URLRouter.build_priority_url_list(
            self.mtg_info.get("wg_name", ""),
            self.mtg_info.get("folder_name") or self.mtg_info.get("meeting_number", ""),
            self.main_ftp_url,
            is_active,
            target_filename=target_filename
        )

        thread = TDocActionThread(base_tdoc, target_filename, url_list, self.meeting_dir,
                                  open_file=not is_silent_compare)
        thread.is_silent_compare = is_silent_compare
        thread.target_filename = target_filename
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))

        self.active_threads[base_tdoc] = thread
        thread.start()

    def _setup_filters(self, layout):
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self._apply_search_filter)

        filter_frame = QFrame()
        filter_frame.setStyleSheet(
            "QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; } "
            "QLabel { font-weight: bold; color: #555; border: none; } "
            "QLineEdit, QComboBox { padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; }"
        )
        filter_layout = QHBoxLayout(filter_frame)

        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search (e.g., "Baseline" -"discussed in call" or -draft)...')
        self.search_input.setToolTip(
            "Search across TDoc numbers, titles, sources, or abstracts.\n"
            "• Use quotes for exact phrases: \"KI#11\"\n"
            "• Prefix with '-' or '!' to exclude: -draft or -\"discussed in call\""
        )
        self.search_input.textChanged.connect(lambda _: self.search_timer.start())
        filter_layout.addWidget(self.search_input)

        self.company_combo = CheckableComboBox("Company")
        self.company_combo.setToolTip("Filter by contributing companies.")
        self.company_combo.selectionChanged.connect(self._on_company_changed)
        filter_layout.addWidget(self.company_combo)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setToolTip("Filter by document type (e.g., pCR, Discussion, Draft).")
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setToolTip("Filter by 3GPP Agenda Item (AI).")
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("TDoc Status")
        self.status_combo.setToolTip("Filter by document status (e.g., Agreed, Noted, Revised).")
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        self.my_status_combo = CheckableComboBox("My Status")
        self.my_status_combo.setToolTip("Filter by your personal color-coded status.")
        self.my_status_combo.selectionChanged.connect(self._on_my_status_changed)
        filter_layout.addWidget(self.my_status_combo)

        if self.is_sa2:
            self.chk_no_comments = QCheckBox("No Comments Only")
            self.chk_no_comments.setToolTip("Hide TDocs that have comments in the secretary's notes.")
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            filter_layout.addWidget(self.chk_no_comments)

        layout.addWidget(filter_frame)

    def _setup_table(self, layout, data, user_data):
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, data, user_data)
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.doubleClicked.connect(self._show_cell_popup)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet(
            "QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; } "
            "QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }"
        )
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(48)

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.html_delegate.linkRightClicked.connect(self._show_related_menu)
        self.table.setItemDelegateForColumn(7, self.html_delegate)
        self.table.setItemDelegateForColumn(12, self.html_delegate)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)
        header.resizeSection(1, 100)
        header.resizeSection(2, 200)
        header.resizeSection(3, 100)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 28)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        header.resizeSection(8, 90)
        header.setSectionResizeMode(9, QHeaderView.Fixed)
        header.resizeSection(9, 28)
        header.resizeSection(10, 80)
        header.resizeSection(12, 160)

        # Set column width for the 14th column ("Emails", index 13)
        if len(self.model._headers) > 13:
            header.resizeSection(13, 85)

        # Enable Right-Click Context Menu for row operations (Mark Read/Unread)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)

        layout.addWidget(self.table)
        self._refresh_comboboxes()
        self._refresh_email_counts()

        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

    def _setup_cache(self):
        agenda_dir = self.meeting_dir / "Agenda"
        if self.is_sa2:
            local_agenda = agenda_dir / "TdocsByAgenda.htm"
            if local_agenda.exists():
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(local_agenda))
                if agenda_data:
                    self.model.merge_agenda_data(agenda_data)
                    self._refresh_comboboxes()

        if self.is_sa2_electronic:
            local_revs = agenda_dir / "revisions.json"
            if local_revs.exists():
                try:
                    with open(local_revs, "r", encoding="utf-8") as f:
                        self.model.revisions = json.load(f)
                        self.model.dataChanged.emit(self.model.index(0, 0),
                                                    self.model.index(self.model.rowCount() - 1, 0))
                except Exception:
                    if self.revisions_url:
                        self._refresh_revisions(silent=True)
            else:
                if self.revisions_url:
                    self._refresh_revisions(silent=True)

    def _get_mod_date_str(self):
        try:
            return f"List last updated: {datetime.datetime.fromtimestamp(os.path.getmtime(self.filepath)).strftime('%Y-%m-%d %H:%M')}"
        except Exception:
            return "List last updated: Unknown"

    def _apply_search_filter(self):
        self.proxy.setGlobalFilter(self.search_input.text())
        QTimer.singleShot(0, self._update_count_label)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)
        QTimer.singleShot(0, self._update_count_label)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)
        QTimer.singleShot(0, self._update_count_label)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)
        QTimer.singleShot(0, self._update_count_label)

    def _on_my_status_changed(self, statuses):
        self.proxy.setMyStatusFilters(statuses)
        QTimer.singleShot(0, self._update_count_label)

    def _on_company_changed(self, companies):
        self.proxy.setCompanyFilters(companies)
        QTimer.singleShot(0, self._update_count_label)

    def _on_no_comments_toggled(self, checked):
        self.proxy.setNoCommentsFilter(checked)
        QTimer.singleShot(0, self._update_count_label)

    def _update_count_label(self):
        self.count_lbl.setText(f"Showing {self.proxy.rowCount()} of {self.model.rowCount()} TDocs")

    def _refresh_comboboxes(self):
        def sanitize(val): return str(val).strip() if val is not None else ""

        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in self.model._data)))
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in self.model._data)),
                            key=natural_sort_key)
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))
        unique_my_statuses = sorted(list(set(sanitize(r.get("My Status", "")) for r in self.model._data)))

        unique_companies = set()
        for r in self.model._data:
            unique_companies.update(r.get("_Sanitized_Companies", ["Other"]))
        sorted_companies = sorted(list(unique_companies), key=lambda x: x.lower())

        self.type_combo.updateItems(unique_types)
        self.ai_combo.updateItems(unique_ais)
        self.status_combo.updateItems(unique_statuses)
        self.company_combo.updateItems(sorted_companies)
        self.my_status_combo.updateItems(unique_my_statuses)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())

        QTimer.singleShot(0, self._update_count_label)

    def _clear_all_filters(self):
        self.search_input.blockSignals(True)
        self.search_input.clear()
        self.search_input.blockSignals(False)
        self.proxy.setGlobalFilter("")
        if self.is_sa2:
            self.chk_no_comments.blockSignals(True)
            self.chk_no_comments.setChecked(False)
            self.chk_no_comments.blockSignals(False)
            self.proxy.setNoCommentsFilter(False)

        for combo in [self.company_combo, self.type_combo, self.ai_combo, self.status_combo, self.my_status_combo]:
            combo.blockSignals(True)
            combo.model().item(0).setCheckState(Qt.Checked)
            for i in range(1, combo.model().rowCount()):
                combo.model().item(i).setCheckState(Qt.Checked)
            combo.updateText()
            combo.blockSignals(False)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())

        QTimer.singleShot(0, self._update_count_label)

    def _handle_tdoc_action(self, base_tdoc: str):
        if base_tdoc in self.model.loading_tdocs or not self.docs_ftp_url:
            return
        revisions = self.model.revisions.get(base_tdoc, [])
        build_action_menu(self.table, base_tdoc, self.docs_ftp_url, self.revisions_url, revisions, self.meeting_dir,
                          self._trigger_download_thread, self._export_llm_single, self._compose_email_draft,
                          QCursor.pos())

    def _show_related_menu(self, target_tdoc: str, pos: QPoint):
        build_related_menu(self, target_tdoc, self.model.valid_tdocs, self.docs_ftp_url, self.revisions_url,
                           self._scroll_to_tdoc, self._trigger_download_thread, self._export_llm_single,
                           self.global_action_requested.emit, self._compose_email_draft, pos)

    def _compose_email_draft(self, tdoc_id: str):
        """Generates a standardized 3GPP email subject and opens a new draft in the default email client."""
        row_data = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == tdoc_id.upper()),
                        None)

        # Fallback to base TDoc metadata if drafting for a newly introduced revision
        if not row_data:
            match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
            row_data = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc), {})

        wg = str(self.mtg_info.get("wg_name", "")).strip().upper()
        mtg_num = str(self.mtg_info.get("meeting_number", "")).strip()
        ai = str(row_data.get("Agenda Item", "")).strip()
        title = str(row_data.get("Title", "")).strip()

        # Build standardized 3GPP bracketed tag: e.g. [SA2#176, 20.6.1.1]
        tag_parts = []
        if wg and mtg_num:
            tag_parts.append(f"{wg}#{mtg_num}")
        elif wg or mtg_num:
            tag_parts.append(f"{wg}{mtg_num}")

        if ai and ai.upper() not in ["N/A", "NONE", "UNKNOWN", ""]:
            tag_parts.append(ai)

        tag_str = f"[{', '.join(tag_parts)}] " if tag_parts else ""
        subject = f"{tag_str}{tdoc_id} {title}".strip()

        encoded_subject = urllib.parse.quote(subject, safe='')
        mailto_url = f"mailto:?subject={encoded_subject}"

        try:
            if hasattr(os, 'startfile'):
                os.startfile(mailto_url)
            else:
                webbrowser.open(mailto_url)
            logging.info(f"📧 [Email Draft] Opened new email draft with subject: '{subject}'")
        except Exception as e:
            logging.warning(f"Failed to launch email client automatically: {e}")
            QApplication.clipboard().setText(subject)
            QToolTip.showText(QCursor.pos(), "📋 Subject copied to clipboard!", self)

    def _show_cell_popup(self, index):
        if not index.isValid():
            return
        col_name = self.model._headers[index.column()]

        # Intercept clicks on the "Emails" column to open the inspection dialog
        if col_name == "Emails":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            self._open_tdoc_emails(row_data.get("TDoc", ""))
            return

        # Double-click on TDoc column opens the complete Info Card Dialog
        if col_name == "TDoc":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            tdoc_id = str(row_data.get("TDoc", "")).strip()
            revs = self.model.revisions.get(tdoc_id, [])
            TDocInfoDialog(
                row_data,
                docs_ftp_url=self.docs_ftp_url,
                revisions=revs,
                parent=self,
            ).exec_()
            return

        if col_name not in [
            "Secretary Remarks",
            "Title",
            "Source",
            "Abstract",
            "My Notes",
            "My Status",
        ]:
            return

        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        tdoc_id = row_data.get("TDoc", "")

        if col_name in ["Title", "Source", "Abstract"]:
            val = next(
                (
                    str(v)
                    for k, v in row_data.items()
                    if str(k).strip().lower() == col_name.lower() and v
                ),
                "",
            )
            ReadOnlyViewerDialog(
                self, f"📄 Viewing: {col_name} ({tdoc_id})", val
            ).exec_()
        else:
            InteractiveNotesDialog(
                self, tdoc_id, row_data, self._save_user_data
            ).exec_()

    def _save_user_data(self, tdoc_id: str, status: str, notes: str):
        self.db.upsert(tdoc_id, status, notes)
        self.model.user_data = self.db.get_all()
        self.model.apply_user_data_refresh()
        self._refresh_comboboxes()

    def _export_reports(self):
        self.export_thread = MarkdownExporterThread(self.meeting_dir, self.model._data, self.docs_ftp_url,
                                                    self.mtg_info)
        self.export_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, False))
        self.export_thread.start()

    def _open_stats_config(self):
        StatisticsSettingsDialog(self).exec_()

    def _generate_statistics(self):
        self.stats_btn.setText("⏳ Generating...")
        self.stats_btn.setEnabled(False)

        config = StatisticsSettingsDialog().load_config()
        self.stats_thread = StatisticsExporterThread(self.meeting_dir, self.model._data, self.mtg_info, config)
        self.stats_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, True))
        self.stats_thread.start()

    def _on_export_finished(self, success: bool, msg: str, is_stats: bool):
        if is_stats:
            self.stats_btn.setText("📊 Statistics")
            self.stats_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Export Complete", f"Successfully generated:\n{msg}")
            if hasattr(os, 'startfile'):
                os.startfile(str(msg))
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    def _scroll_to_tdoc(self, target_tdoc: str):
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', target_tdoc, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else target_tdoc.upper()
        if base_tdoc in self.model.valid_tdocs:
            for row in range(self.proxy.rowCount()):
                if self.proxy.data(self.proxy.index(row, 1), Qt.UserRole) == base_tdoc:
                    self.table.scrollTo(self.proxy.index(row, 1), QTableView.PositionAtCenter)
                    self.table.selectRow(row)
                    return
            if QMessageBox.question(self, "Hidden TDoc",
                                    f"TDoc '{base_tdoc}' is hidden by active filters.\nClear filters to view?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._clear_all_filters()
                self._scroll_to_tdoc(target_tdoc)
        else:
            if QMessageBox.question(self, "External TDoc",
                                    f"{base_tdoc} is not from this meeting.\nSearch global database?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self.global_action_requested.emit(base_tdoc, 'open_meeting')

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc in self.active_threads:
            del self.active_threads[tdoc]
        self.model.set_loading(tdoc, False)
        if not success:
            return QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)
        if getattr(thread, "is_silent_compare", False):
            files = getattr(thread, "extracted_doc_paths", [])
            if files:
                ComparisonManager.get_instance().add_to_cart(thread.target_filename, str(files[0]))
            else:
                QMessageBox.warning(self, "Compare Failed", "No Word document found in TDoc ZIP.")

    def _refresh_both(self):
        self._refresh_excel()
        self._refresh_revisions(silent=True)

    def _refresh_excel(self):
        if not self.mtg_info.get("mtg_id"):
            return QMessageBox.warning(self, "Missing ID", "Cannot refresh: Missing 3GPP Portal ID.")
        self.refresh_btn.setText("⏳ Downloading...")
        self.refresh_btn.setEnabled(False)
        self.dl_thread = TDocsDownloaderThread(self.mtg_info.get("mtg_id"), self.meeting_dir, self)
        self.dl_thread.finished.connect(self._on_refresh_excel_finished)
        self.dl_thread.start()

    def _on_refresh_excel_finished(self, success: bool, result: str, mtg_id: str):
        self.refresh_btn.setText("🔄 Refresh")
        self.refresh_btn.setEnabled(True)
        if success:
            self.filepath = result
            if new_data := TDocsParser.parse_tdocs_excel(self.filepath):
                self.model.update_data(new_data)
                self._refresh_comboboxes()
                self._refresh_email_counts()
                self.last_mod_lbl.setText(self._get_mod_date_str())
            else:
                QMessageBox.warning(self, "Parse Error", "Downloaded, but could not parse the Excel file.")
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to refresh TDocs:\n{result}")

    def _refresh_revisions(self, silent=False):
        if not self.revisions_url:
            return
        self.rev_thread = TDocsRevisionsFetcherThread(self.revisions_url, self.meeting_dir)
        self.rev_thread.finished.connect(lambda s, d, m: self._on_revisions_fetched(s, d, m, silent))
        self.rev_thread.start()

    def _on_revisions_fetched(self, success: bool, data: dict, msg: str, silent: bool):
        if success:
            self.model.revisions = data
            self.model.dataChanged.emit(self.model.index(0, 0), self.model.index(self.model.rowCount() - 1, 0))
            if not silent:
                self.refresh_btn.setText(f"✅ {len(data)} Revs")
                QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        elif not silent:
            QMessageBox.warning(self, "Revisions Error", f"Failed to sync revisions:\n{msg}")

    def _fetch_tdocs_by_agenda(self):
        wg_name = self.mtg_info.get("wg_name", "").upper()
        main_url = self.mtg_info.get("url_key", "")
        if main_url and not main_url.startswith("http"):
            main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

        is_active = self.mtg_info.get("is_active_sync", False)

        candidate_urls = []
        if NetworkState.get_instance().is_local_active():
            local_base = URLRouter._get_local_server_base(wg_name)
            candidate_urls.append(local_base)

        if is_active:
            sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
            candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

        if main_url:
            candidate_urls.append(main_url)

        if not candidate_urls:
            return QMessageBox.warning(self, "No URL Available", "Cannot sync agenda: No valid meeting URL found.")

        self.refresh_btn.setText("⏳ Parsing HTML...")
        self.agenda_thread = TdocsByAgendaThread(candidate_urls, self.meeting_dir)
        self.agenda_thread.finished.connect(self._on_agenda_fetched)
        self.agenda_thread.start()

    def _on_agenda_fetched(self, success: bool, agenda_data: dict):
        if success and agenda_data:
            self.model.merge_agenda_data(agenda_data)
            self._refresh_comboboxes()
            self.refresh_btn.setText(f"✅ {len(agenda_data)} Merged")
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        else:
            self.refresh_btn.setText("🔄 Refresh")
            QMessageBox.warning(self, "Error", "Failed to parse TdocsByAgenda.htm.")

    def _open_meeting_folder(self):
        _open_folder(self.meeting_dir)

    def _show_unmatched_companies(self):
        unmatched = self.model.get_unmatched_sources()

        if not unmatched:
            QMessageBox.information(self, "All Matched",
                                    "Great news! Every source in this meeting was successfully matched by the CompanySanitizer.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Unmatched Companies ({len(unmatched)})")
        dialog.resize(500, 400)

        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "The following raw 'Source' strings were not recognized by the CompanySanitizer and were grouped as 'Other'.\n\nYou can copy these to update your REGEX dictionary:"))

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(unmatched))
        layout.addWidget(text_edit)

        btn = QPushButton("Close")
        btn.clicked.connect(dialog.accept)
        layout.addWidget(btn)

        dialog.exec_()

    def _open_agenda_file(self):
        _open_folder(self.meeting_dir / "Agenda" / "TdocsByAgenda.htm")

    def _open_excel(self):
        _open_folder(Path(self.filepath))

    def _open_email_manager(self):
        self.email_window = EmailManagerWindow(self.meeting_dir, {
            str(r.get("TDoc", "")).strip().upper(): str(r.get("Agenda Item", "N/A")).strip() for r in self.model._data
            if r.get("TDoc")}, self.mtg_info.get("start_date", ""), self.mtg_info.get("end_date", ""))

        self.email_window.tdoc_open_requested.connect(self._open_tdoc_from_signal)
        self.email_window.tdoc_jump_requested.connect(self._jump_to_tdoc_from_signal)
        self.email_window.show()

    def _open_tdoc_from_signal(self, tdoc_id: str):
        main_app_window = self.window()
        main_app_window.setWindowState(main_app_window.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        main_app_window.raise_()
        main_app_window.activateWindow()

        self._scroll_to_tdoc(tdoc_id)

        is_rev = re.search(r'(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        target_url = self.revisions_url if is_rev and self.revisions_url else self.docs_ftp_url

        if not target_url:
            QMessageBox.warning(self, "Missing URL", "No FTP URL available to download this document.")
            return

        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()

        self._trigger_download_thread(base_tdoc, tdoc_id, target_url, is_silent_compare=False)

    def _jump_to_tdoc_from_signal(self, tdoc_id: str):
        main_app_window = self.window()
        main_app_window.setWindowState(main_app_window.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        main_app_window.raise_()
        main_app_window.activateWindow()

        self._scroll_to_tdoc(tdoc_id)

    def _copy_table_selection(self):
        indexes = sorted(self.table.selectionModel().selectedIndexes(), key=lambda x: (x.row(), x.column()))
        if not indexes:
            return
        lines, current_line, current_row = [], [], indexes[0].row()
        for idx in indexes:
            if idx.row() != current_row:
                lines.append("\t".join(current_line))
                current_line = []
                current_row = idx.row()
            cell_text = str(idx.data(Qt.UserRole) or "").strip()
            if not cell_text:
                cell_text = str(idx.data(Qt.DisplayRole) or "").strip()
            current_line.append(cell_text)
        lines.append("\t".join(current_line))
        QApplication.clipboard().setText("\n".join(lines))
        QToolTip.showText(QCursor.pos(), "📋 Copied to clipboard!", self.table)

    def _export_llm_visible(self):
        visible_tdocs = []
        for r in range(self.proxy.rowCount()):
            index = self.proxy.index(r, 0)
            source_index = self.proxy.mapToSource(index)
            row_data = self.model._data[source_index.row()]
            visible_tdocs.append(row_data)

        if not visible_tdocs:
            return QMessageBox.warning(self, "Empty View", "No TDocs currently visible in the table to export.")

        self.export_btn.setText("⏳ Compiling Corpus...")
        self.export_btn.setEnabled(False)

        config = StatisticsSettingsDialog().load_config()
        max_chars = config.get("llm_max_chars", 200000)
        system_prompt = config.get("llm_system_prompt", "")

        self.llm_thread = LLMExporterThread(
            self.meeting_dir,
            visible_tdocs,
            self.docs_ftp_url,
            self.revisions_url,
            is_bulk=True,
            max_chars=max_chars,
            system_prompt=system_prompt
        )
        self.llm_thread.progress.connect(self._on_llm_progress)
        self.llm_thread.finished.connect(self._on_llm_export_finished)
        self.llm_thread.start()

    def _export_llm_single(self, tdoc_id: str):
        row_data = next((r for r in self.model._data if r.get("TDoc") == tdoc_id), None)
        if not row_data:
            return

        config = StatisticsSettingsDialog().load_config()
        max_chars = config.get("llm_max_chars", 200000)
        system_prompt = config.get("llm_system_prompt", "")

        self.llm_thread = LLMExporterThread(
            self.meeting_dir,
            [row_data],
            self.docs_ftp_url,
            self.revisions_url,
            is_bulk=False,
            max_chars=max_chars,
            system_prompt=system_prompt
        )
        self.llm_thread.progress.connect(self._on_llm_progress)
        self.llm_thread.finished.connect(lambda s, m: QMessageBox.information(self, "LLM Export", m))
        self.llm_thread.start()

    def _on_llm_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:35])

    def _on_llm_export_finished(self, success: bool, msg: str):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "LLM Export Complete", msg)
            _open_folder(self.meeting_dir / "Export" / "LLM_Corpus")
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    def _get_tdoc_prefix(self):
        first_tdoc = self.mtg_info.get("first_tdoc", "")
        if first_tdoc:
            match = re.match(r'^([A-Z0-9]+-\d{2})', first_tdoc.upper())
            if match:
                return match.group(1)

        wg = self.mtg_info.get("wg_name", "").upper()
        start_date = self.mtg_info.get("start_date", "")
        year_str = start_date[2:4] if len(start_date) >= 4 else datetime.datetime.now().strftime("%y")

        wg_map = {
            "SA1": "S1", "SA2": "S2", "SA3": "S3", "SA4": "S4", "SA5": "S5", "SA6": "S6", "SA": "SP",
            "RAN1": "R1", "RAN2": "R2", "RAN3": "R3", "RAN4": "R4", "RAN5": "R5", "RAN6": "R6", "RAN": "RP",
            "CT1": "C1", "CT3": "C3", "CT4": "C4", "CT6": "C6", "CT": "CP"
        }
        prefix = wg_map.get(wg, wg)
        return f"{prefix}-{year_str}"

    def _update_routing_indicator(self):
        net_state = NetworkState.get_instance()
        is_active = self.mtg_info.get("is_active_sync", False)

        if net_state.is_local_active():
            self.routing_indicator.setText("🟢 Local Server")
            self.routing_indicator.setStyleSheet(
                "color: #155724; background-color: #C8E6C9; font-weight: bold; padding: 2px 6px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the high-speed local 10.10.10.10 network.")
        elif is_active:
            self.routing_indicator.setText("🔵 SYNC Folder")
            self.routing_indicator.setStyleSheet(
                "color: #004085; background-color: #CCE5FF; font-weight: bold; padding: 2px 6px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the active meeting SYNC folder.")
        else:
            self.routing_indicator.setText("🌐 Standard Web")
            self.routing_indicator.setStyleSheet(
                "color: #383D41; background-color: #E2E3E5; font-weight: bold; padding: 2px 6px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the standard 3GPP web archive.")

    def _on_instant_fetch(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()

        logging.info(f"🚀 [Instant Fetch] Requested {target_filename}. Engaging Smart Router...")
        self._trigger_download_thread(base_tdoc, target_filename, legacy_url=None, is_silent_compare=False)

    # --- Drag & Drop Visual Events ---
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.setGeometry(self.rect())

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.htm', '.html']:
                    self.drop_overlay.setGeometry(self.rect())
                    self.drop_overlay.show()
                    self.drop_overlay.raise_()
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.htm', '.html']:
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.hide()
        event.accept()

    def dropEvent(self, event):
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.hide()

        valid_files = []
        for url in event.mimeData().urls():
            file_path = Path(url.toLocalFile())
            if file_path.suffix.lower() in ['.docx', '.doc', '.htm', '.html']:
                valid_files.append(file_path)

        if valid_files:
            event.acceptProposedAction()
            self._process_imported_agenda_files(valid_files)
        else:
            event.ignore()

    def _import_word_agenda_dialog(self):
        agenda_dir = self.meeting_dir / "Agenda"
        start_dir = str(agenda_dir) if agenda_dir.exists() else str(self.meeting_dir)

        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select TDocsByAgenda / Chairman's Notes File(s)",
            start_dir,
            "Supported Files (*.docx *.doc *.htm *.html);;Word Documents (*.docx *.doc);;DOCX Files (*.docx);;DOC Files (*.doc);;HTML Files (*.htm *.html);;All Files (*.*)"
        )
        if file_paths:
            self._process_imported_agenda_files([Path(p) for p in file_paths])

    def _process_imported_agenda_file(self, source_path: Path):
        """Single-file compatibility wrapper."""
        self._process_imported_agenda_files([source_path])

    def _process_imported_agenda_files(self, source_paths: list):
        """Queues single or multiple agenda/notes files for sequential processing."""
        valid_paths = [p for p in source_paths if p.exists()]
        if not valid_paths:
            QMessageBox.warning(self, "Files Not Found", "None of the specified files exist on disk.")
            return

        self._import_queue.extend(valid_paths)

        if not self._is_importing_agenda:
            self._total_import_batch_count = len(self._import_queue)
            self._import_accumulated_merged = 0
            self._import_success_files = []
            self._import_failed_files = []
            self._is_importing_agenda = True
            self._process_next_in_import_queue()

    def _process_next_in_import_queue(self):
        """Processes the next file in the queue or finalizes the batch import."""
        if not self._import_queue:
            self._finalize_import_batch()
            return

        current_file = self._import_queue.pop(0)
        current_idx = self._total_import_batch_count - len(self._import_queue)

        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setText(f"⏳ ({current_idx}/{self._total_import_batch_count}) Importing...")

        self.word_import_thread = WordAgendaImporterThread(current_file, self.meeting_dir)
        self.word_import_thread.progress.connect(
            lambda msg, idx=current_idx, tot=self._total_import_batch_count: self._on_word_import_progress(
                f"({idx}/{tot}) {msg}")
        )
        self.word_import_thread.finished.connect(self._on_word_import_finished)
        self.word_import_thread.start()

    def _on_word_import_progress(self, msg: str):
        self.refresh_btn.setText(f"⏳ {msg}"[:32])

    def _on_word_import_finished(self, success: bool, agenda_data: dict, filename: str, error_msg: str):
        if success and agenda_data:
            self.model.merge_agenda_data(agenda_data)
            self._import_accumulated_merged += len(agenda_data)
            self._import_success_files.append((filename, len(agenda_data)))
        else:
            self._import_failed_files.append((filename, error_msg or "No valid TDocs table found"))

        # Advance to the next item in the queue
        self._process_next_in_import_queue()

    def _finalize_import_batch(self):
        """Cleans up UI state and displays a completion summary dialog."""
        self._is_importing_agenda = False
        self.refresh_btn.setEnabled(True)
        self.refresh_btn.setText("🔄 Refresh")

        if self._import_accumulated_merged > 0:
            self._refresh_comboboxes()
            self.refresh_btn.setText(f"✅ {self._import_accumulated_merged} Merged")
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))

        # Single-file feedback
        if self._total_import_batch_count == 1:
            if self._import_success_files:
                fname, count = self._import_success_files[0]
                QMessageBox.information(
                    self,
                    "Import Successful",
                    f"Saved to Agenda/\nSuccessfully merged {count} TDocs from:\n{fname}",
                )
            else:
                fname, err = self._import_failed_files[0]
                QMessageBox.warning(self, "Import Failed", f"Failed to import {fname}:\n{err}")
            return

        # Multi-file batch feedback
        summary_lines = [
            f"Batch import completed for {self._total_import_batch_count} files.",
            f"• Successfully imported: {len(self._import_success_files)} file(s)",
            f"• Total TDocs merged: {self._import_accumulated_merged}\n"
        ]

        if self._import_failed_files:
            summary_lines.append(f"⚠️ Failures ({len(self._import_failed_files)}):")
            for fname, err in self._import_failed_files:
                summary_lines.append(f"  - {fname}: {err}")
            QMessageBox.warning(self, "Batch Import Completed with Warnings", "\n".join(summary_lines))
        else:
            QMessageBox.information(self, "Batch Import Successful", "\n".join(summary_lines))

    def _download_chair_notes(self):
        """Asynchronously downloads all Chairman's Notes into Agenda/ChairNotes/."""
        wg_name = self.mtg_info.get("wg_name", "").upper()
        main_url = self.mtg_info.get("url_key", "")
        if main_url and not main_url.startswith("http"):
            main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

        is_active = self.mtg_info.get("is_active_sync", False)

        candidate_urls = []
        if NetworkState.get_instance().is_local_active():
            local_base = URLRouter._get_local_server_base(wg_name)
            candidate_urls.append(local_base)

        if is_active:
            sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
            candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

        if main_url:
            candidate_urls.append(main_url)

        if not candidate_urls:
            return QMessageBox.warning(self, "No URL Available",
                                       "Cannot download Chairman's Notes: No valid meeting URL found.")

        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setText("⏳ Downloading Notes...")

        chair_notes_dir = self.meeting_dir / "Agenda" / "ChairNotes"
        self.chair_notes_thread = ChairNotesDownloaderThread(candidate_urls, chair_notes_dir)
        self.chair_notes_thread.progress.connect(lambda msg: self.refresh_btn.setText(f"⏳ {msg}"[:30]))
        self.chair_notes_thread.finished.connect(self._on_chair_notes_finished)
        self.chair_notes_thread.start()

    def _on_chair_notes_finished(self, success: bool, count: int, files: list, msg: str):
        self.refresh_btn.setEnabled(True)
        self.refresh_btn.setText("🔄 Refresh")

        if success and count > 0:
            self.refresh_btn.setText(f"✅ {count} Notes")
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
            QMessageBox.information(self, "Chairman's Notes Downloaded", msg)
        elif success and count == 0:
            QMessageBox.information(self, "Chairman's Notes", msg)
        else:
            QMessageBox.warning(self, "Download Failed", msg)

    def _open_chair_notes_folder(self):
        """Opens the local Agenda/ChairNotes folder in the OS file explorer."""
        _open_folder(self.meeting_dir / "Agenda" / "ChairNotes")

    def _open_excel_export_dialog(self):
        visible_count = self.proxy.rowCount()
        total_count = self.model.rowCount()

        dialog = ExcelExportDialog(self, visible_count, total_count)
        if dialog.exec_() != QDialog.Accepted:
            return

        selected_columns = dialog.get_selected_columns()
        is_visible_only = dialog.is_visible_only()
        auto_open = dialog.should_auto_open()

        if is_visible_only:
            rows_to_export = []
            for r in range(self.proxy.rowCount()):
                source_idx = self.proxy.mapToSource(self.proxy.index(r, 0))
                rows_to_export.append(self.model._data[source_idx.row()])
        else:
            rows_to_export = list(self.model._data)

        if not rows_to_export:
            QMessageBox.warning(self, "No Data", "There are no rows available to export.")
            return

        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        scope_str = "Filtered" if is_visible_only else "All"
        default_filename = f"TDocs_{wg}_{mtg}_{scope_str}.xlsx"
        default_target = self.meeting_dir / "Export" / default_filename

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDocs Excel Export",
            str(default_target),
            "Excel Workbook (*.xlsx)"
        )

        if not save_path:
            return

        self.export_btn.setText("⏳ Exporting Excel...")
        self.export_btn.setEnabled(False)

        self.excel_exporter_thread = ExcelExporterThread(
            Path(save_path),
            rows_to_export,
            selected_columns,
            self.mtg_info,
            auto_open=auto_open
        )
        self.excel_exporter_thread.progress.connect(self._on_excel_export_progress)
        self.excel_exporter_thread.finished.connect(
            lambda success, msg, auto=auto_open: self._on_excel_export_finished(success, msg, auto)
        )
        self.excel_exporter_thread.start()

    def _on_excel_export_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:25])

    def _on_excel_export_finished(self, success: bool, msg: str, auto_open: bool):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)

        if success:
            QMessageBox.information(self, "Export Complete", f"Successfully exported TDocs to:\n{msg}")
            if auto_open:
                _open_folder(Path(msg))
        else:
            QMessageBox.warning(self, "Export Failed", f"Failed to export Excel:\n{msg}")

    # =========================================================================
    # 📧 GENERAL EMAIL INTEGRATION HANDLERS
    # =========================================================================
    def _refresh_email_counts(self):
        """Queries the SQLite database for total and unread email matches and refreshes the table."""
        try:
            counts = self.general_email_db.get_email_counts_per_tdoc()
            self.model.set_email_counts(counts)
        except Exception as e:
            logging.error(f"Error refreshing email counts: {e}")

    def _open_tdoc_emails(self, tdoc_id: str):
        """Opens the inspection card dialog modelessly as an independent top-level window."""
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()

        # Extract base document if a revision was clicked (e.g., S2-2608457r01 -> S2-2608457)
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_clean, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_clean

        # If already open, restore and bring it to the foreground
        if base_tdoc in self._email_dialogs:
            dlg = self._email_dialogs[base_tdoc]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return

        family = self.model.get_family_tdocs(base_tdoc)
        wg = self.mtg_info.get("wg_name", "SA2")

        dialog = TDocEmailsDialog(base_tdoc, family, self.general_email_db.db_path, wg=wg, parent=None)
        dialog.data_changed.connect(self._refresh_email_counts)

        # ---> Connect link clicks inside reading pane to open related emails for other TDocs
        dialog.tdoc_selected.connect(self._open_tdoc_emails)

        # Retain reference to prevent garbage collection
        self._email_dialogs[base_tdoc] = dialog
        dialog.finished.connect(lambda _, t=base_tdoc: self._email_dialogs.pop(t, None))

        dialog.show()

    def _on_configure_email_folders(self):
        """Opens the Outlook Folder Configuration Dialog for the current Working Group."""
        wg = self.mtg_info.get("wg_name", "SA2")
        dialog = GeneralEmailFoldersDialog(wg, self)
        dialog.exec_()

    def _on_sync_general_emails(self):
        """Opens the sync date buffer dialog and dispatches the background GeneralEmailSyncThread."""
        wg = self.mtg_info.get("wg_name", "SA2")
        folders = load_wg_email_config(wg)
        if not folders:
            if QMessageBox.question(self, "No Folders Configured",
                                    f"No Outlook folders configured for {wg}.\nConfigure folders now?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._on_configure_email_folders()
            return

        dialog = GeneralEmailSyncDialog(wg, self.mtg_info.get("start_date", ""), self.mtg_info.get("end_date", ""), self)
        if dialog.exec_() != QDialog.Accepted:
            return

        params = dialog.get_params()
        self.email_btn.setEnabled(False)
        self.email_btn.setText("⏳ Syncing...")

        self.general_email_sync_thread = GeneralEmailSyncThread(
            folders, self.general_email_db.db_path,
            start_date=params["start_date"], end_date=params["end_date"], days_buffer=params["buffer"]
        )
        self.general_email_sync_thread.progress_msg.connect(lambda m: self.email_btn.setText(f"⏳ {m[:20]}"))
        self.general_email_sync_thread.finished.connect(self._on_general_sync_finished)
        self.general_email_sync_thread.start()

    def _on_general_sync_finished(self, success: bool, msg: str, count: int):
        """Restores button state and prompts feedback once email syncing finishes."""
        self.email_btn.setEnabled(True)
        self.email_btn.setText("📧 Emails ▾")
        self._refresh_email_counts()
        if success:
            QMessageBox.information(self, "Email Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Email Sync Failed", msg)

    def _on_mark_all_emails_read(self):
        """Marks all general emails across the entire meeting database as read."""
        self.general_email_db.mark_all_read()
        self._refresh_email_counts()
        QMessageBox.information(self, "Mark All Read", "All emails have been marked as read.")

    def _on_wipe_generic_emails(self):
        """Fast wipe of the generic emails tables only, preserving SA2 eMeeting tables."""
        reply = QMessageBox.question(
            self,
            "Confirm Wipe Generic Emails",
            "Are you sure you want to completely wipe all generic email records for this meeting?\n\n"
            "• SA2 eMeeting tables will be preserved.\n"
            "• You can re-sync from Outlook anytime.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.general_email_db.wipe_generic_emails()
            self._refresh_email_counts()
            QMessageBox.information(self, "Database Wiped", "Generic email records have been wiped.")

    def _on_table_context_menu(self, pos: QPoint):
        """Displays row-level context actions organized into clean submenus."""
        index = self.table.indexAt(pos)
        if not index.isValid():
            return
        source_idx = self.proxy.mapToSource(index)
        row_data = self.model._data[source_idx.row()]
        tdoc_id = row_data.get("TDoc", "")
        if not tdoc_id:
            return

        revisions = self.model.revisions.get(tdoc_id, [])
        family = self.model.get_family_tdocs(tdoc_id)

        # Retrieve unread email count from table model cache if available
        unread_count = 0
        if hasattr(self.model, "_email_counts") and isinstance(
            self.model._email_counts, dict
        ):
            counts = self.model._email_counts.get(tdoc_id)
            if isinstance(counts, (tuple, list)) and len(counts) > 1:
                unread_count = counts[1]
            elif isinstance(counts, dict):
                unread_count = counts.get("unread", 0)

        build_row_context_menu(
            parent=self.table,
            row_data=row_data,
            revisions=revisions,
            family=family,
            docs_ftp_url=self.docs_ftp_url,
            revisions_url=self.revisions_url,
            meeting_dir=self.meeting_dir,
            download_callback=self._trigger_download_thread,
            export_llm_callback=self._export_llm_single,
            compose_email_callback=self._compose_email_draft,
            open_emails_callback=self._open_tdoc_emails,
            mark_read_callback=lambda: [
                self.general_email_db.set_tdocs_read_status(set(family), True),
                self._refresh_email_counts(),
            ],
            mark_unread_callback=lambda: [
                self.general_email_db.set_tdocs_read_status(
                    set(family), False
                ),
                self._refresh_email_counts(),
            ],
            open_details_callback=lambda: TDocInfoDialog(
                row_data,
                docs_ftp_url=self.docs_ftp_url,
                revisions=revisions,
                parent=self,
            ).exec_(),
            unread_emails_count=unread_count,
            pos=self.table.viewport().mapToGlobal(pos),
        )