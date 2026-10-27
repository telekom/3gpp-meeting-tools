# --- File: src/modules/meetings/ui/tdocs_window.py ---
import datetime
import json
import logging
import os
import re
import subprocess
import urllib.parse
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QPoint
from PyQt5.QtGui import QCursor, QKeySequence
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableView,
    QHeaderView, QLabel, QLineEdit, QFrame,
    QPushButton, QMessageBox, QMenu, QApplication,
    QToolTip, QCheckBox, QDialog, QTextEdit, QFileDialog
)

from core.network.network_state import NetworkState
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_WARNING
)
from modules.emails.ui.email_window import EmailManagerWindow
from modules.meetings.core.agenda_manager import AgendaManager, AgendaDownloaderThread
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.llm_exporter import (
    TDocMarkdownExporterThread,
    LLMExporterThread,
)
from modules.meetings.core.markdown_exporter import MarkdownExporterThread
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.stats.exporter_thread import StatisticsExporterThread
from modules.meetings.core.tdocs_db import TDocsDatabase
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_package_exporter import TDocsPackageExporterThread
from modules.meetings.core.meeting_wi_enricher import MeetingWIEnricherThread
from modules.meetings.core.tdocs_threads import (
    TDocsRevisionsFetcherThread,
    TDocActionThread,
    TdocsByAgendaThread,
    WordAgendaImporterThread,
)
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdoc_delegates import HtmlDelegate, TDocActionDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox, ComparisonCartWidget
from modules.meetings.ui.tdocs_dialogs import (
    ReadOnlyViewerDialog,
    InteractiveNotesDialog,
    StatisticsSettingsDialog, ExcelExportDialog,
    TDocInfoDialog, ALL_AVAILABLE_COLUMNS
)
from modules.meetings.ui.tdocs_menus import build_action_menu, build_related_menu, build_row_context_menu
from modules.meetings.ui.tdocs_menu import handle_table_context_menu
from modules.meetings.ui.tdocs_export_controller import (
    get_package_export_rows,
    export_tdocs_package,
    open_excel_export_dialog,
    export_markdown_visible,
    export_markdown_single,
    open_post_meeting_export_dialog,
)
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel, natural_sort_key
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.ui.general_email_dialog import (
    GeneralEmailFoldersDialog, GeneralEmailSyncDialog, TDocEmailsDialog, load_wg_email_config
)
from modules.meetings.ui.tdoc_triage_dialog import TDocTriageDialog
from core.utils.paths import get_project_root
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.work_items.core.wi_database import WorkItemsDatabase


from core.utils.folder_opener import open_folder as _open_folder
from core.utils.terminal_opener import open_terminal as _open_terminal
from modules.meetings.ui.tdocs_drop_overlay import DropOverlayWidget
from modules.meetings.ui.tdocs_agenda_controller import (
    import_word_agenda_dialog,
    process_imported_agenda_files,
    process_next_in_import_queue,
    on_word_import_progress,
    on_word_import_finished,
    finalize_import_batch,
    download_agenda_csv,
    on_agenda_download_finished,
    download_chair_notes,
    on_chair_notes_finished,
    open_chair_notes_folder,
)


class TDocsWindow(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}
        self._email_dialogs = {}
        self._triage_dialogs = {}

        self.agenda_dir = Path(filepath).parent
        self.agenda_map = AgendaManager.load_local_agenda(self.agenda_dir)
        self.agenda_dl_thread = None

        self.general_email_db = GeneralEmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.general_email_sync_thread = None

        self._import_queue = []
        self._is_importing_agenda = False
        self._total_import_batch_count = 0
        self._import_accumulated_merged = 0
        self._import_success_files = []
        self._import_failed_files = []

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

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        self._setup_header(main_layout, title, is_electronic, len(tdocs_data))
        self._setup_filters(main_layout)
        self._setup_table(main_layout, tdocs_data, user_data)
        self._setup_cache()

        self.drop_overlay = DropOverlayWidget(self)
        self._start_background_wi_enrichment()

    def showEvent(self, event):
        super().showEvent(event)
        self._load_tracked_wi_tokens()
        self._update_filter_facets()

    def _setup_header(self, layout, title, is_electronic, count):
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 16px; color: #1E293B;")
        title_lbl.setToolTip("Electronic Meeting (eMeeting)" if is_electronic else "In-Person Meeting (Face-to-Face)")
        title_lbl.setCursor(Qt.WhatsThisCursor)

        self.routing_indicator = QLabel("⚪ Routing...")
        self.routing_indicator.setStyleSheet("font-weight: bold; padding: 2px 6px; border-radius: 4px;")

        self.routing_timer = QTimer(self)
        self.routing_timer.timeout.connect(self._update_routing_indicator)
        self.routing_timer.start(2000)
        self._update_routing_indicator()

        self.last_mod_lbl = QLabel(self._get_mod_date_str())
        self.last_mod_lbl.setStyleSheet("font-size: 11px; color: #64748B; margin-right: 8px; font-style: italic;")

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.setToolTip("Reload TDocs or fetch the latest revisions from the FTP.")

        refresh_menu = QMenu(self)
        refresh_menu.setToolTipsVisible(True)
        act = refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        act.setToolTip("Download and reload the latest official TDoc Excel list for this meeting.")
        act = refresh_menu.addAction("📋 Download Agenda CSV", self._download_agenda_csv)
        act.setToolTip("Download agenda.csv and refresh Agenda Item names/acronyms used by the TDoc view.")
        if self.is_sa2:
            act = refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)
            act.setToolTip("Fetch TdocsByAgenda.htm and merge chairman/secretary agenda metadata into the TDoc list.")
            act = refresh_menu.addAction("📝 Import Word Document (.docx / .doc)...", self._import_word_agenda_dialog)
            act.setToolTip("Import one or more local Word/HTML agenda or chairman-notes files and merge their TDoc metadata.")
            act = refresh_menu.addAction("📥 Download Chairman's Notes", self._download_chair_notes)
            act.setToolTip("Download available chairman's notes for this meeting into the local Agenda cache.")
        if self.is_sa2_electronic:
            act = refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            act.setToolTip("Refresh the revision list for this electronic meeting without reloading the Excel list.")
            act = refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
            act.setToolTip("Refresh both the official TDoc Excel list and the available revision metadata.")
        self.refresh_btn.setMenu(refresh_menu)

        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.folder_btn.setToolTip("Access local cache folders, on-site server pages, export reports, and remote FTP directories.")

        self.folder_menu = QMenu(self)
        self.folder_menu.setToolTipsVisible(True)
        self.folder_menu.aboutToShow.connect(self._populate_resources_menu)
        self.folder_btn.setMenu(self.folder_menu)

        self.excel_btn = QPushButton("📗 Original Excel")
        self.excel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excel_btn.setToolTip("Open the raw, underlying 3GPP Excel TDocs list.")
        self.excel_btn.clicked.connect(self._open_excel)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.export_btn.setToolTip(
            "Export meeting TDoc metadata, document packages, LLM-ready content, or Markdown reports."
        )

        export_menu = QMenu(self)
        export_menu.setToolTipsVisible(True)

        act_export_excel = export_menu.addAction(
            "📊 Export to Excel (.xlsx)...", self._open_excel_export_dialog
        )
        act_export_excel.setToolTip(
            "Export TDoc metadata to a formatted Excel workbook. "
            "You can choose visible/all rows, columns, their order, and the destination file."
        )
        act_export_excel.setStatusTip(act_export_excel.toolTip())

        act_export_package = export_menu.addAction(
            "📦 Export TDocs + Excel...", self._export_tdocs_package
        )
        act_export_package.setToolTip(
            "Create a self-contained package under this meeting's Export folder. "
            "Exports selected TDoc rows, or all currently visible rows when nothing is selected, "
            "using cached documents where available and including an Excel manifest."
        )
        act_export_package.setStatusTip(act_export_package.toolTip())

        act_export_md = export_menu.addAction(
            "📄 Export TDocs to Markdown (.md)", self._export_markdown_visible
        )
        act_export_md.setToolTip(
            "Convert and export all currently visible/filtered TDocs to individual Markdown (.md) documents with an index manifest."
        )
        act_export_md.setStatusTip(act_export_md.toolTip())

        act_export_reports = export_menu.addAction(
            "📝 Export Markdown Summary Reports", self._export_reports
        )
        act_export_reports.setToolTip(
            "Generate Markdown summary reports for the meeting from the currently loaded TDoc data."
        )
        act_export_reports.setStatusTip(act_export_reports.toolTip())

        act_export_post_meeting = export_menu.addAction(
            "📑 Export Post-Meeting Report (.md)...", self._open_post_meeting_export_dialog
        )
        act_export_post_meeting.setToolTip(
            "Generate a structured, per-release and per-WI post-meeting Markdown dossier covering incoming LS and agreed pCR/CR/LS."
        )
        act_export_post_meeting.setStatusTip(act_export_post_meeting.toolTip())

        self.export_menu = export_menu
        self.export_menu_sep = export_menu.addSeparator()
        self.act_open_export_dir = export_menu.addAction(
            "📂 Open Export Folder", self._open_export_folder
        )
        self.act_open_export_dir.setToolTip(
            "Open this meeting's local Export directory in Windows Explorer (Ctrl+Shift+E)."
        )
        self.act_open_export_dir.setStatusTip(self.act_open_export_dir.toolTip())
        self.act_open_export_dir.setShortcut(QKeySequence("Ctrl+Shift+E"))
        self.addAction(self.act_open_export_dir)

        self.export_menu.aboutToShow.connect(self._on_export_menu_about_to_show)
        self.export_btn.setMenu(export_menu)

        self.draft_cr_btn = QPushButton("📝 Draft CR / pCR")
        self.draft_cr_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.draft_cr_btn.setToolTip("Compose a Change Request (CR) or Pseudo-CR (pCR) contribution for this meeting.")
        self.draft_cr_btn.clicked.connect(lambda: self._open_draft_cr_dialog())

        self.stats_btn = QPushButton("📊 Statistics")
        self.stats_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.stats_btn.setToolTip("Generate an interactive HTML statistics dashboard for this meeting.")
        self.stats_btn.clicked.connect(self._generate_statistics)

        self.stats_cfg_btn = QPushButton("⚙️")
        self.stats_cfg_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.stats_cfg_btn.setFixedWidth(34)
        self.stats_cfg_btn.setToolTip("Configure Statistics Parameters")
        self.stats_cfg_btn.clicked.connect(self._open_stats_config)

        self.email_btn = QPushButton("📧 Emails ▾")
        self.email_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.email_btn.setToolTip("Sync and inspect Outlook emails linked to TDocs and revisions.")

        email_menu = QMenu(self)
        email_menu.setToolTipsVisible(True)
        act = email_menu.addAction("🔄 Sync Related Emails...", self._on_sync_general_emails)
        act.setToolTip("Scan the configured Outlook folders for emails that reference TDocs from this meeting.")
        act = email_menu.addAction("⚙️ Configure Outlook Folders...", self._on_configure_email_folders)
        act.setToolTip("Choose which Outlook folders are searched when syncing TDoc-related emails.")
        act = email_menu.addAction("✔️ Mark All Emails as Read", self._on_mark_all_emails_read)
        act.setToolTip("Mark all locally indexed meeting emails as read.")
        email_menu.addSeparator()
        act = email_menu.addAction("🗑️ Wipe Generic Emails Database...", self._on_wipe_generic_emails)
        act.setToolTip("Delete locally indexed generic meeting emails. They can be rebuilt by syncing Outlook again.")
        if self.is_sa2_electronic:
            email_menu.addSeparator()
            act = email_menu.addAction("📊 Open eMeeting Email Manager (Dashboard)", self._open_email_manager)
            act.setToolTip("Open the electronic-meeting email dashboard for detailed email/TDoc inspection.")
        self.email_btn.setMenu(email_menu)

        self.count_lbl = QLabel(f"Showing {count} of {count} TDocs")
        self.count_lbl.setStyleSheet("font-size: 12px; color: #64748B;")

        self.instant_fetch_input = QLineEdit()
        self.instant_fetch_input.setPlaceholderText("Instant Fetch...")
        self.instant_fetch_input.setToolTip("Instantly fetch a TDoc, bypassing the table entirely. Press Enter to launch.")
        self.instant_fetch_input.setText(self._get_tdoc_prefix())
        self.instant_fetch_input.setFixedWidth(120)
        self.instant_fetch_input.returnPressed.connect(self._on_instant_fetch)
        self.instant_fetch_input.setFocus()
        self.instant_fetch_input.setCursorPosition(len(self.instant_fetch_input.text()))

        self.instant_cmp_btn = QPushButton("⚖️ vs Word")
        self.instant_cmp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.instant_cmp_btn.setCursor(Qt.PointingHandCursor)
        self.instant_cmp_btn.setToolTip(
            "Download the TDoc entered in Instant Fetch and immediately compare it with your currently active Microsoft Word document."
        )
        self.instant_cmp_btn.clicked.connect(self._on_instant_compare_active_word)

        header_layout.addWidget(title_lbl)
        header_layout.addWidget(self.routing_indicator)
        header_layout.addStretch()

        header_layout.addWidget(QLabel("🚀"))
        header_layout.addWidget(self.instant_fetch_input)
        header_layout.addWidget(self.instant_cmp_btn)
        header_layout.addSpacing(10)

        header_layout.addWidget(self.last_mod_lbl)
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addWidget(self.export_btn)
        header_layout.addWidget(self.draft_cr_btn)
        header_layout.addWidget(self.stats_btn)
        header_layout.addWidget(self.stats_cfg_btn)
        header_layout.addWidget(self.email_btn)
        header_layout.addSpacing(10)
        header_layout.addWidget(self.count_lbl)
        layout.addLayout(header_layout)

    def _populate_resources_menu(self):
        self.folder_menu.clear()
        self.folder_menu.addAction("📁 Local: Meeting Folder", self._open_meeting_folder)
        if self.meeting_dir.is_dir():
            act = self.folder_menu.addAction(
                "💻 Open Terminal in Meeting Cache",
                lambda: _open_terminal(self.meeting_dir, self),
            )
            act.setToolTip("Open Windows Terminal with this meeting's local cache folder as its working directory.")
        export_dir = self.meeting_dir / "Export"
        if export_dir.is_dir():
            act_exp = self.folder_menu.addAction("📂 Local: Export Folder", self._open_export_folder)
            act_exp.setToolTip("Open this meeting's local Export directory containing packages and Markdown dossiers.")
        local_agenda_csv = self.agenda_dir / "agenda.csv"
        if local_agenda_csv.exists():
            self.folder_menu.addAction("📋 Local: agenda.csv", lambda: _open_folder(local_agenda_csv))
        if self.is_sa2:
            self.folder_menu.addAction("📄 Local: TdocsByAgenda.htm", self._open_agenda_file)

        self.folder_menu.addSeparator()
        self.folder_menu.addAction("⚠️ View Unmatched Companies", self._show_unmatched_companies)
        act_resource_md_export = self.folder_menu.addAction("📝 Export Markdown Reports", self._export_reports)
        act_resource_md_export.setToolTip(
            "Generate Markdown summary reports for the meeting from the currently loaded TDoc data."
        )
        act_resource_md_export.setStatusTip(act_resource_md_export.toolTip())

        if NetworkState.get_instance().is_local_active():
            self.folder_menu.addSeparator()
            wg_name = self.mtg_info.get("wg_name", "").upper()
            local_base = URLRouter._get_local_server_base(wg_name)
            self.folder_menu.addAction("🟢 Local Server: Main WG Folder", lambda u=local_base: webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Docs Folder", lambda u=f"{local_base}/Docs": webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Inbox Folder", lambda u=f"{local_base}/Inbox": webbrowser.open(u))
            if self.is_sa2:
                self.folder_menu.addAction("🟢 Local Server: Revisions Folder", lambda u=f"{local_base}/Inbox/Revisions": webbrowser.open(u))
                self.folder_menu.addAction("🟢 Local Server: TdocsByAgenda.htm", lambda u=f"{local_base}/TdocsByAgenda.htm": webbrowser.open(u))
            self.folder_menu.addAction("🟢 Local Server: Home (10.10.10.10)", lambda: webbrowser.open("http://10.10.10.10/"))

        self.folder_menu.addSeparator()
        if self.main_ftp_url:
            self.folder_menu.addAction("🌐 Web FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
        if self.docs_ftp_url:
            self.folder_menu.addAction("🌐 Web FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
        if self.revisions_url:
            self.folder_menu.addAction("🌐 Web FTP: Revisions Folder", lambda: webbrowser.open(self.revisions_url))

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, legacy_url: str = None,
                                 is_silent_compare: bool = False, auto_compare_active_word: bool = False):
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
        thread.auto_compare_active_word = auto_compare_active_word
        thread.target_filename = target_filename
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))

        self.active_threads[base_tdoc] = thread
        thread.start()

    def _compare_tdoc_with_active_word(self, base_tdoc: str, target_filename: str):
        """
        1-click action to compare the currently active Microsoft Word document (Slot A)
        against the specified TDoc or revision (Slot B).
        """
        mgr = ComparisonManager.get_instance()
        active_doc = mgr.stage_active_word_document(slot="A")
        if not active_doc:
            QMessageBox.warning(
                self,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and the document you wish to compare is open on your screen.",
            )
            return

        self._trigger_download_thread(
            base_tdoc,
            target_filename,
            is_silent_compare=True,
            auto_compare_active_word=True,
        )

    def _setup_filters(self, layout):
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self._apply_search_filter)

        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
            QLabel {
                font-weight: bold;
                color: #475569;
                border: none;
            }
            QLineEdit, QComboBox {
                padding: 4px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background: #FFFFFF;
            }
            QComboBox[activeFilter="true"] {
                background-color: #EFF6FF;
                border: 1.5px solid #3B82F6;
                color: #1D4ED8;
                font-weight: bold;
            }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(8, 5, 8, 5)
        filter_layout.setSpacing(6)

        search_lbl = QLabel("🔍")
        search_lbl.setToolTip("Full-text search across TDoc number, title, source, remarks, and related metadata.")
        filter_layout.addWidget(search_lbl)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search TDocs (e.g. "CR", -draft)...')
        self.search_input.setToolTip(
            'Search across TDoc metadata.\n'
            '• Exact phrases: "draft baseline"\n'
            '• Exclude words: -draft or !draft\n'
            '• Substrings & numbers: S2-260855, TS 23.501'
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(160)
        self.search_input.textChanged.connect(lambda _: self.search_timer.start())
        filter_layout.addWidget(self.search_input, stretch=2)

        self.btn_followed_wis = QPushButton("🎯")
        self.btn_followed_wis.setCheckable(True)
        self.btn_followed_wis.setCursor(Qt.PointingHandCursor)
        self.btn_followed_wis.setFixedSize(30, 28)
        self.btn_followed_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_followed_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_followed_wis.setToolTip("Filter to TDocs associated with Work Items you are following (assigned responsibility).\nRight-click to toggle including unassigned Incoming LS.")
        self.btn_followed_wis.clicked.connect(self._toggle_followed_wis_filter)
        filter_layout.addWidget(self.btn_followed_wis)

        self.btn_starred_wis = QPushButton("⭐")
        self.btn_starred_wis.setCheckable(True)
        self.btn_starred_wis.setCursor(Qt.PointingHandCursor)
        self.btn_starred_wis.setFixedSize(30, 28)
        self.btn_starred_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_starred_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_starred_wis.setToolTip("Filter to TDocs associated with Work Items you have starred (personal watchlist).\nRight-click to toggle including unassigned Incoming LS.")
        self.btn_starred_wis.clicked.connect(self._toggle_starred_wis_filter)
        filter_layout.addWidget(self.btn_starred_wis)
        self._update_tracked_wi_button_styles()

        def _make_vsep():
            sep = QFrame()
            sep.setFrameShape(QFrame.VLine)
            sep.setFrameShadow(QFrame.Plain)
            sep.setStyleSheet("color: #CBD5E1; background-color: #E2E8F0; width: 1px; margin: 2px 1px;")
            return sep

        filter_layout.addWidget(_make_vsep())

        self.company_combo = CheckableComboBox("Company")
        self.company_combo.setMaximumWidth(160)
        self.company_combo.setToolTip('Filter by contributing company. Grey entries currently have no intersection with the other active filters but remain selectable.')
        self.company_combo.selectionChanged.connect(self._on_company_changed)
        filter_layout.addWidget(self.company_combo)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMaximumWidth(110)
        self.type_combo.setToolTip('Filter by TDoc type. Grey entries currently have no intersection with the other active filters but remain selectable.')
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMaximumWidth(110)
        self.ai_combo.setToolTip('Filter by Agenda Item. When agenda.csv is available, associated acronyms are shown in parentheses.')
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.wi_combo = CheckableComboBox("WI")
        self.wi_combo.setMaximumWidth(120)
        self.wi_combo.setToolTip('Filter by the Related WIs metadata from the official TDoc list. Grey entries currently have no intersection with the other active filters.')
        self.wi_combo.selectionChanged.connect(self._on_wi_changed)
        filter_layout.addWidget(self.wi_combo)

        self.status_combo = CheckableComboBox("TDoc Status")
        self.status_combo.setMaximumWidth(125)
        self.status_combo.setToolTip('Filter by the meeting/secretary TDoc status.')
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        self.my_status_combo = CheckableComboBox("My Status")
        self.my_status_combo.setMaximumWidth(120)
        self.my_status_combo.setToolTip('Filter by your personal TDoc status.')
        self.my_status_combo.selectionChanged.connect(self._on_my_status_changed)
        filter_layout.addWidget(self.my_status_combo)

        filter_layout.addWidget(_make_vsep())

        if self.is_sa2:
            self.chk_no_comments = QCheckBox("No Comments")
            self.chk_no_comments.setToolTip('Show only TDocs without secretary/chair comments.')
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            filter_layout.addWidget(self.chk_no_comments)

        self.btn_reset_filters = QPushButton("✕ Clear")
        self.btn_reset_filters.setToolTip("Reset all search text and active filters to show all TDocs.")
        self.btn_reset_filters.setCursor(Qt.PointingHandCursor)
        self.btn_reset_filters.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #64748B;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                border-color: #FCA5A5;
                color: #DC2626;
            }
        """)
        self.btn_reset_filters.clicked.connect(self._clear_all_filters)
        filter_layout.addWidget(self.btn_reset_filters)

        layout.addWidget(filter_frame)

    def _setup_table(self, layout, data, user_data):
        self.settings = MeetingsSettings()
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, data, user_data, agenda_map=self.agenda_map)
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setIncludeUnassignedLsIn(self.settings.get_include_unassigned_ls_in())
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.doubleClicked.connect(self._show_cell_popup)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet("""
            QTableView {
                gridline-color: #F1F5F9;
                border: 1px solid #CBD5E1;
                background-color: #FFFFFF;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                padding: 4px;
                font-weight: bold;
                border: 1px solid #E2E8F0;
                color: #1E293B;
            }
        """)
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

        if len(self.model._headers) > 13:
            header.resizeSection(13, 85)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)

        layout.addWidget(self.table)

        # Bottom Status & Shared Comparison Cart Bar
        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(2, 2, 2, 0)
        bottom_bar.setSpacing(8)
        bottom_bar.addStretch()

        self.cart_frame = ComparisonCartWidget(self)
        self.btn_cart_launcher = self.cart_frame.create_launcher_button(self)
        bottom_bar.addWidget(self.btn_cart_launcher)
        bottom_bar.addWidget(self.cart_frame)
        layout.addLayout(bottom_bar)

        self._load_tracked_wi_tokens()
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
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_wi_changed(self, wis):
        self.proxy.setWIFilters(wis)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_my_status_changed(self, statuses):
        self.proxy.setMyStatusFilters(statuses)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_company_changed(self, companies):
        self.proxy.setCompanyFilters(companies)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_followed_wis_filter(self):
        self._load_tracked_wi_tokens()
        self.proxy.setTrackedWIFilters(
            self.btn_followed_wis.isChecked(),
            self.btn_starred_wis.isChecked()
        )
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_starred_wis_filter(self):
        self._load_tracked_wi_tokens()
        self.proxy.setTrackedWIFilters(
            self.btn_followed_wis.isChecked(),
            self.btn_starred_wis.isChecked()
        )
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _show_tracked_wi_context_menu(self, pos):
        sender = self.sender()
        menu = QMenu(self)
        act = menu.addAction("Include unassigned Incoming LS (LS in)")
        act.setCheckable(True)
        act.setChecked(getattr(self.proxy, "include_unassigned_ls_in", True))
        act.toggled.connect(self._on_toggle_include_unassigned_ls_in)
        if sender and hasattr(sender, "mapToGlobal"):
            menu.exec_(sender.mapToGlobal(pos))
        else:
            from PyQt5.QtGui import QCursor
            menu.exec_(QCursor.pos())

    def _on_toggle_include_unassigned_ls_in(self, checked: bool):
        if not hasattr(self, "settings"):
            self.settings = MeetingsSettings()
        self.settings.set_include_unassigned_ls_in(checked)
        self.proxy.setIncludeUnassignedLsIn(checked)
        self._update_facet_counts()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _update_tracked_wi_button_styles(self):
        """Update active/inactive visual styles of Followed and Starred WI quick-filter chips."""
        if not hasattr(self, "btn_followed_wis") or not hasattr(self, "btn_starred_wis"):
            return

        if self.btn_followed_wis.isChecked():
            self.btn_followed_wis.setStyleSheet("""
                QPushButton {
                    background-color: #DBEAFE;
                    color: #1E40AF;
                    border: 1.5px solid #3B82F6;
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #BFDBFE;
                    border-color: #2563EB;
                }
            """)
        else:
            self.btn_followed_wis.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC;
                    color: #475569;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    border-color: #94A3B8;
                }
            """)

        if self.btn_starred_wis.isChecked():
            self.btn_starred_wis.setStyleSheet("""
                QPushButton {
                    background-color: #FEF3C7;
                    color: #92400E;
                    border: 1.5px solid #F59E0B;
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #FDE68A;
                    border-color: #D97706;
                }
            """)
        else:
            self.btn_starred_wis.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC;
                    color: #475569;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    border-color: #94A3B8;
                }
            """)

    def _load_tracked_wi_tokens(self):
        """Loads followed/starred WI codes from settings, resolves tokens from 3gpp_data.db, and updates proxy."""
        if not hasattr(self, "proxy"):
            return
        try:
            settings = WorkItemsSettings()
            starred_codes, following_codes = settings.get_tracking_sets()
            db_path = get_project_root() / "db" / "3gpp_data.db"
            if db_path.exists():
                wi_db = WorkItemsDatabase(db_path)
                followed_tokens = wi_db.get_tracked_wi_tokens(following_codes)
                starred_tokens = wi_db.get_tracked_wi_tokens(starred_codes)
            else:
                followed_tokens = set(following_codes)
                starred_tokens = set(starred_codes)
            self.proxy.setTrackedWITokens(followed_tokens, starred_tokens)
        except Exception as e:
            logging.error(f"Failed to load tracked WI tokens: {e}", exc_info=True)
            self.proxy.setTrackedWITokens(set(), set())

    def _on_no_comments_toggled(self, checked):
        self.proxy.setNoCommentsFilter(checked)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _update_count_label(self):
        self.count_lbl.setText(f"Showing {self.proxy.rowCount()} of {self.model.rowCount()} TDocs")

    def _update_filter_facets(self):
        """Update zero-intersection styling for every metadata filter."""
        if not hasattr(self, "proxy") or not hasattr(self, "model"):
            return

        facets = {
            "company": set(),
            "type": set(),
            "ai": set(),
            "wi": set(),
            "status": set(),
            "my_status": set(),
        }

        followed_count = 0
        starred_count = 0
        has_followed_tokens = bool(self.proxy.followed_wi_tokens)
        has_starred_tokens = bool(self.proxy.starred_wi_tokens)

        # Each facet is evaluated with all active constraints except itself.
        # This answers: "Would selecting this option yield rows given my other filters?"
        for row_idx, row in enumerate(self.model._data):
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="company"):
                facets["company"].update(row.get("_Sanitized_Companies", ["Other"]))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="type"):
                facets["type"].add(str(row.get("Type", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="ai"):
                facets["ai"].add(str(row.get("Agenda Item", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="wi"):
                facets["wi"].update(self.model.get_related_wis(row_idx))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="status"):
                facets["status"].add(str(row.get("TDoc Status", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="my_status"):
                facets["my_status"].add(str(row.get("My Status", "") or "").strip())

            if has_followed_tokens or has_starred_tokens:
                if self.proxy.rowMatchesFilters(row_idx, exclude_filter="tracked_wi"):
                    row_tokens = row.get("_wi_tokens")
                    if row_tokens is None:
                        row_tokens = self.model._extract_row_wi_tokens(row)
                        row["_wi_tokens"] = row_tokens

                    doc_type = str(row.get("Type", "") or "").strip().lower()
                    raw_wis = self.model._split_related_wis(row.get("Related WIs", ""))
                    has_assigned_wi = any(w and w.lower() not in {"(no wi)", "none", "nan", "-"} for w in raw_wis)
                    is_unassigned_ls = ("ls in" in doc_type or doc_type == "ls") and not has_assigned_wi

                    wi_followed = bool(has_followed_tokens and (row_tokens & self.proxy.followed_wi_tokens))
                    wi_starred = bool(has_starred_tokens and (row_tokens & self.proxy.starred_wi_tokens))
                    ls_included = bool(self.proxy.include_unassigned_ls_in and is_unassigned_ls)

                    if wi_followed or (has_followed_tokens and ls_included):
                        followed_count += 1
                    if wi_starred or (has_starred_tokens and ls_included):
                        starred_count += 1

        self.company_combo.setAvailability(facets["company"])
        self.type_combo.setAvailability(facets["type"])
        self.ai_combo.setAvailability(facets["ai"])
        self.wi_combo.setAvailability(facets["wi"])
        self.status_combo.setAvailability(facets["status"])
        self.my_status_combo.setAvailability(facets["my_status"])

        ls_suffix = " (incl. unassigned LS in)" if getattr(self.proxy, "include_unassigned_ls_in", True) else ""

        if hasattr(self, "btn_followed_wis"):
            self.btn_followed_wis.setText("🎯")
            self.btn_followed_wis.setToolTip(
                f"🎯 Followed Work Items ({followed_count} matching{ls_suffix})\n"
                "Filter to TDocs associated with Work Items you are following (assigned responsibility).\n"
                "Right-click to toggle including unassigned Incoming LS."
            )
        if hasattr(self, "btn_starred_wis"):
            self.btn_starred_wis.setText("⭐")
            self.btn_starred_wis.setToolTip(
                f"⭐ Starred Work Items ({starred_count} matching{ls_suffix})\n"
                "Filter to TDocs associated with Work Items you have starred (personal watchlist).\n"
                "Right-click to toggle including unassigned Incoming LS."
            )

    def _refresh_comboboxes(self):
        def sanitize(val): return str(val).strip() if val is not None else ""

        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in self.model._data)))
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in self.model._data)),
                            key=natural_sort_key)
        unique_wis = sorted({
            wi
            for row_idx in range(self.model.rowCount())
            for wi in self.model.get_related_wis(row_idx)
        }, key=lambda x: x.lower())
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))
        unique_my_statuses = sorted(list(set(sanitize(r.get("My Status", "")) for r in self.model._data)))

        unique_companies = set()
        for r in self.model._data:
            unique_companies.update(r.get("_Sanitized_Companies", ["Other"]))
        sorted_companies = sorted(list(unique_companies), key=lambda x: x.lower())

        ai_display_labels = {}
        for ai_num in unique_ais:
            agenda_item = self.agenda_map.get(ai_num)
            if agenda_item:
                ai_display_labels[ai_num] = agenda_item.display_label
            else:
                excel_desc = next((str(r.get("Agenda item description", "")).strip()
                                   for r in self.model._data
                                   if str(r.get("Agenda Item", "")).strip() == ai_num and r.get("Agenda item description")), "")
                if excel_desc:
                    clean = re.sub(r'\s+', ' ', excel_desc)
                    max_desc_len = max(15, 38 - len(ai_num))
                    label_desc = clean if len(clean) <= max_desc_len else clean[:max_desc_len - 3].rstrip() + "..."
                    ai_display_labels[ai_num] = f"{ai_num} - {label_desc}"
                else:
                    ai_display_labels[ai_num] = ai_num

        self.type_combo.updateItems(unique_types)
        self.ai_combo.updateItems(unique_ais, display_labels=ai_display_labels)
        self._update_ai_tooltips()
        self.wi_combo.updateItems(unique_wis)
        self.status_combo.updateItems(unique_statuses)
        self.company_combo.updateItems(sorted_companies)
        self.my_status_combo.updateItems(unique_my_statuses)
        self._load_tracked_wi_tokens()
        self._update_tracked_wi_button_styles()

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setWIFilters(self.wi_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())
        self._update_filter_facets()

        QTimer.singleShot(0, self._update_count_label)

    def _update_ai_tooltips(self):
        if not hasattr(self, 'ai_combo') or not self.ai_combo.model():
            return
        model = self.ai_combo.model()
        for i in range(1, model.rowCount()):
            item = model.item(i)
            if not item:
                continue
            ai_num = str(item.data(Qt.UserRole) or "").strip()
            agenda_item = self.agenda_map.get(ai_num)
            if agenda_item:
                if hasattr(agenda_item, 'full_tooltip'):
                    full_desc = agenda_item.full_tooltip
                elif hasattr(agenda_item, 'description'):
                    full_desc = f"AI {ai_num}: {agenda_item.description}"
                else:
                    full_desc = f"AI {ai_num}: {agenda_item}"
                item.setData(full_desc, Qt.ToolTipRole)
            else:
                excel_desc = next((str(r.get("Agenda item description", "")).strip()
                                   for r in self.model._data
                                   if str(r.get("Agenda Item", "")).strip() == ai_num and r.get("Agenda item description")), "")
                if excel_desc:
                    item.setData(f"AI {ai_num}: {excel_desc}", Qt.ToolTipRole)
                else:
                    item.setData(f"Agenda Item {ai_num}", Qt.ToolTipRole)

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

        if hasattr(self, "btn_followed_wis"):
            self.btn_followed_wis.blockSignals(True)
            self.btn_followed_wis.setChecked(False)
            self.btn_followed_wis.blockSignals(False)

        if hasattr(self, "btn_starred_wis"):
            self.btn_starred_wis.blockSignals(True)
            self.btn_starred_wis.setChecked(False)
            self.btn_starred_wis.blockSignals(False)

        self.proxy.setTrackedWIFilters(False, False)
        self._update_tracked_wi_button_styles()

        for combo in [self.company_combo, self.type_combo, self.ai_combo, self.wi_combo, self.status_combo, self.my_status_combo]:
            combo.blockSignals(True)
            combo.model().item(0).setCheckState(Qt.Checked)
            for i in range(1, combo.model().rowCount()):
                combo.model().item(i).setCheckState(Qt.Checked)
            combo.updateText()
            combo.blockSignals(False)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setWIFilters(self.wi_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())
        self._update_filter_facets()

        QTimer.singleShot(0, self._update_count_label)

    def _handle_tdoc_action(self, base_tdoc: str):
        if base_tdoc in self.model.loading_tdocs or not self.docs_ftp_url:
            return
        revisions = self.model.revisions.get(base_tdoc, [])
        build_action_menu(
            self.table,
            base_tdoc,
            self.docs_ftp_url,
            self.revisions_url,
            revisions,
            self.meeting_dir,
            self._trigger_download_thread,
            self._export_llm_single,
            self._compose_email_draft,
            QCursor.pos(),
            compare_active_word_callback=self._compare_tdoc_with_active_word,
        )

    def _show_related_menu(self, target_tdoc: str, pos: QPoint):
        build_related_menu(
            self,
            target_tdoc,
            self.model.valid_tdocs,
            self.docs_ftp_url,
            self.revisions_url,
            self._scroll_to_tdoc,
            self._trigger_download_thread,
            self._export_llm_single,
            self.global_action_requested.emit,
            self._compose_email_draft,
            pos,
            triage_callback=self._open_tdoc_triage,
            compare_active_word_callback=self._compare_tdoc_with_active_word,
        )

    def _compose_email_draft(self, tdoc_id: str):
        row_data = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == tdoc_id.upper()),
                        None)

        if not row_data:
            match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
            row_data = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc), {})

        wg = str(self.mtg_info.get("wg_name", "")).strip().upper()
        mtg_num = str(self.mtg_info.get("meeting_number", "")).strip()
        ai = str(row_data.get("Agenda Item", "")).strip()
        title = str(row_data.get("Title", "")).strip()

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
            logging.info(f"📧 [Email Draft] Opened draft with subject: '{subject}'")
        except Exception as e:
            logging.warning(f"Failed to launch email client automatically: {e}")
            QApplication.clipboard().setText(subject)
            QToolTip.showText(QCursor.pos(), "📋 Subject copied to clipboard!", self)

    def _show_cell_popup(self, index):
        if not index.isValid():
            return
        col_name = self.model._headers[index.column()]

        if col_name == "Emails":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            self._open_tdoc_emails(row_data.get("TDoc", ""))
            return

        if col_name == "TDoc":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            tdoc_id = str(row_data.get("TDoc", "")).strip()
            revs = self.model.revisions.get(tdoc_id, [])
            TDocInfoDialog(
                row_data,
                docs_ftp_url=self.docs_ftp_url,
                revisions=revs,
                parent=self,
                triage_callback=self._open_tdoc_triage,  # <--- Added
            ).exec_()
            return

        if col_name not in [
            "Secretary Remarks", "Title", "Source", "Abstract", "My Notes", "My Status"
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
            ReadOnlyViewerDialog(self, f"📄 Viewing: {col_name} ({tdoc_id})", val).exec_()
        else:
            InteractiveNotesDialog(self, tdoc_id, row_data, self._save_user_data).exec_()

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

        if not success:
            QMessageBox.warning(self, "Export Failed", msg)
            return

        info_msg = msg if "successfully" in msg.lower() else f"Successfully generated:\n{msg}"
        QMessageBox.information(self, "Export Complete", info_msg)

        target_path = None
        for line in reversed(str(msg).splitlines()):
            candidate = Path(line.strip().strip('"').strip("'"))
            if candidate.exists():
                target_path = candidate
                break

        if not target_path or not target_path.exists():
            fallback_dir = self.meeting_dir / "Export"
            if fallback_dir.exists():
                target_path = fallback_dir

        if target_path and target_path.exists():
            _open_folder(target_path)

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
            from modules.word_tools.core.document_opener import select_best_word_document

            files = getattr(thread, "extracted_doc_paths", [])
            best_word = select_best_word_document(files, thread.target_filename)
            if best_word:
                mgr = ComparisonManager.get_instance()
                if getattr(thread, "auto_compare_active_word", False):
                    mgr.set_slot_b(thread.target_filename, str(best_word))
                    if hasattr(self, "cart_frame") and hasattr(self.cart_frame, "run_comparison"):
                        self.cart_frame.run_comparison()
                else:
                    mgr.add_to_cart(thread.target_filename, str(best_word))
            else:
                found_exts = ", ".join(sorted({Path(f).suffix.lower() for f in files})) or "none"
                QMessageBox.warning(
                    self,
                    "Compare Failed",
                    f"No Word document (.docx / .doc) found in {thread.target_filename}.zip (found: {found_exts}).",
                )

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
                self._start_background_wi_enrichment()
            else:
                QMessageBox.warning(self, "Parse Error", "Downloaded, but could not parse the Excel file.")
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to refresh TDocs:\n{result}")

    def _start_background_wi_enrichment(self):
        """Dispatches an unnoticeable, asynchronous background thread to enrich missing Work Items."""
        try:
            if not getattr(self, "filepath", None):
                return
            excel_path = Path(self.filepath)
            if not excel_path.exists():
                return
            if "wi_enrichment" in self.active_threads:
                existing = self.active_threads["wi_enrichment"]
                if existing.isRunning():
                    return
            thread = MeetingWIEnricherThread(excel_path, parent=self)
            thread.finished_enrichment.connect(self._on_wi_enrichment_finished)
            self.active_threads["wi_enrichment"] = thread
            thread.start()
        except Exception as e:
            logging.debug(f"[TDocsWindow] Failed to dispatch background WI enrichment: {e}")

    def _on_wi_enrichment_finished(self, success: bool, count: int, msg: str):
        if "wi_enrichment" in self.active_threads:
            del self.active_threads["wi_enrichment"]
        if success and count > 0:
            try:
                self._load_tracked_wi_tokens()
                self._update_filter_facets()
            except Exception as e:
                logging.debug(f"[TDocsWindow] Filter refresh after WI enrichment: {e}")

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

    def _on_export_menu_about_to_show(self):
        export_dir = self.meeting_dir / "Export"
        has_export = export_dir.is_dir()
        self.export_menu_sep.setVisible(has_export)
        self.act_open_export_dir.setVisible(has_export)

    def _open_export_folder(self):
        export_dir = self.meeting_dir / "Export"
        if export_dir.is_dir():
            _open_folder(export_dir)
        else:
            QMessageBox.information(
                self,
                "Export Folder Not Found",
                f"No export folder exists yet for this meeting:\n{export_dir}\n\n"
                "Run an export first to create it.",
            )

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
        layout.addWidget(QLabel("The following raw 'Source' strings were not recognized and grouped as 'Other':"))

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(unmatched))
        layout.addWidget(text_edit)

        btn = QPushButton("Close")
        btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn.clicked.connect(dialog.accept)
        layout.addWidget(btn)

        dialog.exec_()

    def _open_agenda_file(self):
        _open_folder(self.meeting_dir / "Agenda" / "TdocsByAgenda.htm")

    def _open_excel(self):
        _open_folder(Path(self.filepath))

    def _open_draft_cr_dialog(self, tdoc_row_data: dict = None):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

        mtg_info = {
            "wg_name": self.mtg_info.get("wg_name", ""),
            "meeting_number": self.mtg_info.get("meeting_number", ""),
            "location": self.mtg_info.get("location", ""),
            "dates": self.mtg_info.get("dates", ""),
            "start_date": self.mtg_info.get("start_date", ""),
            "end_date": self.mtg_info.get("end_date", ""),
            "title": self.mtg_info.get("title", ""),
        }

        tdoc_info = tdoc_row_data or {}
        spec_num = tdoc_info.get("Spec", "") or tdoc_info.get("Spec#", "")
        version = tdoc_info.get("Version", "")

        dialog = CrPcrDialog(
            spec_num=spec_num,
            version=version,
            meeting_info=mtg_info,
            tdoc_info=tdoc_info,
            parent=self,
        )
        dialog.show()

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

    def _export_markdown_visible(self):
        export_markdown_visible(self)

    # Backwards compatibility alias
    _export_llm_visible = _export_markdown_visible

    def _export_markdown_single(self, tdoc_id: str):
        export_markdown_single(self, tdoc_id)

    # Backwards compatibility alias
    _export_llm_single = _export_markdown_single


    # Backwards compatibility alias
    _export_llm_single = _export_markdown_single

    def _open_post_meeting_export_dialog(self):
        open_post_meeting_export_dialog(self)

    def _on_md_single_finished(self, success: bool, msg: str):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Markdown Export", msg)
            target = getattr(self.md_export_thread, "export_dir", self.meeting_dir / "Export")
            _open_folder(target if target.exists() else self.meeting_dir / "Export")
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    _on_llm_single_finished = _on_md_single_finished

    def _on_md_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:35])

    _on_llm_progress = _on_md_progress

    def _on_md_export_finished(self, success: bool, msg: str):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Markdown Export Complete", msg)
            target = getattr(self.md_export_thread, "export_dir", self.meeting_dir / "Export")
            _open_folder(target if target.exists() else self.meeting_dir / "Export")
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    _on_llm_export_finished = _on_md_export_finished

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
            self.routing_indicator.setStyleSheet("color: #065F46; background-color: #D1FAE5; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the high-speed local 10.10.10.10 network.")
        elif is_active:
            self.routing_indicator.setText("🔵 SYNC Folder")
            self.routing_indicator.setStyleSheet("color: #1E40AF; background-color: #DBEAFE; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the active meeting SYNC folder.")
        else:
            self.routing_indicator.setText("🌐 Standard Web")
            self.routing_indicator.setStyleSheet("color: #334155; background-color: #F1F5F9; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the standard 3GPP web archive.")

    def _on_instant_fetch(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._trigger_download_thread(base_tdoc, target_filename, legacy_url=None, is_silent_compare=False)

    def _on_instant_compare_active_word(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._compare_tdoc_with_active_word(base_tdoc, target_filename)

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
        import_word_agenda_dialog(self)

    def _process_imported_agenda_files(self, source_paths: list):
        process_imported_agenda_files(self, source_paths)

    def _process_next_in_import_queue(self):
        process_next_in_import_queue(self)

    def _on_word_import_progress(self, msg: str):
        on_word_import_progress(self, msg)

    def _on_word_import_finished(self, success: bool, agenda_data: dict, filename: str, error_msg: str):
        on_word_import_finished(self, success, agenda_data, filename, error_msg)

    def _finalize_import_batch(self):
        finalize_import_batch(self)

    def _download_agenda_csv(self):
        download_agenda_csv(self)

    def _on_agenda_download_finished(self, success: bool, msg: str):
        on_agenda_download_finished(self, success, msg)

    def _download_chair_notes(self):
        download_chair_notes(self)

    def _on_chair_notes_finished(self, success: bool, count: int, files: list, msg: str):
        on_chair_notes_finished(self, success, count, files, msg)

    def _open_chair_notes_folder(self):
        open_chair_notes_folder(self)

    def _get_package_export_rows(self) -> list:
        return get_package_export_rows(self)

    def _export_tdocs_package(self):
        export_tdocs_package(self)

    def _on_package_export_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:34])

    def _open_excel_export_dialog(self):
        open_excel_export_dialog(self)

    def _on_excel_export_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:25])


    def _refresh_email_counts(self):
        try:
            counts = self.general_email_db.get_email_counts_per_tdoc()
            self.model.set_email_counts(counts)
        except Exception as e:
            logging.error(f"Error refreshing email counts: {e}")

    def _open_tdoc_emails(self, tdoc_id: str):
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_clean, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_clean

        if base_tdoc in self._email_dialogs:
            dlg = self._email_dialogs[base_tdoc]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return

        family = self.model.get_family_tdocs(base_tdoc)
        wg = self.mtg_info.get("wg_name", "SA2")

        # Save callback supporting partial note appends without forcing a status overwrite
        def _save_notes_proxy(target_id: str, new_status: str, new_notes: str):
            existing_row = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == target_id.upper()), {})
            status_to_use = new_status if new_status else existing_row.get("My Status", "⚪ Neutral")
            curr_notes = str(existing_row.get("My Notes", "")).strip()
            merged_notes = f"{curr_notes}\n\n{new_notes}" if curr_notes else new_notes
            self._save_user_data(target_id, status_to_use, merged_notes)

        dialog = TDocEmailsDialog(
            base_tdoc,
            family,
            self.general_email_db.db_path,
            wg=wg,
            save_callback=_save_notes_proxy,
            parent=None
        )
        dialog.data_changed.connect(self._refresh_email_counts)
        dialog.tdoc_selected.connect(self._open_tdoc_emails)

        self._email_dialogs[base_tdoc] = dialog
        dialog.finished.connect(lambda _, t=base_tdoc: self._email_dialogs.pop(t, None))
        dialog.show()

    def _on_configure_email_folders(self):
        wg = self.mtg_info.get("wg_name", "SA2")
        dialog = GeneralEmailFoldersDialog(wg, self)
        dialog.exec_()

    def _on_sync_general_emails(self):
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
        self.email_btn.setEnabled(True)
        self.email_btn.setText("📧 Emails ▾")
        self._refresh_email_counts()
        if success:
            QMessageBox.information(self, "Email Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Email Sync Failed", msg)

    def _on_mark_all_emails_read(self):
        self.general_email_db.mark_all_read()
        self._refresh_email_counts()
        QMessageBox.information(self, "Mark All Read", "All emails have been marked as read.")

    def _on_wipe_generic_emails(self):
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
        handle_table_context_menu(self, pos)


    def _open_tdoc_triage(self, tdoc_id: str):
        """Opens a modeless AI Triage dialog for the selected TDoc."""
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()

        if tdoc_clean in self._triage_dialogs:
            dlg = self._triage_dialogs[tdoc_clean]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return

        # Find the row data for this TDoc or its base
        row_data = next(
            (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == tdoc_clean),
            None
        )
        if not row_data:
            match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_clean, re.IGNORECASE)
            base_tdoc = match.group(1).upper() if match else tdoc_clean
            row_data = next(
                (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc),
                {"TDoc": tdoc_clean}
            )

        dialog = TDocTriageDialog(
            tdoc_data=row_data,
            meeting_dir=self.meeting_dir,
            docs_ftp_url=self.docs_ftp_url,
            revisions_url=self.revisions_url,
            save_callback=self._save_user_data,
            parent=None  # Independent top-level modeless window
        )
        self._triage_dialogs[tdoc_clean] = dialog
        dialog.finished.connect(lambda _, t=tdoc_clean: self._triage_dialogs.pop(t, None))
        dialog.show()

    def closeEvent(self, event):
        if hasattr(self, 'routing_timer'):
            self.routing_timer.stop()
        if hasattr(self, 'search_timer'):
            self.search_timer.stop()

        if getattr(self, 'email_window', None):
            try:
                self.email_window.close()
            except Exception:
                pass

        for dlg in list(getattr(self, '_email_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_email_dialogs'):
            self._email_dialogs.clear()

        for dlg in list(getattr(self, '_triage_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_triage_dialogs'):
            self._triage_dialogs.clear()

        all_threads = [
            getattr(self, 'rev_thread', None),
            getattr(self, 'agenda_dl_thread', None),
            getattr(self, 'general_email_sync_thread', None),
            getattr(self, 'dl_thread', None),
            getattr(self, 'package_export_thread', None),
            getattr(self, 'package_excel_exporter_thread', None),
            getattr(self, 'export_thread', None),
            getattr(self, 'excel_exporter_thread', None),
            getattr(self, 'stats_thread', None),
            getattr(self, 'agenda_thread', None),
            getattr(self, 'llm_thread', None),
            getattr(self, 'word_import_thread', None),
            getattr(self, 'chair_notes_thread', None),
        ]
        if hasattr(self, 'active_threads'):
            all_threads.extend(self.active_threads.values())

        for th in filter(None, all_threads):
            if hasattr(th, 'is_cancelled'):
                th.is_cancelled = True
            if th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(100)

        super().closeEvent(event)
