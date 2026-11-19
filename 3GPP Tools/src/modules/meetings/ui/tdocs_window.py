# --- File: src/modules/meetings/ui/tdocs_window.py ---
import datetime
import json
import logging
import os
import re
import urllib.parse
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QRect, QTimer, pyqtSignal, QPoint
from PyQt5.QtGui import QCursor, QKeySequence, QTextDocument
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableView,
    QHeaderView, QLabel, QLineEdit, QFrame,
    QPushButton, QMessageBox, QMenu, QApplication, QShortcut,
    QToolTip, QCheckBox, QDialog, QTextEdit, QComboBox, QSizePolicy
)

from core.network.network_state import NetworkState
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import (
    get_font_scale_notifier,
    scale_px,
)
from core.ui.world_clock_service import WorldClockService
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY
)
from modules.emails.ui.email_window import EmailManagerWindow
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.exporters.markdown_exporter import MarkdownExporterThread
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.stats.exporter_thread import StatisticsExporterThread
from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.core.meeting_wi_enricher import MeetingWIEnricherThread
from modules.meetings.core.tdocs.tdocs_threads import (
    TDocsRevisionsFetcherThread,
    TDocActionThread,
    TdocsByAgendaThread,
)
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdoc_delegates import (
    HtmlDelegate,
    TDocActionDelegate,
    AgendaItemDelegate,
    SourceCompanyDelegate,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox, ComparisonCartWidget, make_rich_tooltip
from modules.meetings.ui.tdocs_dialogs import (
    ReadOnlyViewerDialog,
    InteractiveNotesDialog,
    StatisticsSettingsDialog, TDocInfoDialog
)
from modules.meetings.ui.meetings_menu import add_meeting_portal_actions
from modules.meetings.ui.tdocs_menus import build_action_menu, build_related_menu
from modules.meetings.ui.tdocs_menu import handle_table_context_menu
from modules.meetings.ui.dialogs.tdocs_export_controller import (
    get_package_export_rows,
    export_tdocs_package,
    open_excel_export_dialog,
    export_markdown_visible,
    export_markdown_single,
    open_post_meeting_export_dialog,
    open_pre_meeting_export_dialog,
    launch_cli_for_tdocs,
)
from core.ai.cli_launcher import populate_cli_agent_menu
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel, natural_sort_key
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    pin_thread_by_subject_or_cid,
    pin_thread_from_msg,
    pin_thread_from_outlook_selection,
)
from modules.emails.core.tracked_threads_workers import TrackedThreadsRefreshWorker
from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar
from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog
from modules.emails.ui.tracked_thread_dialog import (
    PinFromOutlookDialog,
    TrackedThreadDialog,
)
from core.ui.toast import notify_info, notify_success, notify_toast
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
    scan_existing_chair_notes,
    on_chair_notes_finished,
    open_chair_notes_folder,
    start_window_tdocs_caching,
)
from core.ui.theme import C


class TDocsWindow(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}
        self.chair_agreements = []
        self._email_dialogs = {}
        self._triage_dialogs = {}
        self._highlight_clear_timer = QTimer(self)
        self._highlight_clear_timer.setSingleShot(True)
        self._highlight_clear_timer.setInterval(4500)
        self._highlight_clear_timer.timeout.connect(
            lambda: self.model.set_highlighted_tdoc(None) if hasattr(self, "model") else None
        )

        self.agenda_dir = Path(filepath).parent
        self.agenda_map = AgendaManager.load_local_agenda(self.agenda_dir)
        self.agenda_dl_thread = None

        self.general_email_db = GeneralEmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.general_email_sync_thread = None
        self.tracked_threads_db = TrackedThreadsDatabase(self.meeting_dir / "Agenda" / "tracked_threads.db")
        self._tracked_thread_dialogs = {}
        self.tracked_threads_timer = QTimer(self)
        self.tracked_threads_timer.setInterval(300000)  # 5 minutes auto-refresh
        self.tracked_threads_timer.timeout.connect(lambda: self._on_refresh_tracked_threads(silent=True))

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
        self.is_ran = bool(re.search(r'\bRAN\d?\b', wg_name) or wg_name.startswith("RAN"))
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

        self._row_height_timer = QTimer(self)
        self._row_height_timer.setSingleShot(True)
        self._row_height_timer.setInterval(35)
        self._row_height_timer.timeout.connect(self._adjust_table_row_heights)
        self._user_resized_cols = False
        self._programmatic_column_resize = False

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
        self._schedule_column_fit()
        self._schedule_row_height_adjustment()

    def _setup_header(self, layout, title, is_electronic, count):
        header_container = QWidget(self)
        header_container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        header_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        header_sp.setHeightForWidth(True)
        header_container.setSizePolicy(header_sp)
        header_layout = FlowLayout(header_container, margin=0, h_spacing=6, v_spacing=6)
        self._header_container = header_container
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet(f"font-size: 16px; color: {C.TEXT_STRONG};")
        title_lbl.setToolTip("Electronic Meeting (eMeeting)" if is_electronic else "In-Person Meeting (Face-to-Face)")
        title_lbl.setCursor(Qt.WhatsThisCursor)

        self.routing_indicator = QLabel("⚪ Routing...")
        self.routing_indicator.setStyleSheet("font-weight: bold; padding: 2px 6px; border-radius: 4px;")

        self._clock_service = WorldClockService()
        self.venue_clock_btn = QPushButton("📍 --:--")
        self.venue_clock_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.venue_clock_btn.setCursor(Qt.PointingHandCursor)
        self.venue_clock_btn.clicked.connect(self._open_venue_clock_dialog)

        self.routing_timer = QTimer(self)
        self.routing_timer.timeout.connect(self._update_routing_indicator)
        self.routing_timer.start(2000)
        self._update_routing_indicator()

        self.last_mod_lbl = QLabel(self._get_mod_date_str())
        self.last_mod_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; margin-right: 8px; font-style: italic;")

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.setToolTip(make_rich_tooltip(
            title="🔄 Refresh Meeting Data",
            description="Reload official TDoc list, download Agenda CSV, cache documents offline, or scan notes.",
            tip="Click arrow ▾ to view options including offline caching and notes scanning."
        ))

        refresh_menu = QMenu(self)
        refresh_menu.setToolTipsVisible(True)
        act = refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        act.setToolTip("Download and reload the latest official TDoc Excel list for this meeting.")
        act = refresh_menu.addAction("📋 Download Agenda CSV", self._download_agenda_csv)
        act.setToolTip("Download agenda.csv and refresh Agenda Item names/acronyms used by the TDoc view.")
        act = refresh_menu.addAction("📥 Cache All TDocs (Docs/)", self._cache_all_tdocs)
        act.setToolTip("Download and cache all meeting TDocs (.zip) from the 3GPP Docs directory locally using parallel downloads.")
        act = refresh_menu.addAction("📥 Download Chairman's Notes", self._download_chair_notes)
        act.setToolTip("Browse and download Chairman's Notes from /Inbox/Chair_Notes with pre-selected latest files per chair/track.")
        if self.is_ran:
            act = refresh_menu.addAction("🔍 Scan Existing Chairman's Notes for Agreements", self._scan_existing_chair_notes)
            act.setToolTip("Scan local Chairman's Notes files (.docx) in Agenda/ChairNotes, extract current agreements, and merge into the meeting table.")
        act = refresh_menu.addAction("📝 Import Chairman's Notes / Agenda (.docx / .doc / .zip)...", self._import_word_agenda_dialog)
        act.setToolTip("Import one or more local Word, ZIP, or HTML agenda/notes files and merge their TDoc remarks into the table.")
        act = refresh_menu.addAction("⚖️ Track Study Item && Meeting Agreements...", self._open_agreement_tracker)
        act.setToolTip("Open the Cross-Meeting Study Item & Topic Agreement Tracker pre-filtered to this meeting's Working Group.")
        if self.is_sa2:
            act = refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)
            act.setToolTip("Fetch TdocsByAgenda.htm and merge chairman/secretary agenda metadata into the TDoc list.")
        if self.is_sa2_electronic:
            act = refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            act.setToolTip("Refresh the revision list for this electronic meeting without reloading the Excel list.")
            act = refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
            act.setToolTip("Refresh both the official TDoc Excel list and the available revision metadata.")
        self.refresh_btn.setMenu(refresh_menu)

        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.folder_btn.setToolTip(make_rich_tooltip(
            title="🗂️ Resources & Directories",
            description="Quick access to local cache directories (Docs, Agenda, ChairNotes, SessionPlans), remote 3GPP FTP folders, and official 3GU Meeting Portal links.",
            tip="You can also open a Windows Terminal directly inside the local meeting cache folder."
        ))

        self.folder_menu = QMenu(self)
        self.folder_menu.setToolTipsVisible(True)
        self.folder_menu.aboutToShow.connect(self._populate_resources_menu)
        self.folder_btn.setMenu(self.folder_menu)

        self.excel_btn = QPushButton("📗 Original Excel")
        self.excel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excel_btn.setToolTip(make_rich_tooltip(
            title="📗 Open Original Excel",
            description="Launch Microsoft Excel with the raw, official 3GPP TDoc list spreadsheet for this meeting.",
            tip="Use this to view unedited official committee columns or run external Excel formulas."
        ))
        self.excel_btn.clicked.connect(self._open_excel)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.export_btn.setToolTip(make_rich_tooltip(
            title="📤 Export Reports & Dossiers",
            shortcut="Ctrl+Shift+E",
            description="Export meeting TDoc metadata, document packages, LLM-ready content, or Markdown reports.",
            tip="Press Ctrl+Shift+E anytime to reveal the export folder in Windows File Explorer."
        ))

        self.session_plan_btn = QPushButton("🗓️ Session Plan")
        self.session_plan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.session_plan_btn.setToolTip(make_rich_tooltip(
            title="🗓️ Session Plan & Timetable",
            description="Interactive visual schedule viewer parsing session plan documents. Displays parallel tracks, room locations, and timings.",
            tip="Automatically refreshes when new session plan documents are uploaded during the meeting."
        ))
        self.session_plan_btn.clicked.connect(self._open_session_plan_window)

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

        act_export_pre_meeting = export_menu.addAction(
            "🧭 Export Pre-Meeting Report (.md)...", self._open_pre_meeting_export_dialog
        )
        act_export_pre_meeting.setToolTip(
            "Generate a structured, per-release and per-WI pre-meeting Markdown briefing covering incoming LS, "
            "competing LS reply proposals, proposed CRs/pCRs, and company discussion papers."
        )
        act_export_pre_meeting.setStatusTip(act_export_pre_meeting.toolTip())

        act_export_post_meeting = export_menu.addAction(
            "📑 Export Post-Meeting Report (.md)...", self._open_post_meeting_export_dialog
        )
        act_export_post_meeting.setToolTip(
            "Generate a structured, per-release and per-WI post-meeting Markdown dossier covering incoming LS and agreed pCR/CR/LS."
        )
        act_export_post_meeting.setStatusTip(act_export_post_meeting.toolTip())

        act_export_dossier = export_menu.addAction(
            "📑 Export Navigable Dossier (HTML && PDF)...", self._export_dossier
        )
        act_export_dossier.setToolTip(
            "Compile selected or visible TDocs into an interactive HTML dossier and a bookmarked vector PDF with companion AI citation prompt."
        )
        act_export_dossier.setStatusTip(act_export_dossier.toolTip())

        self.cli_export_submenu = export_menu.addMenu("💬 Export && Open in CLI Agent")
        self.cli_export_submenu.setToolTip(
            "Pre-convert selected (or visible) TDocs to Markdown and launch Codex, Claude, or Gemini CLI."
        )
        self.cli_export_submenu.aboutToShow.connect(self._populate_export_cli_submenu)

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

        self.btn_meeting_agreements = QPushButton("⚖️ Agreements")
        self.btn_meeting_agreements.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_meeting_agreements.setToolTip(make_rich_tooltip(
            title="⚖️ Chairman's Notes Agreements (RAN)",
            description="View, search, and export all Chairman's Notes agreements and working assumptions for this meeting.",
            tip="Double-clicking an agreement cell in the table will open this viewer focused on that agreement."
        ))
        self.btn_meeting_agreements.clicked.connect(lambda: self._open_meeting_agreements_dialog())

        self.draft_cr_btn = QPushButton("📝 Draft CR / pCR")
        self.draft_cr_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.draft_cr_btn.setToolTip(make_rich_tooltip(
            title="📝 Draft CR / pCR Contribution",
            description="Compose an official Change Request (CR) or Pseudo-CR (pCR) contribution adhering strictly to MCC formatting rules.",
            tip="Select a baseline spec or TDoc in the table to pre-populate revision headers and clauses."
        ))
        self.draft_cr_btn.clicked.connect(lambda: self._open_draft_cr_dialog())

        self.stats_btn = QPushButton("📊 Statistics")
        self.stats_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.stats_btn.setToolTip(make_rich_tooltip(
            title="📊 Meeting Statistics & Analytics",
            description="Generate an interactive HTML analytics dashboard showing company contribution volumes, document types, and agenda distributions.",
            tip="Click ⚙️ to configure custom company groupings or export parameters."
        ))
        self.stats_btn.clicked.connect(self._generate_statistics)

        self.stats_cfg_btn = QPushButton("⚙")
        self.stats_cfg_btn.setStyleSheet(
            BUTTON_STYLE_TOOLBAR_SECONDARY
            + "\nQPushButton { padding: 4px 8px; font-size: 13px; }"
        )
        self.stats_cfg_btn.setMinimumWidth(scale_px(34))
        self.stats_cfg_btn.setToolTip(make_rich_tooltip(
            title="⚙️ Configure Statistics",
            description="Customize company alias unification, contribution categorization, and output charts."
        ))
        self.stats_cfg_btn.clicked.connect(self._open_stats_config)

        self.email_btn = QPushButton("📧 Emails ▾")
        self.email_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.email_btn.setToolTip(make_rich_tooltip(
            title="📧 Email Discussions & Reflector Sync",
            description="Sync and inspect Outlook emails linked to TDocs and revisions, strip quotes, and categorize stances.",
            tip="Discussion counts are displayed in the 'Emails' table column with stance breakdown badges."
        ))

        self.btn_help = QPushButton("❓ Help")
        self.btn_help.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_help.setToolTip(make_rich_tooltip(
            title="❓ Help & Shortcuts Guide",
            shortcut="F1",
            description="Open the comprehensive interactive user guide, table double-click cheatsheet, search syntax, and keyboard shortcuts.",
            tip="Press F1 anywhere in the TDocs window to open instantly."
        ))
        self.btn_help.clicked.connect(self._open_help_dialog)

        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self._open_help_dialog)

        email_menu = QMenu(self)
        email_menu.setToolTipsVisible(True)
        act = email_menu.addAction("📬 Unread && Flagged Email Inbox...", self._open_meeting_email_inbox)
        act.setToolTip("Open the unified Email Triage Cockpit combining all unread, flagged, and recent emails across Tracked Threads and TDoc Emails.")
        email_menu.addSeparator()
        act = email_menu.addAction("📌 Pin Email Thread from Outlook Folder...", self._on_pin_thread_from_outlook)
        act.setToolTip("Browse and pin email discussion threads from your configured WG Outlook folders.")
        act = email_menu.addAction("📥 Pin Email Thread from .msg File...", self._on_pin_thread_from_msg_dialog)
        act.setToolTip("Select a local .msg file, locate its full thread in configured Outlook folders, and pin it to the meeting table.")
        act = email_menu.addAction("🔄 Refresh Tracked Email Threads Now", lambda: self._on_refresh_tracked_threads(silent=False))
        act.setToolTip("Check configured Outlook folders for new replies on all pinned/tracked email threads.")
        self.act_auto_refresh_threads = email_menu.addAction("⏱️ Auto-Refresh Tracked Threads (Every 5m)")
        self.act_auto_refresh_threads.setCheckable(True)
        self.act_auto_refresh_threads.setChecked(False)
        self.act_auto_refresh_threads.toggled.connect(self._on_toggle_auto_refresh_threads)
        email_menu.addSeparator()
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
        self.count_lbl.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {C.TEXT_SECONDARY};")

        self.instant_fetch_input = QLineEdit()
        self.instant_fetch_input.setPlaceholderText("Instant Fetch...")
        self.instant_fetch_input.setToolTip(make_rich_tooltip(
            title="⚡ Instant TDoc Fetch",
            shortcut="Enter",
            description="Type any TDoc number (e.g. R2-2601234 or S2-2605678) and press Enter to download, unpack, and launch it immediately.",
            tip="Bypasses table scrolling completely when documents are called up live in the meeting room."
        ))
        self.instant_fetch_input.setText(self._get_tdoc_prefix())
        self.instant_fetch_input.setMinimumWidth(scale_px(120))
        self.instant_fetch_input.setMaximumWidth(scale_px(160))
        self.instant_fetch_input.returnPressed.connect(self._on_instant_fetch)
        self.instant_fetch_input.setFocus()
        self.instant_fetch_input.setCursorPosition(len(self.instant_fetch_input.text()))

        self.instant_cmp_btn = QPushButton("⚖️ vs Word")
        self.instant_cmp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.instant_cmp_btn.setCursor(Qt.PointingHandCursor)
        self.instant_cmp_btn.setToolTip(make_rich_tooltip(
            title="⚖️ Compare with Active Word",
            description="Download the TDoc entered in Instant Fetch and immediately run a Track Changes Word comparison against your active Word draft.",
            tip="Perfect for reviewing a newly posted revision against your local edits in seconds."
        ))
        self.instant_cmp_btn.clicked.connect(self._on_instant_compare_active_word)

        identity_group = QWidget(header_container)
        identity_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        id_layout = QHBoxLayout(identity_group)
        id_layout.setContentsMargins(0, 0, 2, 0)
        id_layout.setSpacing(6)
        id_layout.addWidget(title_lbl)
        id_layout.addWidget(self.routing_indicator)
        id_layout.addWidget(self.venue_clock_btn)
        header_layout.addWidget(identity_group)

        instant_group = QWidget(header_container)
        instant_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        inst_layout = QHBoxLayout(instant_group)
        inst_layout.setContentsMargins(0, 0, 2, 0)
        inst_layout.setSpacing(4)
        inst_layout.addWidget(QLabel("🚀"))
        inst_layout.addWidget(self.instant_fetch_input)
        inst_layout.addWidget(self.instant_cmp_btn)
        header_layout.addWidget(instant_group)

        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addWidget(self.export_btn)
        header_layout.addWidget(self.session_plan_btn)
        if self.is_ran:
            header_layout.addWidget(self.btn_meeting_agreements)
        else:
            self.btn_meeting_agreements.setVisible(False)
        header_layout.addWidget(self.draft_cr_btn)

        stats_group = QWidget(header_container)
        stats_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        stats_layout = QHBoxLayout(stats_group)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setSpacing(2)
        stats_layout.addWidget(self.stats_btn)
        stats_layout.addWidget(self.stats_cfg_btn)
        header_layout.addWidget(stats_group)

        header_layout.addWidget(self.email_btn)
        header_layout.addWidget(self.btn_help)
        layout.addWidget(header_container)

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
        session_plans_dir = self.meeting_dir / "Agenda" / "SessionPlans"
        if session_plans_dir.is_dir():
            self.folder_menu.addAction("🗓️ Local: Session Plans Folder", lambda: _open_folder(session_plans_dir))
        chair_notes_dir = self.meeting_dir / "Agenda" / "ChairNotes"
        if chair_notes_dir.is_dir():
            local_docs = []
            try:
                local_docs = [
                    f for f in chair_notes_dir.iterdir()
                    if f.is_file() and f.suffix.lower() in ('.docx', '.doc') and not f.name.startswith(('~$', '.'))
                ]
            except Exception:
                local_docs = []

            has_any_file = False
            try:
                has_any_file = any(chair_notes_dir.iterdir())
            except Exception:
                pass

            if has_any_file:
                self.folder_menu.addAction("📂 Local: ChairNotes Folder", self._open_chair_notes_folder)

            if local_docs:
                if len(local_docs) == 1:
                    doc = local_docs[0]
                    act = self.folder_menu.addAction(f"📝 Open Notes: {doc.name}", lambda d=doc: self._open_doc_in_word(d))
                    act.setToolTip(f"Open {doc.name} in Microsoft Word.")
                else:
                    notes_sub = self.folder_menu.addMenu("📝 Open Chairman's Notes (.docx)")
                    notes_sub.setToolTipsVisible(True)
                    sorted_docs = sorted(local_docs, key=lambda f: f.stat().st_mtime, reverse=True)
                    for doc in sorted_docs:
                        act = notes_sub.addAction(doc.name, lambda d=doc: self._open_doc_in_word(d))
                        act.setToolTip(f"Open {doc.name} in Microsoft Word.")
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

        has_ftp = bool(self.main_ftp_url or self.docs_ftp_url or self.revisions_url)
        if has_ftp:
            self.folder_menu.addSeparator()
            if self.main_ftp_url:
                self.folder_menu.addAction("🌐 Web FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
            if self.docs_ftp_url:
                self.folder_menu.addAction("🌐 Web FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
                act_cache = self.folder_menu.addAction("📥 Cache TDocs (Docs/)", self._cache_all_tdocs)
                act_cache.setToolTip("Cache all official TDoc archives from the 3GPP Docs directory into the local meeting cache.")
            if self.revisions_url:
                self.folder_menu.addAction("🌐 Web FTP: Revisions Folder", lambda: webbrowser.open(self.revisions_url))

        if self.mtg_info.get("mtg_id"):
            self.folder_menu.addSeparator()
            add_meeting_portal_actions(self.folder_menu, self.mtg_info)

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, legacy_url: str = None,
                                 is_silent_compare: bool = False, auto_compare_active_word: bool = False,
                                 sync_table_on_success: bool = False, send_email_on_finish: bool = False):
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
                                  open_file=not (is_silent_compare or send_email_on_finish))
        thread.is_silent_compare = is_silent_compare
        thread.auto_compare_active_word = auto_compare_active_word
        thread.sync_table_on_success = sync_table_on_success
        thread.send_email_on_finish = send_email_on_finish
        thread.target_filename = target_filename
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))

        self.active_threads[base_tdoc] = thread
        thread.start()

    def _compare_tdoc_with_active_word(self, base_tdoc: str, target_filename: str,
                                       sync_table_on_success: bool = False):
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
            sync_table_on_success=sync_table_on_success,
        )

    def _setup_filters(self, layout):
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self._apply_search_filter)

        filter_frame = QFrame()
        filter_frame.setObjectName("tdocFilterFrame")
        filter_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        filter_sp.setHeightForWidth(True)
        filter_frame.setSizePolicy(filter_sp)
        filter_frame.setStyleSheet(f"""
            QFrame#tdocFilterFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
            QLabel {{
                font-weight: bold;
                color: {C.TEXT_SECONDARY};
                border: none;
            }}
            QLineEdit, QComboBox {{
                padding: 4px 8px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background: {C.BG_CARD};
            }}
            QComboBox[activeFilter="true"] {{
                background-color: {C.INFO_SOFT};
                border: 1.5px solid {C.INFO_ACCENT};
                color: {C.INFO};
                font-weight: bold;
            }}
        """)
        filter_layout = FlowLayout(filter_frame, margin=0, h_spacing=5, v_spacing=4)
        filter_layout.setContentsMargins(6, 4, 6, 4)
        self._filter_frame = filter_frame

        search_group = QWidget(filter_frame)
        search_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(5)

        search_lbl = QLabel("🔍")
        search_lbl.setToolTip("Full-text search across TDoc number, title, source, remarks, and related metadata.")
        search_layout.addWidget(search_lbl)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search TDocs (e.g. "CR", -draft)...')
        self.search_input.setToolTip(
            'Search across TDoc metadata.\n'
            '• Exact phrases: "draft baseline"\n'
            '• Exclude words: -draft or !draft\n'
            '• Substrings & numbers: S2-260855, TS 23.501'
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(scale_px(180))
        self.search_input.setMaximumWidth(scale_px(240))
        self.search_input.textChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.search_input, stretch=1)

        self.btn_followed_wis = QPushButton("🎯")
        self.btn_followed_wis.setCheckable(True)
        self.btn_followed_wis.setCursor(Qt.PointingHandCursor)
        self.btn_followed_wis.setFixedSize(scale_px(28), scale_px(28))
        self.btn_followed_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_followed_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_followed_wis.setToolTip(make_rich_tooltip(
            title="🎯 Followed Work Items Filter",
            description="Filter to TDocs associated with Work Items you have selected to follow.",
            tip="Right-click to manage your Followed Work Items watchlist."
        ))
        self.btn_followed_wis.clicked.connect(self._toggle_followed_wis_filter)
        search_layout.addWidget(self.btn_followed_wis)

        self.btn_starred_wis = QPushButton("⭐")
        self.btn_starred_wis.setCheckable(True)
        self.btn_starred_wis.setCursor(Qt.PointingHandCursor)
        self.btn_starred_wis.setFixedSize(scale_px(28), scale_px(28))
        self.btn_starred_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_starred_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_starred_wis.setToolTip(make_rich_tooltip(
            title="⭐ Starred Work Items Filter",
            description="Filter to TDocs associated with Work Items you have starred (personal priority watchlist).",
            tip="Right-click to toggle including unassigned incoming Liaison Statements (LS)."
        ))
        self.btn_starred_wis.clicked.connect(self._toggle_starred_wis_filter)
        search_layout.addWidget(self.btn_starred_wis)

        self.btn_agreements_filter = QPushButton("⚖️")
        self.btn_agreements_filter.setCheckable(True)
        self.btn_agreements_filter.setCursor(Qt.PointingHandCursor)
        self.btn_agreements_filter.setFixedSize(scale_px(28), scale_px(28))
        self.btn_agreements_filter.setToolTip(make_rich_tooltip(
            title="⚖️ Chairman's Notes Agreements Filter (RAN)",
            description="Filter table to only TDocs with Chairman's Notes text agreements or working assumptions.",
            tip="Agreements are highlighted in soft green and working assumptions in warm amber."
        ))
        self.btn_agreements_filter.clicked.connect(self._toggle_agreements_filter)
        if self.is_ran:
            search_layout.addWidget(self.btn_agreements_filter)
        else:
            self.btn_agreements_filter.setVisible(False)

        self.btn_unread_emails = QPushButton("🔵")
        self.btn_unread_emails.setCheckable(True)
        self.btn_unread_emails.setCursor(Qt.PointingHandCursor)
        self.btn_unread_emails.setFixedSize(scale_px(28), scale_px(28))
        self.btn_unread_emails.setToolTip(make_rich_tooltip(
            title="🔵 Unread Emails Filter",
            description="Show only TDocs and Tracked Email Threads that have unread emails.",
            tip="Click again to clear, or click 🚩 to switch to flagged emails."
        ))
        self.btn_unread_emails.clicked.connect(self._toggle_unread_emails_filter)
        search_layout.addWidget(self.btn_unread_emails)

        self.btn_flagged_emails = QPushButton("🚩")
        self.btn_flagged_emails.setCheckable(True)
        self.btn_flagged_emails.setCursor(Qt.PointingHandCursor)
        self.btn_flagged_emails.setFixedSize(scale_px(28), scale_px(28))
        self.btn_flagged_emails.setToolTip(make_rich_tooltip(
            title="🚩 Flagged Emails Filter",
            description="Show only TDocs and Tracked Email Threads that have Outlook-flagged emails.",
            tip="Click again to clear, or click 🔵 to switch to unread emails."
        ))
        self.btn_flagged_emails.clicked.connect(self._toggle_flagged_emails_filter)
        search_layout.addWidget(self.btn_flagged_emails)

        self._update_tracked_wi_button_styles()
        filter_layout.addWidget(search_group)

        self.company_combo = CheckableComboBox("Company")
        self.company_combo.setMaximumWidth(scale_px(165))
        self.company_combo.setContextMenuPolicy(Qt.CustomContextMenu)
        self.company_combo.customContextMenuRequested.connect(self._show_company_filter_context_menu)
        self.company_combo.setToolTip(
            'Filter by contributing company (excludes unconfirmed `?` and `[...]` sources by default).\n'
            'Right-click to toggle including unconfirmed (?, [...]) carry-over sources.'
        )
        self.company_combo.selectionChanged.connect(self._on_company_changed)
        filter_layout.addWidget(self.company_combo)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMaximumWidth(scale_px(125))
        self.type_combo.setToolTip('Filter by TDoc type. Grey entries currently have no intersection with the other active filters but remain selectable.')
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMaximumWidth(scale_px(130))
        self.ai_combo.setToolTip('Filter by Agenda Item. When agenda.csv is available, associated acronyms are shown in parentheses.')
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.sub_item_combo = CheckableComboBox("Sub-Item")
        self.sub_item_combo.setMaximumWidth(scale_px(150))
        self.sub_item_combo.setToolTip('Filter by Agenda Sub-Item / topic separator from TdocsByAgenda.htm or Chairman notes.')
        self.sub_item_combo.selectionChanged.connect(self._on_sub_item_changed)
        self.sub_item_combo.setVisible(False)
        filter_layout.addWidget(self.sub_item_combo)

        self.wi_combo = CheckableComboBox("WI")
        self.wi_combo.setMaximumWidth(scale_px(135))
        self.wi_combo.setToolTip('Filter by the Related WIs metadata from the official TDoc list. Grey entries currently have no intersection with the other active filters.')
        self.wi_combo.selectionChanged.connect(self._on_wi_changed)
        filter_layout.addWidget(self.wi_combo)

        self.status_combo = CheckableComboBox("TDoc Status")
        self.status_combo.setMaximumWidth(scale_px(155))
        self.status_combo.setToolTip('Filter by the meeting/secretary TDoc status.')
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        self.my_status_combo = CheckableComboBox("My Status")
        self.my_status_combo.setMaximumWidth(scale_px(145))
        self.my_status_combo.setToolTip('Filter by your personal TDoc status.')
        self.my_status_combo.selectionChanged.connect(self._on_my_status_changed)
        filter_layout.addWidget(self.my_status_combo)

        self.tags_combo = CheckableComboBox("Tags")
        self.tags_combo.setMaximumWidth(scale_px(145))
        self.tags_combo.setToolTip('Filter TDocs and tracked email threads by your personal tags.')
        self.tags_combo.selectionChanged.connect(self._on_tags_changed)
        self.tags_combo.setVisible(False)
        filter_layout.addWidget(self.tags_combo)

        tail_group = QWidget(filter_frame)
        tail_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tail_layout = QHBoxLayout(tail_group)
        tail_layout.setContentsMargins(0, 0, 0, 0)
        tail_layout.setSpacing(5)

        if self.is_sa2:
            self.chk_no_comments = QCheckBox("No Comments")
            self.chk_no_comments.setToolTip('Show only TDocs without secretary/chair comments.')
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            tail_layout.addWidget(self.chk_no_comments)

        order_lbl = QLabel("Order:")
        order_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-weight: bold;")
        tail_layout.addWidget(order_lbl)

        self.order_combo = QComboBox()
        self.order_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.order_combo.setMaximumWidth(scale_px(165))
        self.order_combo.setToolTip(
            "Change the display ordering of TDocs.\n"
            "• 📋 Agenda Order: Sort by Agenda Item and Sub-Item discussion order (superseded by TdocsByAgenda.htm).\n"
            "• 🔢 TDoc Number: Sort alphanumerically by TDoc identifier.\n"
            "• 🏢 Company: Group by contributing company.\n"
            "• ✉️ Most Emails: Show most discussed TDoc families first.\n"
            "• 🗂️ Default / File Order: Original document arrival order."
        )
        self.order_combo.addItem("📋 Agenda Order", "AGENDA")
        self.order_combo.addItem("🔢 TDoc Number", "TDOC")
        self.order_combo.addItem("🏢 Company", "COMPANY")
        self.order_combo.addItem("✉️ Most Emails", "EMAILS")
        self.order_combo.addItem("🗂️ Default / File Order", "DEFAULT")
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        tail_layout.addWidget(self.order_combo)

        self.row_mode_combo = QComboBox()
        self.row_mode_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.row_mode_combo.setMaximumWidth(scale_px(135))
        self.row_mode_combo.setToolTip("Filter between all rows, official TDocs only, or pinned 📧 Email Threads only.")
        self.row_mode_combo.addItem("📋 All Rows", "ALL")
        self.row_mode_combo.addItem("📄 TDocs Only", "TDOCS")
        self.row_mode_combo.addItem("📧 Threads Only", "THREADS")
        self.row_mode_combo.currentIndexChanged.connect(
            lambda _: self.proxy.setRowModeFilter(self.row_mode_combo.currentData()) if hasattr(self, "proxy") else None
        )
        tail_layout.addWidget(self.row_mode_combo)

        self.btn_reset_filters = QPushButton("✕ Clear")
        self.btn_reset_filters.setToolTip("Reset all search text and active filters to show all TDocs.")
        self.btn_reset_filters.setCursor(Qt.PointingHandCursor)
        self.btn_reset_filters.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_MUTED};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {C.DANGER_SOFT_HOVER};
                border-color: {C.DANGER_BORDER_ALT};
                color: {C.DANGER_ALT};
            }}
        """)
        self.btn_reset_filters.clicked.connect(self._clear_all_filters)
        tail_layout.addWidget(self.btn_reset_filters)
        filter_layout.addWidget(tail_group)

        layout.addWidget(filter_frame)

        self.thread_sync_bar = AsyncSyncProgressBar(self)
        self.thread_sync_bar.cancel_requested.connect(self._on_cancel_tracked_threads_sync)
        layout.addWidget(self.thread_sync_bar)
        self._tracked_threads_sync_worker = None
        self._sync_silent_mode = False

    def _setup_table(self, layout, data, user_data):
        self.settings = MeetingsSettings()
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, data, user_data, agenda_map=self.agenda_map)
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setIncludeUnassignedLsIn(self.settings.get_include_unassigned_ls_in())
        self.proxy.setIncludeTentativeSources(self.settings.get_include_tentative_sources())
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)
        self.proxy.layoutChanged.connect(self._schedule_row_height_adjustment)
        self.proxy.modelReset.connect(self._schedule_row_height_adjustment)
        self.proxy.rowsInserted.connect(lambda *_: self._schedule_row_height_adjustment())
        self.proxy.rowsRemoved.connect(lambda *_: self._schedule_row_height_adjustment())
        self.model.dataChanged.connect(lambda *_: self._schedule_row_height_adjustment())

        self.table.setModel(self.proxy)
        self.table.setWordWrap(True)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.doubleClicked.connect(self._show_cell_popup)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet(f"""
            QTableView {{
                gridline-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_STRONG};
                background-color: {C.BG_CARD};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                padding: 4px;
                font-weight: bold;
                border: 1px solid {C.BORDER_SUBTLE};
                color: {C.TEXT_STRONG};
            }}
        """)
        vheader = self.table.verticalHeader()
        vheader.setSectionResizeMode(QHeaderView.Fixed)
        vheader.setMinimumSectionSize(scale_px(48))
        vheader.setDefaultSectionSize(scale_px(48))

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.source_delegate = SourceCompanyDelegate(self.table)
        self.table.setItemDelegateForColumn(3, self.source_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.html_delegate.linkRightClicked.connect(self._show_related_menu)
        self.table.setItemDelegateForColumn(7, self.html_delegate)
        self.table.setItemDelegateForColumn(12, self.html_delegate)

        self.agenda_delegate = AgendaItemDelegate(self.table)
        self.table.setItemDelegateForColumn(10, self.agenda_delegate)

        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(scale_px(26))
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.setSectionResizeMode(9, QHeaderView.Fixed)
        self._apply_scaled_table_columns()
        header.sectionClicked.connect(self._on_header_section_clicked)
        header.sectionResized.connect(self._on_header_section_resized)
        get_font_scale_notifier().scale_changed.connect(self._on_font_scale_changed)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)

        if self.is_sa2:
            self.order_combo.blockSignals(True)
            self.order_combo.setCurrentIndex(0)  # 📋 Agenda Order
            self.order_combo.blockSignals(False)
            self.proxy.setSortMode("AGENDA")
            self.table.sortByColumn(10, Qt.AscendingOrder)

        layout.addWidget(self.table)

        # Bottom Status & Shared Comparison Cart Bar
        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(2, 2, 2, 0)
        bottom_bar.setSpacing(8)

        self._footer_status_widget = QWidget(self)
        self._footer_status_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        footer_status_layout = QHBoxLayout(self._footer_status_widget)
        footer_status_layout.setContentsMargins(0, 0, 0, 0)
        footer_status_layout.setSpacing(8)
        footer_status_layout.addWidget(self.count_lbl)
        sep_dot = QLabel("·")
        sep_dot.setStyleSheet(f"color: {C.TEXT_FAINT}; font-weight: bold;")
        footer_status_layout.addWidget(sep_dot)
        footer_status_layout.addWidget(self.last_mod_lbl)
        bottom_bar.addWidget(self._footer_status_widget)

        bottom_bar.addStretch()

        self.cart_frame = ComparisonCartWidget(self)
        self.btn_cart_launcher = self.cart_frame.create_launcher_button(self)
        bottom_bar.addWidget(self.btn_cart_launcher)
        bottom_bar.addWidget(self.cart_frame)
        layout.addLayout(bottom_bar)

        self._load_tracked_wi_tokens()
        self._sync_tracked_threads_to_table()
        self._refresh_comboboxes()
        self._refresh_email_counts()
        self._schedule_row_height_adjustment()

        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

    def _schedule_row_height_adjustment(self):
        """Coalesce table row-height recalculations onto a short single-shot timer."""
        if hasattr(self, "_row_height_timer") and self._row_height_timer is not None:
            self._row_height_timer.start()

    def _adjust_table_row_heights(self):
        """Expand visible TDocs rows dynamically for wrapped multi-line titles, sources, and remarks."""
        if not hasattr(self, "table") or not hasattr(self, "proxy"):
            return
        row_count = self.proxy.rowCount()
        min_h = scale_px(48)
        max_h = scale_px(124)

        vheader = self.table.verticalHeader()
        vheader.setMinimumSectionSize(min_h)
        vheader.setDefaultSectionSize(min_h)

        if row_count <= 0:
            return

        hheader = self.table.horizontalHeader()
        w_title = max(scale_px(80), hheader.sectionSize(2))
        w_source = max(scale_px(60), hheader.sectionSize(3))
        w_remarks = max(scale_px(140), hheader.sectionSize(7))
        w_related = max(scale_px(80), hheader.sectionSize(12))

        fm = self.table.fontMetrics()
        pad_v = scale_px(8)
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(self.table.font())

        has_emails = len(self.model._headers) > 13
        title_wrap_w = max(40, w_title - 10)
        source_wrap_w = max(40, w_source - 10)
        proxy = self.proxy

        for r in range(row_count):
            needed = min_h

            title_txt = str(proxy.data(proxy.index(r, 2), Qt.DisplayRole) or "")
            if title_txt and (len(title_txt) * fm.averageCharWidth() > title_wrap_w or "\n" in title_txt):
                h_t = fm.boundingRect(QRect(0, 0, title_wrap_w, 1000), Qt.TextWordWrap, title_txt).height() + pad_v
                if h_t > needed:
                    needed = h_t

            source_txt = str(proxy.data(proxy.index(r, 3), Qt.DisplayRole) or "")
            if source_txt and (len(source_txt) * fm.averageCharWidth() > source_wrap_w or "\n" in source_txt):
                h_s = fm.boundingRect(QRect(0, 0, source_wrap_w, 1000), Qt.TextWordWrap, source_txt).height() + pad_v
                if h_s > needed:
                    needed = h_s

            remarks_html = str(proxy.data(proxy.index(r, 7), Qt.DisplayRole) or "")
            if remarks_html and ("<" in remarks_html or len(remarks_html) > 35):
                doc.setTextWidth(w_remarks)
                doc.setHtml(remarks_html)
                h_r = int(doc.size().height()) + scale_px(6)
                if h_r > needed:
                    needed = h_r

            if has_emails:
                rel_html = str(proxy.data(proxy.index(r, 12), Qt.DisplayRole) or "")
                if rel_html and len(rel_html) > 25:
                    doc.setTextWidth(w_related)
                    doc.setHtml(rel_html)
                    h_rel = int(doc.size().height()) + scale_px(6)
                    if h_rel > needed:
                        needed = h_rel

            clamped = max(min_h, min(max_h, needed))
            if self.table.rowHeight(r) != clamped:
                self.table.setRowHeight(r, clamped)

    def _on_header_section_resized(self, _logical_index: int, _old_size: int, _new_size: int):
        """Track user-initiated column resizing so responsive relayout preserves manual widths."""
        if not getattr(self, "_programmatic_column_resize", False):
            self._user_resized_cols = True

    @staticmethod
    def compute_fitted_column_widths(viewport_width: int, col_count: int, scale_fn=None) -> dict:
        """
        Computes responsive TDocs table column widths that expand on wide monitors
        and proportionally shrink flexible columns on narrower windows so the
        rightmost Emails unread/flag column remains visible without horizontal clipping.
        """
        sc = scale_fn or scale_px
        base_widths = {
            0: sc(92),
            1: sc(86),
            2: sc(200),
            3: sc(112),
            4: sc(64),
            5: sc(70),
            6: sc(26),
            7: sc(220),
            8: sc(70),
            9: sc(26),
            10: sc(114),
            11: sc(86),
            12: sc(110),
        }
        if col_count > 13:
            base_widths[13] = sc(100)

        min_widths = {
            0: sc(74),
            1: sc(74),
            2: sc(100),
            3: sc(64),
            4: sc(52),
            5: sc(48),
            6: sc(26),
            7: sc(108),
            8: sc(54),
            9: sc(26),
            10: sc(68),
            11: sc(58),
            12: sc(58),
            13: sc(92),
        }

        widths = {col: w for col, w in base_widths.items() if col < col_count}
        total_base = sum(widths.values())
        if viewport_width > total_base > 0:
            surplus = viewport_width - total_base
            rem_extra = int(surplus * 0.52)
            title_extra = int(surplus * 0.34)
            source_extra = surplus - rem_extra - title_extra
            if 7 in widths:
                widths[7] += rem_extra
            if 2 in widths:
                widths[2] += title_extra
            if 3 in widths:
                widths[3] += source_extra
        elif 0 < viewport_width < total_base:
            deficit = total_base - viewport_width
            shrinkable = {
                c: max(0, widths[c] - min_widths.get(c, widths[c]))
                for c in widths
            }
            total_shrinkable = sum(shrinkable.values())
            if total_shrinkable > 0:
                target_cut = min(deficit, total_shrinkable)
                removed = 0
                for c, avail in shrinkable.items():
                    if avail > 0:
                        cut = int(target_cut * avail / total_shrinkable)
                        widths[c] -= cut
                        removed += cut
                # Distribute remaining rounding pixels from widest flexible columns
                leftover = target_cut - removed
                for c in sorted(shrinkable, key=lambda k: widths[k] - min_widths.get(k, widths[k]), reverse=True):
                    if leftover <= 0:
                        break
                    room = widths[c] - min_widths.get(c, widths[c])
                    if room > 0:
                        take = min(room, leftover)
                        widths[c] -= take
                        leftover -= take
        return widths

    def _schedule_column_fit(self):
        """Fits columns on the next event-loop tick so viewport().width() reflects post-layout & scrollbar geometry."""
        if getattr(self, "_user_resized_cols", False):
            return
        if not hasattr(self, "_col_fit_timer"):
            self._col_fit_timer = QTimer(self)
            self._col_fit_timer.setSingleShot(True)
            self._col_fit_timer.timeout.connect(self._apply_scaled_table_columns)
        self._col_fit_timer.start(15)

    def _apply_scaled_table_columns(self):
        """Scale and fit TDocsWindow table column widths to the active viewport."""
        if not hasattr(self, "table") or not hasattr(self, "model"):
            return
        header = self.table.horizontalHeader()
        col_count = len(self.model._headers)
        vp_w = self.table.viewport().width()
        if vp_w > 40:
            # Reserve 2px safety margin so Qt never triggers an unnecessary 1-2px horizontal scrollbar
            vp_w -= 2
            if not self.table.verticalScrollBar().isVisible() and self.proxy.rowCount() > 6:
                vp_w = max(200, vp_w - max(16, self.table.verticalScrollBar().sizeHint().width()))
        widths = self.compute_fitted_column_widths(vp_w, col_count)

        self._programmatic_column_resize = True
        try:
            header.setMinimumSectionSize(scale_px(26))
            for col, width in widths.items():
                if header.sectionSize(col) != width:
                    header.resizeSection(col, width)
        finally:
            self._programmatic_column_resize = False
        self._schedule_row_height_adjustment()

    def _on_font_scale_changed(self, _scale: float = 0.0):
        """Update filter dropdown caps, toolbar buttons, column widths, and row heights on font scale changes."""
        self.instant_fetch_input.setMinimumWidth(scale_px(120))
        self.instant_fetch_input.setMaximumWidth(scale_px(160))
        self.stats_cfg_btn.setMinimumWidth(scale_px(34))
        self.search_input.setMinimumWidth(scale_px(180))
        self.search_input.setMaximumWidth(scale_px(240))
        for btn in (
            self.btn_followed_wis,
            self.btn_starred_wis,
            self.btn_agreements_filter,
            self.btn_unread_emails,
            self.btn_flagged_emails,
        ):
            btn.setFixedSize(scale_px(28), scale_px(28))
        caps = (
            (self.company_combo, 165),
            (self.type_combo, 125),
            (self.ai_combo, 130),
            (self.sub_item_combo, 150),
            (self.wi_combo, 135),
            (self.status_combo, 155),
            (self.my_status_combo, 145),
            (self.order_combo, 165),
        )
        for combo, base_max in caps:
            combo.setMaximumWidth(scale_px(base_max))
            combo.updateGeometry()
        if not getattr(self, "_user_resized_cols", False):
            self._apply_scaled_table_columns()
        self._schedule_row_height_adjustment()

    def _setup_cache(self):
        agenda_dir = self.meeting_dir / "Agenda"
        if self.is_sa2:
            local_agenda = agenda_dir / "TdocsByAgenda.htm"
            if local_agenda.exists():
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(local_agenda))
                if agenda_data:
                    self.model.merge_agenda_data(agenda_data)
                    self._sync_tracked_threads_to_table()
                    self._refresh_comboboxes()

        if getattr(self, "is_ran", False):
            from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
            cached_chair_agenda, cached_agreements = ChairNotesAgreementsParser.load_cached_chair_notes_data(self.meeting_dir)
            if cached_chair_agenda:
                self.model.merge_agenda_data(cached_chair_agenda)
                self._refresh_comboboxes()
            if cached_agreements:
                self.chair_agreements = cached_agreements
                self.model.set_chair_agreements(cached_agreements)
            self._update_agreements_counter()
        else:
            if hasattr(self, "btn_meeting_agreements"):
                self.btn_meeting_agreements.setVisible(False)
            if hasattr(self, "btn_agreements_filter"):
                self.btn_agreements_filter.setVisible(False)

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

    def _on_tags_changed(self, tags):
        all_checked = (
            self.tags_combo.model().rowCount() > 0
            and self.tags_combo.model().item(0).checkState() == Qt.Checked
        )
        self.proxy.setTagFilters(tags, active=not all_checked)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_company_changed(self, companies):
        self.proxy.setCompanyFilters(companies)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_sub_item_changed(self, sub_items):
        self.proxy.setSubItemFilters(sub_items)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_order_changed(self, index):
        mode = self.order_combo.currentData() or "AGENDA"
        self.proxy.setSortMode(mode)
        if mode == "AGENDA":
            self.table.sortByColumn(10, Qt.AscendingOrder)
        elif mode == "TDOC":
            self.table.sortByColumn(1, Qt.AscendingOrder)
        elif mode == "COMPANY":
            self.table.sortByColumn(3, Qt.AscendingOrder)
        elif mode == "EMAILS":
            self.table.sortByColumn(13, Qt.DescendingOrder)
        elif mode == "DEFAULT":
            self.proxy.sort(-1)
        QTimer.singleShot(0, self._update_count_label)

    def _on_header_section_clicked(self, logical_index):
        if not hasattr(self, 'order_combo'):
            return
        self.order_combo.blockSignals(True)
        if logical_index == 10:  # Agenda Item
            idx = self.order_combo.findData("AGENDA")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("AGENDA")
        elif logical_index == 1:  # TDoc
            idx = self.order_combo.findData("TDOC")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("TDOC")
        elif logical_index == 3:  # Source / Company
            idx = self.order_combo.findData("COMPANY")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("COMPANY")
        elif logical_index == 13:  # Emails
            idx = self.order_combo.findData("EMAILS")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("EMAILS")
        else:
            self.proxy.setSortMode("OTHER")
        self.order_combo.blockSignals(False)

    def _toggle_followed_wis_filter(self):
        self._load_tracked_wi_tokens()
        self.proxy.setTrackedWIFilters(
            self.btn_followed_wis.isChecked(),
            self.btn_starred_wis.isChecked()
        )
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_agreements_filter(self):
        """Toggles table filtering to only TDocs with agreements or working assumptions."""
        checked = self.btn_agreements_filter.isChecked()
        self.proxy.setAgreementsFilter(checked)
        self._update_tracked_wi_button_styles()

    def _toggle_unread_emails_filter(self):
        """Toggles 1-click filtering to rows with unread emails."""
        checked = self.btn_unread_emails.isChecked()
        if checked and hasattr(self, "btn_flagged_emails"):
            self.btn_flagged_emails.blockSignals(True)
            self.btn_flagged_emails.setChecked(False)
            self.btn_flagged_emails.blockSignals(False)
        self.proxy.setEmailStateFilter("UNREAD" if checked else "ALL")
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_flagged_emails_filter(self):
        """Toggles 1-click filtering to rows with Outlook-flagged emails."""
        checked = self.btn_flagged_emails.isChecked()
        if checked and hasattr(self, "btn_unread_emails"):
            self.btn_unread_emails.blockSignals(True)
            self.btn_unread_emails.setChecked(False)
            self.btn_unread_emails.blockSignals(False)
        self.proxy.setEmailStateFilter("FLAGGED" if checked else "ALL")
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def locate_tdoc(self, tdoc_id: str):
        """Public API to scroll to and highlight a specific TDoc in the table."""
        self._scroll_to_tdoc(tdoc_id)

    def _update_agreements_counter(self):
        """Updates the ⚖️ Agreements toolbar button counter badge and active styling."""
        if not hasattr(self, "btn_meeting_agreements"):
            return
        if not getattr(self, "is_ran", False):
            self.btn_meeting_agreements.setVisible(False)
            return
        ag_list = getattr(self, "chair_agreements", [])
        non_recap = [a for a in ag_list if not a.get("is_previous_meeting_recap")]
        cnt = len(non_recap)
        if cnt == 0 and hasattr(self, "model") and hasattr(self.model, "_data"):
            cnt = sum(1 for r in self.model._data if r.get("_has_agreement"))

        if cnt > 0:
            self.btn_meeting_agreements.setText(f"⚖️ Agreements ({cnt})")
            self.btn_meeting_agreements.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.SUCCESS_SOFT};
                    color: {C.SUCCESS_DARK};
                    border: 1.5px solid {C.SUCCESS_BORDER_MEDIUM};
                    border-radius: 4px;
                    padding: 4px 10px;
                    font-size: 12px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {C.SUCCESS_SOFT_ALT};
                    border-color: #22C55E;
                }}
            """)
        else:
            self.btn_meeting_agreements.setText("⚖️ Agreements")
            self.btn_meeting_agreements.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)

    def _open_meeting_agreements_dialog(self, target_tdoc: str = None):
        """Opens the dedicated Meeting Agreements Viewer & Management Console."""
        if not getattr(self, "is_ran", False):
            QMessageBox.information(
                self,
                "RAN-Specific Feature",
                "Chairman's Notes agreements scanning and viewing are specific to RAN working groups (RAN1, RAN2, RAN3, etc.).\n\n"
                "In SA/CT working groups, contribution agreements are formally tracked via TDoc Status (agreed/approved) in the meeting table."
            )
            return
        from modules.meetings.ui.meeting_agreements_dialog import MeetingAgreementsDialog

        ag_list = getattr(self, "chair_agreements", [])
        if not ag_list:
            from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
            _, ag_list = ChairNotesAgreementsParser.load_cached_chair_notes_data(self.meeting_dir)
            if ag_list:
                self.chair_agreements = ag_list

        if not ag_list:
            # Fallback: construct lightweight agreement entries from model rows if remarks contain agreement tags
            synthetic_ags = []
            if hasattr(self, "model") and hasattr(self.model, "_data"):
                for row in self.model._data:
                    if row.get("_has_agreement"):
                        synthetic_ags.append({
                            "kind": row.get("_agreement_kind", "Agreement"),
                            "tdoc": row.get("TDoc", ""),
                            "ai": row.get("Agenda Item", ""),
                            "wi": row.get("Related WIs", ""),
                            "title": row.get("Title", ""),
                            "company": row.get("Source", ""),
                            "text": row.get("Secretary Remarks", ""),
                        })
            if synthetic_ags:
                ag_list = synthetic_ags
                self.chair_agreements = ag_list

        if not ag_list:
            reply = QMessageBox.question(
                self,
                "No Agreements Loaded",
                "No Chairman's Notes agreements have been scanned or cached for this meeting yet.\n\n"
                "Would you like to scan existing Chairman's Notes in Agenda/ChairNotes now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                self._scan_existing_chair_notes()
            return

        mtg_name = self.mtg_info.get("meeting_name", "") or f"{self.mtg_info.get('wg_name', '')} {self.mtg_info.get('meeting_number', '')}".strip()
        dlg = MeetingAgreementsDialog(
            agreements=ag_list,
            meeting_name=mtg_name,
            meeting_dir=self.meeting_dir,
            parent=self,
        )
        dlg.locate_tdoc_requested.connect(self.locate_tdoc)
        if target_tdoc:
            dlg.focus_tdoc(target_tdoc)
        dlg.exec_()

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
        if hasattr(self, "_update_facet_counts"):
            self._update_facet_counts()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _show_company_filter_context_menu(self, pos):
        sender = self.sender()
        menu = QMenu(self)
        act = menu.addAction("Include unconfirmed (?, [...]) carry-over sources")
        act.setCheckable(True)
        act.setChecked(getattr(self.proxy, "include_tentative_sources", False))
        act.setToolTip(
            "When checked, filtering by Company also matches revisions where the company is marked with ? or [...] as unconfirmed."
        )
        act.toggled.connect(self._on_toggle_include_tentative_sources)
        if sender and hasattr(sender, "mapToGlobal"):
            menu.exec_(sender.mapToGlobal(pos))
        else:
            menu.exec_(QCursor.pos())

    def _on_toggle_include_tentative_sources(self, checked: bool):
        if not hasattr(self, "settings"):
            self.settings = MeetingsSettings()
        self.settings.set_include_tentative_sources(checked)
        self.proxy.setIncludeTentativeSources(checked)
        self._refresh_comboboxes()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _update_tracked_wi_button_styles(self):
        """Update active/inactive visual styles of Followed and Starred WI quick-filter chips."""
        if not hasattr(self, "btn_followed_wis") or not hasattr(self, "btn_starred_wis"):
            return

        if self.btn_followed_wis.isChecked():
            self.btn_followed_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.INFO_SOFT_STRONG};
                    color: {C.INFO_DARK};
                    border: 1.5px solid {C.INFO_ACCENT};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.INFO_BORDER};
                    border-color: {C.ACCENT_ALT};
                }}
            """)
        else:
            self.btn_followed_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_SECONDARY};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    border-color: {C.TEXT_FAINT};
                }}
            """)

        if self.btn_starred_wis.isChecked():
            self.btn_starred_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.WARNING_SOFT_ALT};
                    color: {C.WARNING_DARK};
                    border: 1.5px solid {C.WARNING_ACCENT};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.WARNING_BORDER};
                    border-color: {C.WARNING_AMBER};
                }}
            """)
        else:
            self.btn_starred_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_SECONDARY};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    border-color: {C.TEXT_FAINT};
                }}
            """)

        if hasattr(self, "btn_agreements_filter"):
            if self.btn_agreements_filter.isChecked():
                self.btn_agreements_filter.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.SUCCESS_SOFT_ALT};
                        color: {C.SUCCESS_DARK};
                        border: 1.5px solid #22C55E;
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: {C.SUCCESS_BORDER};
                        border-color: {C.GREEN_ACCENT};
                    }}
                """)
            else:
                self.btn_agreements_filter.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
                """)

        if hasattr(self, "btn_unread_emails"):
            if self.btn_unread_emails.isChecked():
                self.btn_unread_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.INFO_SOFT_STRONG};
                        color: {C.INFO_DARK};
                        border: 1.5px solid {C.INFO_ACCENT};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: {C.INFO_BORDER};
                        border-color: {C.ACCENT_ALT};
                    }}
                """)
            else:
                self.btn_unread_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
                """)

        if hasattr(self, "btn_flagged_emails"):
            if self.btn_flagged_emails.isChecked():
                self.btn_flagged_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #FEE2E2;
                        color: #B91C1C;
                        border: 1.5px solid #DC2626;
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: #FECACA;
                        border-color: #B91C1C;
                    }}
                """)
            else:
                self.btn_flagged_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
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
        self._schedule_row_height_adjustment()

    def _update_filter_facets(self):
        """Update zero-intersection styling for every metadata filter."""
        if not hasattr(self, "proxy") or not hasattr(self, "model"):
            return

        facets = {
            "company": set(),
            "type": set(),
            "ai": set(),
            "sub_item": set(),
            "wi": set(),
            "status": set(),
            "my_status": set(),
            "tags": set(),
        }

        followed_count = 0
        starred_count = 0
        has_followed_tokens = bool(self.proxy.followed_wi_tokens)
        has_starred_tokens = bool(self.proxy.starred_wi_tokens)

        # Each facet is evaluated with all active constraints except itself.
        # This answers: "Would selecting this option yield rows given my other filters?"
        for row_idx, row in enumerate(self.model._data):
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="company"):
                comp_key = "_All_Sanitized_Companies" if getattr(self.proxy, "include_tentative_sources", False) else "_Sanitized_Companies"
                facets["company"].update(row.get(comp_key) or row.get("_Sanitized_Companies", ["Other"]))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="type"):
                facets["type"].add(str(row.get("Type", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="ai"):
                if row.get("_linked_ais"):
                    facets["ai"].update(row["_linked_ais"])
                else:
                    facets["ai"].add(str(row.get("Agenda Item", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="sub_item"):
                if row.get("_linked_sub_items"):
                    facets["sub_item"].update(row["_linked_sub_items"])
                else:
                    sub_val = str(row.get("Sub-Item", "") or "").strip()
                    if sub_val:
                        facets["sub_item"].add(sub_val)
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="wi"):
                facets["wi"].update(self.model.get_related_wis(row_idx))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="status"):
                facets["status"].add(str(row.get("TDoc Status", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="my_status"):
                facets["my_status"].add(str(row.get("My Status", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="tags"):
                row_tags = row.get("_tag_list") or []
                if row_tags:
                    facets["tags"].update(row_tags)
                else:
                    facets["tags"].add("(No Tags)")

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
        if hasattr(self, "sub_item_combo") and self.sub_item_combo.isVisible():
            self.sub_item_combo.setAvailability(facets["sub_item"])
        self.wi_combo.setAvailability(facets["wi"])
        self.status_combo.setAvailability(facets["status"])
        self.my_status_combo.setAvailability(facets["my_status"])
        if hasattr(self, "tags_combo") and self.tags_combo.isVisible():
            self.tags_combo.setAvailability(facets["tags"])

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
        all_ais = set()
        for r in self.model._data:
            if r.get("_linked_ais"):
                all_ais.update(sanitize(a) for a in r["_linked_ais"])
            all_ais.add(sanitize(r.get("Agenda Item", "")))
        unique_ais = sorted(list(all_ais), key=natural_sort_key)
        unique_wis = sorted({
            wi
            for row_idx in range(self.model.rowCount())
            for wi in self.model.get_related_wis(row_idx)
        }, key=lambda x: x.lower())
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))
        unique_my_statuses = sorted(list(set(sanitize(r.get("My Status", "")) for r in self.model._data)))

        unique_companies = set()
        comp_key = "_All_Sanitized_Companies" if getattr(self.proxy, "include_tentative_sources", False) else "_Sanitized_Companies"
        for r in self.model._data:
            unique_companies.update(r.get(comp_key) or r.get("_Sanitized_Companies", ["Other"]))
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

        unique_sub_items = []
        seen_subs = set()
        for r in self.model._data:
            for sub in (list(r.get("_linked_sub_items") or []) or [str(r.get("Sub-Item", "") or "").strip()]):
                sub = str(sub or "").strip()
                if sub and sub not in seen_subs:
                    seen_subs.add(sub)
                    unique_sub_items.append(sub)

        if hasattr(self, "sub_item_combo"):
            if unique_sub_items:
                self.sub_item_combo.setVisible(True)
                self.sub_item_combo.updateItems(unique_sub_items)
                self.proxy.setSubItemFilters(self.sub_item_combo.getCheckedItems())
            else:
                self.sub_item_combo.setVisible(False)
                self.proxy.setSubItemFilters([])

        unique_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        if hasattr(self, "tags_combo"):
            if unique_tags:
                tag_items = list(unique_tags) + ["(No Tags)"]
                self.tags_combo.setVisible(True)
                self.tags_combo.updateItems(tag_items)
                all_checked = (
                    self.tags_combo.model().rowCount() > 0
                    and self.tags_combo.model().item(0).checkState() == Qt.Checked
                )
                self.proxy.setTagFilters(self.tags_combo.getCheckedItems(), active=not all_checked)
            else:
                self.tags_combo.setVisible(False)
                self.proxy.setTagFilters([], active=False)

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

        if hasattr(self, "btn_agreements_filter"):
            self.btn_agreements_filter.blockSignals(True)
            self.btn_agreements_filter.setChecked(False)
            self.btn_agreements_filter.blockSignals(False)

        for btn_attr in ("btn_unread_emails", "btn_flagged_emails"):
            if hasattr(self, btn_attr):
                b = getattr(self, btn_attr)
                b.blockSignals(True)
                b.setChecked(False)
                b.blockSignals(False)

        self.proxy.setTrackedWIFilters(False, False)
        self.proxy.setAgreementsFilter(False)
        self.proxy.setEmailStateFilter("ALL")
        if hasattr(self, "row_mode_combo"):
            self.row_mode_combo.blockSignals(True)
            self.row_mode_combo.setCurrentIndex(0)
            self.row_mode_combo.blockSignals(False)
        self.proxy.setRowModeFilter("ALL")
        self._update_tracked_wi_button_styles()

        for combo in [self.company_combo, self.type_combo, self.ai_combo, self.wi_combo, self.status_combo, self.my_status_combo]:
            combo.blockSignals(True)
            combo.model().item(0).setCheckState(Qt.Checked)
            for i in range(1, combo.model().rowCount()):
                combo.model().item(i).setCheckState(Qt.Checked)
            combo.updateText()
            combo.blockSignals(False)

        if hasattr(self, 'sub_item_combo') and self.sub_item_combo.isVisible():
            self.sub_item_combo.blockSignals(True)
            if self.sub_item_combo.model() and self.sub_item_combo.model().rowCount() > 0:
                self.sub_item_combo.model().item(0).setCheckState(Qt.Checked)
                for i in range(1, self.sub_item_combo.model().rowCount()):
                    self.sub_item_combo.model().item(i).setCheckState(Qt.Checked)
                self.sub_item_combo.updateText()
            self.sub_item_combo.blockSignals(False)
            self.proxy.setSubItemFilters(self.sub_item_combo.getCheckedItems())

        if hasattr(self, 'tags_combo') and self.tags_combo.isVisible():
            self.tags_combo.blockSignals(True)
            if self.tags_combo.model() and self.tags_combo.model().rowCount() > 0:
                self.tags_combo.model().item(0).setCheckState(Qt.Checked)
                for i in range(1, self.tags_combo.model().rowCount()):
                    self.tags_combo.model().item(i).setCheckState(Qt.Checked)
                self.tags_combo.updateText()
            self.tags_combo.blockSignals(False)
            self.proxy.setTagFilters(self.tags_combo.getCheckedItems(), active=False)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setWIFilters(self.wi_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())
        self._update_filter_facets()

        QTimer.singleShot(0, self._update_count_label)

    def _handle_tdoc_action(self, base_tdoc: str):
        if str(base_tdoc or "").strip().upper().startswith("THR-"):
            self._open_tracked_thread_dialog(str(base_tdoc).strip().upper())
            return
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
            launch_cli_callback=self._launch_cli_for_tdocs,
            send_tdoc_email_callback=self._send_tdoc_via_email,
        )

    @staticmethod
    def _extract_base_tdoc(tdoc_id: str) -> str:
        clean = str(tdoc_id or "").strip().upper()
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', clean, re.IGNORECASE)
        return match.group(1).upper() if match else clean

    def _build_email_subject(self, tdoc_id: str) -> str:
        clean_tdoc = str(tdoc_id or "").strip()
        row_data = next(
            (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == clean_tdoc.upper()),
            None,
        )

        if not row_data:
            base_tdoc = self._extract_base_tdoc(clean_tdoc)
            row_data = next(
                (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc),
                {},
            )

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
        return f"{tag_str}{clean_tdoc} {title}".strip()

    def _compose_email_draft(self, tdoc_id: str):
        subject = self._build_email_subject(tdoc_id)
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

    def _send_tdoc_via_email(self, target_tdoc: str):
        """Drafts an email with the TDoc subject and attaches the TDoc document(s), downloading first if needed."""
        clean_target = str(target_tdoc or "").strip()
        if not clean_target:
            return

        local_files = TDocFileHandler.find_local_tdoc_files(self.meeting_dir, clean_target)
        if local_files:
            self._open_email_draft_with_attachments(clean_target, local_files)
            return

        base_tdoc = self._extract_base_tdoc(clean_target)
        if base_tdoc in self.model.loading_tdocs:
            return
        self._trigger_download_thread(
            base_tdoc,
            clean_target,
            send_email_on_finish=True,
        )

    def _open_email_draft_with_attachments(self, target_tdoc: str, attachments: list):
        subject = self._build_email_subject(target_tdoc)
        valid_paths = [Path(p) for p in (attachments or []) if p and Path(p).is_file()]
        if not valid_paths:
            QMessageBox.warning(
                self,
                "Attachment Missing",
                f"Could not locate a local document file to attach for {target_tdoc}.",
            )
            return

        ok, err = OutlookClient.create_email_draft(subject, valid_paths)
        if ok:
            logging.info(
                f"📧 [Send TDoc] Opened Outlook draft for '{target_tdoc}' with {len(valid_paths)} attachment(s)."
            )
            return

        # Graceful fallback when Classic Outlook COM is unavailable
        QApplication.clipboard().setText(subject)
        encoded_subject = urllib.parse.quote(subject, safe='')
        mailto_url = f"mailto:?subject={encoded_subject}"
        try:
            if hasattr(os, 'startfile'):
                os.startfile(mailto_url)
            else:
                webbrowser.open(mailto_url)
        except Exception:
            pass

        try:
            _open_folder(valid_paths[0].parent)
        except Exception:
            pass

        QMessageBox.information(
            self,
            "Attach TDoc Manually",
            f"Could not attach files automatically via Outlook COM ({err}).\n\n"
            f"The email subject has been copied to your clipboard and the local folder for {target_tdoc} was opened so you can drag-and-drop the file(s).",
        )

    def _open_interactive_notes_for_row(self, row_data: dict):
        if not row_data:
            return
        tdoc_id = str(row_data.get("TDoc", "") or "").strip()
        if not tdoc_id:
            return
        all_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        InteractiveNotesDialog(
            self,
            tdoc_id,
            row_data,
            self._save_user_data,
            all_existing_tags=all_tags,
        ).exec_()

    def _show_cell_popup(self, index):
        if not index.isValid():
            return
        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        col_name = self.model._headers[index.column()]
        if row_data.get("_is_tracked_thread"):
            if col_name in ["My Notes", "My Status"]:
                self._open_interactive_notes_for_row(row_data)
            else:
                self._open_tracked_thread_dialog(row_data.get("_thread_id") or row_data.get("TDoc", ""))
            return

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

        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        tdoc_id = row_data.get("TDoc", "")

        if col_name == "Secretary Remarks":
            if row_data.get("_has_agreement"):
                self._open_meeting_agreements_dialog(target_tdoc=tdoc_id)
                return
            remarks = str(row_data.get("Secretary Remarks", "") or "")
            if remarks.strip():
                ReadOnlyViewerDialog(self, f"📄 Viewing: Secretary Remarks ({tdoc_id})", remarks).exec_()
            else:
                self._open_interactive_notes_for_row(row_data)
            return

        if col_name == "TDoc Status" and row_data.get("_has_agreement"):
            self._open_meeting_agreements_dialog(target_tdoc=tdoc_id)
            return

        if col_name not in [
            "Title", "Source", "Abstract", "My Notes", "My Status"
        ]:
            return

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
            self._open_interactive_notes_for_row(row_data)

    def _save_user_data(self, tdoc_id: str, status: str, notes: str, tags: str = None):
        vbar = self.table.verticalScrollBar() if hasattr(self, "table") else None
        prev_scroll = vbar.value() if vbar else 0
        self.db.upsert(tdoc_id, status, notes, tags=tags)
        if str(tdoc_id or "").strip().upper().startswith("THR-") and hasattr(self, "tracked_threads_db"):
            tag_list = None
            if tags is not None:
                tag_list = [t.strip() for t in str(tags).replace(";", ",").split(",") if t.strip()]
            self.tracked_threads_db.update_thread_annotations(
                str(tdoc_id).strip(),
                user_status=status,
                user_notes=notes,
                user_tags=tag_list,
            )
            open_dlg = getattr(self, "_tracked_thread_dialogs", {}).get(str(tdoc_id).strip())
            if open_dlg is not None:
                try:
                    open_dlg._load_data()
                except Exception:
                    pass
        self.model.user_data = self.db.get_all()
        self.model.apply_user_data_refresh()
        self._refresh_comboboxes()
        self._adjust_table_row_heights()
        if vbar is not None:
            vbar.setValue(prev_scroll)

    def _export_reports(self):
        self.export_thread = MarkdownExporterThread(
            self.meeting_dir,
            self.model._data,
            self.docs_ftp_url,
            self.mtg_info,
            agenda_map=getattr(self, "agenda_map", None),
        )
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

    def _find_proxy_row(self, base_tdoc: str):
        for row in range(self.proxy.rowCount()):
            if self.proxy.data(self.proxy.index(row, 1), Qt.UserRole) == base_tdoc:
                return row
        return None

    def _focus_and_highlight_proxy_row(self, row: int, base_tdoc: str):
        idx = self.proxy.index(row, 1)
        if not idx.isValid():
            return
        self.model.set_highlighted_tdoc(base_tdoc)
        if hasattr(self, "_highlight_clear_timer") and self._highlight_clear_timer is not None:
            self._highlight_clear_timer.start()
        self.table.selectRow(row)
        self.table.setCurrentIndex(idx)
        self.table.scrollTo(idx, QTableView.PositionAtCenter)

    def _reveal_tdoc_in_table(self, target_tdoc: str):
        """
        Syncs the table view after a successful Quick Open:
        - If the current filter already includes the TDoc, scroll to and highlight it.
        - If hidden by active filters, clear filters and filter to the TDoc, then scroll & highlight.
        """
        base_tdoc = self._extract_base_tdoc(target_tdoc)
        if not base_tdoc or base_tdoc not in self.model.valid_tdocs:
            return

        row = self._find_proxy_row(base_tdoc)
        if row is not None:
            self._focus_and_highlight_proxy_row(row, base_tdoc)
            return

        self._clear_all_filters()
        self.search_timer.stop()
        self.search_input.blockSignals(True)
        self.search_input.setText(base_tdoc)
        self.search_input.blockSignals(False)
        self.proxy.setGlobalFilter(base_tdoc)
        self._update_filter_facets()
        self._update_count_label()
        self._adjust_table_row_heights()

        row = self._find_proxy_row(base_tdoc)
        if row is not None:
            self._focus_and_highlight_proxy_row(row, base_tdoc)

    def _scroll_to_tdoc(self, target_tdoc: str):
        base_tdoc = self._extract_base_tdoc(target_tdoc)
        if base_tdoc in self.model.valid_tdocs:
            row = self._find_proxy_row(base_tdoc)
            if row is not None:
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
        if getattr(thread, "sync_table_on_success", False):
            self._reveal_tdoc_in_table(getattr(thread, "target_filename", "") or tdoc)
        if getattr(thread, "send_email_on_finish", False):
            target_name = getattr(thread, "target_filename", "") or tdoc
            attachments = list(getattr(thread, "extracted_doc_paths", []) or [])
            if not attachments:
                zip_candidate = self.meeting_dir / tdoc / f"{target_name}.zip"
                if zip_candidate.is_file():
                    attachments = [zip_candidate]
            self._open_email_draft_with_attachments(target_name, attachments)
            return
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
                self._setup_cache()
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

    def _cache_all_tdocs(self):
        start_window_tdocs_caching(self)

    def _open_excel(self):
        _open_folder(Path(self.filepath))

    def _open_session_plan_window(self):
        """Launches the dedicated Meeting Session Plan viewer."""
        if not hasattr(self, "_session_plan_window") or self._session_plan_window is None:
            from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow
            followed = getattr(self.proxy, "followed_wi_tokens", set()) if hasattr(self, "proxy") else set()
            starred = getattr(self.proxy, "starred_wi_tokens", set()) if hasattr(self, "proxy") else set()
            tracked_tokens = set(followed) | set(starred)
            self._session_plan_window = SessionPlanWindow(
                meeting_dir=self.meeting_dir,
                mtg_info=self.mtg_info,
                agenda_map=self.agenda_map,
                tracked_tokens=tracked_tokens,
            )
        self._session_plan_window.show()
        self._session_plan_window.raise_()
        self._session_plan_window.activateWindow()

    def _load_tdoc_as_session_plan(self, tdoc_id: str):
        """Fetches the target TDoc, stages it to Agenda/SessionPlans, and launches SessionPlanWindow."""
        from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
        from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
        from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
        from modules.meetings.core.url_router import URLRouter
        from core.network.network_state import NetworkState

        wg_name = self.mtg_info.get("wg_name", "").upper()
        candidate_urls = []
        if NetworkState.get_instance().is_local_active():
            candidate_urls.append(URLRouter._get_local_server_base(wg_name) + "/Docs")
        if self.docs_ftp_url:
            candidate_urls.append(self.docs_ftp_url)

        tdoc_dir = self.meeting_dir / tdoc_id
        extracted_docs = []
        try:
            for url in candidate_urls:
                extracted_docs = TDocFileHandler.download_and_extract_tdoc(tdoc_id, url, tdoc_dir, timeout=10)
                if extracted_docs:
                    break
        except Exception as e:
            logging.warning(f"Failed to fetch {tdoc_id}: {e}")

        if not extracted_docs and tdoc_dir.exists():
            extracted_docs = list(tdoc_dir.glob("*.docx")) + list(tdoc_dir.glob("*.doc"))

        if not extracted_docs:
            QMessageBox.warning(self, "Download Failed", f"Could not acquire document files for {tdoc_id}.")
            return

        doc_file = extracted_docs[0]
        staged_file = SessionPlanParser.stage_session_plan_file(doc_file, self.meeting_dir, tdoc_id)
        db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")

        if staged_file.suffix.lower() == ".docx":
            snap = SessionPlanParser.parse_docx(staged_file, meeting_number=str(self.mtg_info.get("meeting_number", "")))
        else:
            snap = SessionPlanParser.parse_text_or_markdown(
                staged_file.read_text(encoding="utf-8", errors="replace"),
                meeting_number=str(self.mtg_info.get("meeting_number", "")),
                cached_file_path=str(staged_file),
            )

        db.save_snapshot(snap)
        self._open_session_plan_window()
        if hasattr(self, "_session_plan_window") and self._session_plan_window:
            self._session_plan_window._load_data()
        QMessageBox.information(self, "Session Plan Loaded", f"Successfully loaded {tdoc_id} into Session Plan viewer!")

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

        if not hasattr(self, "_active_cr_dialogs"):
            self._active_cr_dialogs = []
        dialog = CrPcrDialog(
            spec_num=spec_num,
            version=version,
            meeting_info=mtg_info,
            tdoc_info=tdoc_info,
            parent=self,
        )
        self._active_cr_dialogs.append(dialog)
        dialog.destroyed.connect(lambda: self._active_cr_dialogs.remove(dialog) if dialog in self._active_cr_dialogs else None)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

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

    def _open_pre_meeting_export_dialog(self):
        open_pre_meeting_export_dialog(self)

    def _open_post_meeting_export_dialog(self):
        open_post_meeting_export_dialog(self)

    def _export_dossier(self, target_rows=None):
        from modules.meetings.ui.dialogs.tdocs_export_controller import open_dossier_export_dialog
        open_dossier_export_dialog(self, target_rows)

    def _launch_cli_for_tdocs(self, rows_data, agent_id: str, task_mode: str = "interactive", custom_question: str = ""):
        launch_cli_for_tdocs(self, rows_data, agent_id, task_mode=task_mode, custom_question=custom_question)

    def _populate_export_cli_submenu(self):
        rows = get_package_export_rows(self)
        populate_cli_agent_menu(
            self.cli_export_submenu,
            on_launch_callback=lambda aid, mode, r=rows: self._launch_cli_for_tdocs(r, aid, mode),
            has_multiple_or_revs=(len(rows) > 1),
            multi_count=len(rows),
            parent=self,
        )

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
            self.routing_indicator.setStyleSheet(f"color: {C.WG_RAN_TEXT}; background-color: #D1FAE5; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the high-speed local 10.10.10.10 network.")
        elif is_active:
            self.routing_indicator.setText("🔵 SYNC Folder")
            self.routing_indicator.setStyleSheet(f"color: {C.INFO_DARK}; background-color: {C.INFO_SOFT_STRONG}; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the active meeting SYNC folder.")
        else:
            self.routing_indicator.setText("🌐 Standard Web")
            self.routing_indicator.setStyleSheet(f"color: {C.TEXT_BODY}; background-color: {C.BG_CANVAS}; font-weight: bold; padding: 2px 8px; border-radius: 4px;")
            self.routing_indicator.setToolTip("Downloads are routed through the standard 3GPP web archive.")
        self._update_venue_clock()

    def _update_venue_clock(self):
        if not hasattr(self, "venue_clock_btn") or not hasattr(self, "_clock_service"):
            return
        try:
            state = self._clock_service.resolve_single_meeting(
                self.mtg_info, meeting_dir=self.meeting_dir
            )
            if state.venue_dt is not None and state.tz_offset_minutes is not None:
                is_elec = bool(self.mtg_info.get("is_electronic", 0)) or state.location == "Online (UTC)"
                if is_elec:
                    self.venue_clock_btn.setText(f"🌐 {state.venue_dt:%H:%M} UTC (eMeeting)")
                else:
                    off_str = format_utc_offset(state.tz_offset_minutes)
                    if off_str.endswith(":00"):
                        off_str = off_str[:-3]
                    day_tag = f" ({state.day_diff_vs_laptop:+d}d)" if state.day_diff_vs_laptop != 0 else ""
                    self.venue_clock_btn.setText(f"📍 {state.location} {state.venue_dt:%H:%M}{day_tag} ({off_str})")
                self.venue_clock_btn.setToolTip(
                    f"Meeting Venue: {state.location}\n"
                    f"Venue Time: {state.venue_dt:%a %d %b %Y, %H:%M:%S}\n"
                    f"Timezone: {state.tz_label}\n"
                    f"Source: {state.source}\n\n"
                    f"Click to manage venue timezone, sync from web, or set manual override"
                )
            else:
                self.venue_clock_btn.setText(f"📍 {state.location or 'Venue'} (--:--)")
                self.venue_clock_btn.setToolTip(
                    "Venue timezone unresolved. Click to configure or sync from web."
                )
        except Exception as e:
            logging.debug(f"[TDocsWindow] Venue clock update failed: {e}")

    def _open_venue_clock_dialog(self):
        from PyQt5.QtWidgets import QApplication
        from core.ui.world_clock_dialog import WorldClockDialog

        dlg = WorldClockDialog(
            self._clock_service,
            focus_meeting=self.mtg_info,
            focus_meeting_dir=self.meeting_dir,
            parent=self,
        )

        def _sync_all_clocks():
            self._update_venue_clock()
            for top in QApplication.topLevelWidgets():
                sb = getattr(top, "status_bar", None)
                if sb is not None and hasattr(sb, "refresh_world_clock"):
                    sb.refresh_world_clock()

        dlg.settings_changed.connect(_sync_all_clocks)
        dlg.exec_()
        _sync_all_clocks()

    def _on_instant_fetch(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._trigger_download_thread(
            base_tdoc,
            target_filename,
            legacy_url=None,
            is_silent_compare=False,
            sync_table_on_success=True,
        )

    def _on_instant_compare_active_word(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._compare_tdoc_with_active_word(base_tdoc, target_filename, sync_table_on_success=True)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.setGeometry(self.rect())
        if not getattr(self, "_user_resized_cols", False):
            self._apply_scaled_table_columns()
            self._schedule_column_fit()
        self._schedule_row_height_adjustment()

    @staticmethod
    def _has_outlook_drag_payload(mime_data) -> bool:
        if not mime_data:
            return False
        for fmt in mime_data.formats():
            if any(k in fmt for k in ("FileGroupDescriptor", "RenPrivateMessages", "RenPrivateItem")):
                return True
        return False

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        accepted = False
        if mime and mime.hasUrls():
            for url in mime.urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.zip', '.htm', '.html', '.msg']:
                    accepted = True
                    break
        if not accepted and self._has_outlook_drag_payload(mime):
            accepted = True
        if accepted:
            self.drop_overlay.setGeometry(self.rect())
            self.drop_overlay.show()
            self.drop_overlay.raise_()
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        mime = event.mimeData()
        if mime and mime.hasUrls():
            for url in mime.urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.zip', '.htm', '.html', '.msg']:
                    event.acceptProposedAction()
                    return
        if self._has_outlook_drag_payload(mime):
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

        mime = event.mimeData()
        valid_files = []
        msg_files = []
        if mime and mime.hasUrls():
            for url in mime.urls():
                file_path = Path(url.toLocalFile())
                ext = file_path.suffix.lower()
                if ext == '.msg':
                    msg_files.append(file_path)
                elif ext in ['.docx', '.doc', '.zip', '.htm', '.html']:
                    valid_files.append(file_path)

        outlook_dropped = False
        if not msg_files and not valid_files and self._has_outlook_drag_payload(mime):
            event.acceptProposedAction()
            outlook_dropped = True
            self._process_dropped_outlook_items(mime)

        if msg_files:
            event.acceptProposedAction()
            self._process_dropped_msg_files(msg_files)

        if valid_files:
            event.acceptProposedAction()
            # Smart routing: check if dropped file is a session plan
            session_plan_files = []
            agenda_files = []
            for vf in valid_files:
                name_lower = vf.name.lower()
                if "session" in name_lower or "time_allocation" in name_lower or "time allocation" in name_lower:
                    session_plan_files.append(vf)
                else:
                    agenda_files.append(vf)

            if session_plan_files:
                for spf in session_plan_files:
                    from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
                    from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
                    staged = SessionPlanParser.stage_session_plan_file(spf, self.meeting_dir)
                    db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")
                    if staged.suffix.lower() == ".docx":
                        snap = SessionPlanParser.parse_docx(staged, meeting_number=str(self.mtg_info.get("meeting_number", "")))
                    else:
                        snap = SessionPlanParser.parse_text_or_markdown(
                            staged.read_text(encoding="utf-8", errors="replace"),
                            meeting_number=str(self.mtg_info.get("meeting_number", "")),
                            cached_file_path=str(staged),
                        )
                    db.save_snapshot(snap)
                self._open_session_plan_window()
                if hasattr(self, "_session_plan_window") and self._session_plan_window:
                    self._session_plan_window._load_data()

            if agenda_files:
                self._process_imported_agenda_files(agenda_files)
        elif not msg_files and not outlook_dropped:
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

    def _scan_existing_chair_notes(self):
        scan_existing_chair_notes(self)

    def _on_chair_notes_finished(self, success: bool, count: int, files: list, msg: str):
        on_chair_notes_finished(self, success, count, files, msg)

    def _open_chair_notes_folder(self):
        open_chair_notes_folder(self)

    def _open_doc_in_word(self, doc_path):
        import os
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(doc_path))
            else:
                import webbrowser
                webbrowser.open(f"file:///{doc_path}")
        except Exception as e:
            QMessageBox.warning(self, "Open Failed", f"Could not open document:\n{e}")

    def _open_agreement_tracker(self):
        from modules.meetings.ui.agreement_tracker_dialog import AgreementTrackerDialog
        from modules.meetings.core.meetings_db import MeetingsDatabase
        from PyQt5.QtCore import QDate

        if not hasattr(self, "_agreement_tracker_dlg") or self._agreement_tracker_dlg is None:
            db_path = get_project_root() / "db" / "3gpp_data.db"
            self._agreement_tracker_dlg = AgreementTrackerDialog(MeetingsDatabase(db_path), self)
            self._agreement_tracker_dlg.open_meeting_requested.connect(
                lambda wg, num: self.global_action_requested.emit(f"{wg}#{num}", "open_meeting")
            )

        wg = str(self.mtg_info.get("wg_name", "")).strip()
        if wg:
            self._agreement_tracker_dlg.set_active_wgs([wg])
        s_date = QDate.fromString(str(self.mtg_info.get("start_date", "")), "yyyy-MM-dd")
        e_date = QDate.fromString(str(self.mtg_info.get("end_date", "")), "yyyy-MM-dd")
        if s_date.isValid() and e_date.isValid():
            self._agreement_tracker_dlg.set_date_range(s_date.addDays(-7), e_date.addDays(7))

        checked_wis = self.wi_combo.getCheckedItems() if hasattr(self, "wi_combo") else []
        all_wis = (
            self.wi_combo.model().rowCount() - 1
            if hasattr(self, "wi_combo") and self.wi_combo.model()
            else 0
        )
        if checked_wis and len(checked_wis) < all_wis:
            self._agreement_tracker_dlg._clear_wi_tokens()
            self._agreement_tracker_dlg.add_wi_tokens(checked_wis)

        self._agreement_tracker_dlg.show()
        self._agreement_tracker_dlg.raise_()
        self._agreement_tracker_dlg.activateWindow()

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


    def _get_meeting_id_label(self) -> str:
        wg = str(self.mtg_info.get("wg_name", "")).strip().upper()
        num = str(self.mtg_info.get("meeting_number", "")).strip()
        return f"{wg}#{num}" if (wg and num) else (wg or num or "MEETING")

    def _expand_tdoc_family_with_merges(self, tdocs: list) -> list:
        """Expands TDoc IDs across revision chains and merged-to/replaced-by contributions."""
        if not tdocs or not hasattr(self, "model"):
            return list(tdocs or [])
        tdoc_map = getattr(self.model, "_tdoc_map", {}) or {}
        seen = []
        seen_set = set()
        queue = [str(t).strip().upper() for t in tdocs if str(t).strip()]
        tdoc_token_re = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{5,8}(?:r\d{1,2}[a-z]?)?\b", re.IGNORECASE)

        while queue:
            cur = queue.pop(0)
            if cur in seen_set:
                continue
            seen_set.add(cur)
            seen.append(cur)

            # Revision family from model
            for fam in self.model.get_family_tdocs(cur):
                f_up = str(fam).strip().upper()
                if f_up and f_up not in seen_set:
                    queue.append(f_up)

            # Check merge / replacement hints in row metadata
            row = tdoc_map.get(cur)
            if row:
                for col in ("Merged to", "Merged into", "Replaced by", "Revised to", "Secretary Remarks", "TDoc Status"):
                    val = str(row.get(col, "") or "")
                    if not val:
                        continue
                    if col in ("Secretary Remarks", "TDoc Status") and not any(k in val.lower() for k in ("merg", "revis", "replac", "supersed")):
                        continue
                    for m in tdoc_token_re.findall(val):
                        m_up = m.upper()
                        if m_up not in seen_set:
                            queue.append(m_up)
        return seen

    def _build_thread_autocomplete_suggestions(self):
        """Builds AI, WI, TDoc, Sub-Topic lists and AI -> Sub-Topics mapping for TrackedThreadDialog."""
        ai_items = []
        seen_ais = set()
        for ai_num, item in sorted((self.agenda_map or {}).items(), key=lambda kv: natural_sort_key(kv[0])):
            desc = getattr(item, "description", "") or ""
            label = f"{ai_num} — {desc}" if desc else str(ai_num)
            if ai_num not in seen_ais:
                seen_ais.add(ai_num)
                ai_items.append(label)

        wi_set = set()
        tdoc_items = []
        seen_tdocs = set()
        sub_items = []
        seen_subs = set()
        ai_to_sub_items = {}
        for row in (getattr(self.model, "_data", None) or []):
            if row.get("_is_tracked_thread"):
                continue
            ai_val = str(row.get("Agenda Item", "") or "").strip()
            valid_ai = bool(ai_val and ai_val not in ("N/A", "-", "None"))
            if valid_ai and ai_val not in seen_ais:
                seen_ais.add(ai_val)
                ai_desc = str(row.get("Agenda item description", "") or "").strip()
                ai_items.append(f"{ai_val} — {ai_desc}" if ai_desc else ai_val)
            sub_val = str(row.get("Sub-Item", "") or "").strip()
            if sub_val:
                if valid_ai:
                    bucket = ai_to_sub_items.setdefault(ai_val, [])
                    if sub_val not in bucket:
                        bucket.append(sub_val)
                if sub_val not in seen_subs:
                    seen_subs.add(sub_val)
                    sub_items.append(f"{sub_val} — AI {ai_val}" if valid_ai else sub_val)
            for wi in self.model._split_related_wis(row.get("Related WIs", "")):
                if wi and wi.lower() not in ("(no wi)", "none", "nan", "-"):
                    wi_set.add(wi)
            tid = str(row.get("TDoc", "") or "").strip().upper()
            if tid and tid not in seen_tdocs:
                seen_tdocs.add(tid)
                title = str(row.get("Title", "") or "").strip()
                tdoc_items.append(f"{tid} — {title[:70]}" if title else tid)
        return ai_items, sorted(wi_set), tdoc_items, sub_items, ai_to_sub_items

    def _sync_tracked_threads_to_table(self):
        if not hasattr(self, "tracked_threads_db") or not hasattr(self, "model"):
            return
        try:
            threads = self.tracked_threads_db.get_threads_for_meeting(self._get_meeting_id_label())
            tracked_unread_ids = set()
            for t in threads:
                cur_tdocs = t.get("linked_tdocs") or []
                if cur_tdocs:
                    expanded = self._expand_tdoc_family_with_merges(cur_tdocs)
                    if [x.upper() for x in expanded] != [x.upper() for x in cur_tdocs]:
                        t["linked_tdocs"] = expanded
                        self.tracked_threads_db.update_thread_links(t["thread_id"], linked_tdocs=expanded)
                msgs = self.tracked_threads_db.get_thread_messages(t["thread_id"])
                for m in msgs:
                    if int(m.get("is_read", 0) or 0) == 0 and m.get("id"):
                        tracked_unread_ids.add(str(m["id"]))
                comps = list(dict.fromkeys(
                    str(m.get("company") or "").strip()
                    for m in msgs
                    if str(m.get("company") or "").strip() and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
                ))
                if comps:
                    t["participants_companies"] = ", ".join(comps)
            self.model.sync_tracked_thread_rows(threads)
            tdoc_unread_ids = set()
            if hasattr(self, "general_email_db") and self.general_email_db is not None:
                for stats in (self.general_email_db.get_email_counts_per_tdoc() or {}).values():
                    tdoc_unread_ids.update(str(eid) for eid in (stats.get("unread_ids") or ()) if eid)
            total_unread = len(tracked_unread_ids | tdoc_unread_ids)
            if hasattr(self, "email_btn"):
                if total_unread > 0:
                    self.email_btn.setText(f"📧 Emails ({total_unread} new) ▾")
                    self.email_btn.setToolTip(
                        f"Email Discussions & Tracked Threads\n"
                        f"{total_unread} unread email(s) total "
                        f"({len(tracked_unread_ids)} in Tracked Threads, {len(tdoc_unread_ids)} in TDoc Emails)."
                    )
                else:
                    self.email_btn.setText("📧 Emails ▾")
                    self.email_btn.setToolTip("Email Discussions & Tracked Threads (no unread emails).")
        except Exception as e:
            logging.error(f"Error syncing tracked email threads to table: {e}")

    def _open_tdoc_document_directly(self, tdoc_id: str):
        """
        Downloads/opens a TDoc contribution (or revision) directly without stealing focus from email dialogs.
        Automatically routes external / cross-WG TDocs (e.g. R3-263595, older SA2 TDocs) via
        global_action_requested -> MeetingsTab instead of failing against this meeting's FTP folder.
        """
        td = str(tdoc_id or "").strip().upper()
        if not td:
            return
        match = re.search(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", td, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else td
        valid_tdocs = getattr(self.model, "valid_tdocs", set()) if hasattr(self, "model") else set()
        in_current_meeting = (base_tdoc in valid_tdocs) or (td in valid_tdocs)

        if not in_current_meeting:
            try:
                from modules.meetings.core.meetings_db import MeetingsDatabase
                db_path = get_project_root() / "db" / "3gpp_data.db"
                if db_path.is_file():
                    ext_mtg = MeetingsDatabase(db_path).find_meeting_by_tdoc(base_tdoc)
                    if ext_mtg and ext_mtg.get("mtg_id") != self.mtg_info.get("mtg_id"):
                        self.global_action_requested.emit(td, "open_doc")
                        return
            except Exception as exc:
                logging.debug(f"Global TDoc lookup skipped for {td}: {exc}")

            cur_prefix = str(self._get_tdoc_prefix() or "").upper()
            if cur_prefix and not base_tdoc.startswith(cur_prefix):
                QMessageBox.information(
                    self,
                    "External TDoc",
                    f"{td} does not belong to {self._get_meeting_id_label()} and its meeting "
                    f"could not be resolved in the local Meetings Database.\n\n"
                    f"Sync that Working Group / meeting in the Meetings tab to open it directly.",
                )
                return

        is_rev = bool(re.search(r"(?:r|rev)\d{1,2}[a-zA-Z]?$", td, re.IGNORECASE))
        target_url = (
            getattr(self, "revisions_url", "")
            if is_rev and getattr(self, "revisions_url", "")
            else getattr(self, "docs_ftp_url", "")
        )
        if not target_url:
            self._scroll_to_tdoc(td)
            return
        self._trigger_download_thread(base_tdoc, td, target_url, is_silent_compare=False)

    def _open_tracked_thread_dialog(self, thread_id: str):
        tid = str(thread_id or "").strip()
        if not tid:
            return
        if tid in self._tracked_thread_dialogs:
            dlg = self._tracked_thread_dialogs[tid]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return
        ai_sugg, wi_sugg, tdoc_sugg, sub_sugg, ai_to_sub_items = self._build_thread_autocomplete_suggestions()
        all_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        dlg = TrackedThreadDialog(
            self.tracked_threads_db,
            tid,
            parent=None,
            ai_suggestions=ai_sugg,
            wi_suggestions=wi_sugg,
            tdoc_suggestions=tdoc_sugg,
            tdoc_family_expander=self._expand_tdoc_family_with_merges,
            sub_item_suggestions=sub_sugg,
            all_existing_tags=all_tags,
            annotations_save_callback=self._save_user_data,
            ai_to_sub_items=ai_to_sub_items,
            open_tdoc_callback=self._open_tdoc_document_directly,
        )
        dlg.thread_changed.connect(self._on_tracked_thread_changed)
        self._tracked_thread_dialogs[tid] = dlg
        dlg.finished.connect(lambda _, k=tid: self._tracked_thread_dialogs.pop(k, None))
        dlg.show()

    def _refresh_inbox_dialog_if_open(self):
        inbox = getattr(self, "_inbox_dialog", None)
        if inbox is None:
            return
        try:
            if inbox.isVisible():
                inbox.reload_from_dbs()
        except RuntimeError:
            self._inbox_dialog = None

    def _on_tracked_thread_changed(self):
        self._sync_tracked_threads_to_table()
        self._refresh_comboboxes()
        self._schedule_row_height_adjustment()
        self._refresh_inbox_dialog_if_open()

    def _build_tdoc_meta_map(self) -> dict:
        """Maps TDoc IDs to their Agenda Item, Sub-Item, and Related WIs for email link inheritance."""
        meta_map = {}
        if not hasattr(self, "model") or not hasattr(self.model, "_data"):
            return meta_map
        for r in self.model._data:
            if r.get("_is_tracked_thread"):
                continue
            td = str(r.get("TDoc") or "").strip().upper()
            if not td:
                continue
            ai = str(r.get("Agenda Item") or "").strip()
            if ai in ("N/A", "-", "None"):
                ai = ""
            sub = str(r.get("Sub-Item") or "").strip()
            if sub in ("N/A", "-", "None"):
                sub = ""
            wis = [
                w for w in self.model._split_related_wis(r.get("Related WIs", ""))
                if w and w.lower() not in ("(no wi)", "none", "nan", "-")
            ]
            meta_map[td] = {"ai": ai, "sub_item": sub, "wis": wis}
        return meta_map

    def _get_pin_meeting_context(self):
        """Returns (agenda_map, valid_wis) for resolving KI/AI/WI tokens against the active meeting."""
        _, wi_sugg, _, _, _ = self._build_thread_autocomplete_suggestions()
        return getattr(self, "agenda_map", None), wi_sugg

    def _process_dropped_outlook_items(self, mime_data):
        """Handles emails dragged directly from Microsoft Outlook UI (FileGroupDescriptor / RenPrivateMessages)."""
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        meeting_id = self._get_meeting_id_label()
        agenda_map, valid_wis = self._get_pin_meeting_context()
        tdoc_meta_map = self._build_tdoc_meta_map()
        try:
            # First try reading Outlook ActiveExplorer Selection (most reliable for .msg drags in Windows)
            pinned_ids = pin_thread_from_outlook_selection(
                self.tracked_threads_db,
                meeting_id=meeting_id,
                wg=wg,
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=tdoc_meta_map,
            )
            if not pinned_ids and mime_data:
                # Fallback: try extracting FileContents if non-empty
                import struct
                formats = mime_data.formats()
                fgd_fmt = next((f for f in formats if "FileGroupDescriptor" in f), None)
                fc_fmt = next((f for f in formats if "FileContents" in f), None)
                if fgd_fmt and fc_fmt:
                    contents = bytes(mime_data.data(fc_fmt))
                    desc = bytes(mime_data.data(fgd_fmt))
                    if contents and len(desc) >= 4:
                        is_uni = "FileGroupDescriptorW" in fgd_fmt
                        rec_size = 592 if is_uni else 332
                        if len(desc) >= 4 + rec_size:
                            raw_name = desc[4 + 72 : 4 + rec_size]
                            fname = (raw_name.decode("utf-16le", errors="ignore") if is_uni else raw_name.decode("mbcs", errors="ignore")).split("\0")[0].strip()
                            if fname.lower().endswith(".msg"):
                                drop_dir = self.meeting_dir / "Agenda" / "DroppedEmails"
                                drop_dir.mkdir(parents=True, exist_ok=True)
                                tmp_msg = drop_dir / fname
                                tmp_msg.write_bytes(contents)
                                tid = pin_thread_from_msg(
                                    self.tracked_threads_db,
                                    tmp_msg,
                                    meeting_id=meeting_id,
                                    wg=wg,
                                    agenda_map=agenda_map,
                                    valid_wis=valid_wis,
                                    tdoc_meta_map=tdoc_meta_map,
                                )
                                if tid:
                                    pinned_ids.append(tid)
            if pinned_ids:
                self._on_tracked_thread_changed()
                self._open_tracked_thread_dialog(pinned_ids[-1])
            else:
                QMessageBox.warning(
                    self,
                    "Outlook Drag & Drop",
                    "Could not read the dragged email from Microsoft Outlook.\n"
                    "Ensure the email is selected in Classic Outlook or save it as a .msg file.",
                )
        except Exception as e:
            logging.error(f"Error handling Outlook drag & drop: {e}", exc_info=True)
            QMessageBox.warning(self, "Outlook Drag & Drop Failed", f"Failed to process dragged Outlook email:\n{e}")

    def _process_dropped_msg_files(self, msg_files: list):
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        meeting_id = self._get_meeting_id_label()
        agenda_map, valid_wis = self._get_pin_meeting_context()
        tdoc_meta_map = self._build_tdoc_meta_map()
        pinned_ids = []
        try:
            for mp in msg_files:
                tid = pin_thread_from_msg(
                    self.tracked_threads_db,
                    Path(mp),
                    meeting_id=meeting_id,
                    wg=wg,
                    agenda_map=agenda_map,
                    valid_wis=valid_wis,
                    tdoc_meta_map=tdoc_meta_map,
                )
                if tid:
                    pinned_ids.append(tid)
            if pinned_ids:
                self._on_tracked_thread_changed()
                self._open_tracked_thread_dialog(pinned_ids[-1])
            else:
                QMessageBox.warning(
                    self,
                    "Could Not Parse .msg",
                    "Failed to parse the dropped .msg file via Outlook COM. Ensure Microsoft Outlook is available.",
                )
        except Exception as e:
            logging.error(f"Error processing dropped .msg file(s): {e}", exc_info=True)
            QMessageBox.warning(self, "Error Processing .msg", f"Failed to process dropped .msg file:\n{e}")

    def _on_pin_thread_from_msg_dialog(self):
        from PyQt5.QtWidgets import QFileDialog
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Outlook .msg File(s) to Pin Thread", "", "Outlook Message (*.msg)")
        if paths:
            self._process_dropped_msg_files([Path(p) for p in paths])

    def _on_pin_thread_from_outlook(self):
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        folders = load_wg_email_config(wg)
        if not folders:
            if QMessageBox.question(
                self,
                "No Folders Configured",
                f"No Outlook folders configured for {wg}.\nConfigure folders now?",
                QMessageBox.Yes | QMessageBox.No,
            ) == QMessageBox.Yes:
                self._on_configure_email_folders()
            return
        agenda_map, valid_wis = self._get_pin_meeting_context()
        dlg = PinFromOutlookDialog(
            self.tracked_threads_db,
            self._get_meeting_id_label(),
            wg,
            self,
            agenda_map=agenda_map,
            valid_wis=valid_wis,
            meeting_dir=getattr(self, "meeting_dir", None),
            general_email_db=getattr(self, "general_email_db", None),
            valid_tdocs=getattr(self.model, "valid_tdocs", set()),
            mtg_start_date=str(self.mtg_info.get("start_date") or ""),
            mtg_end_date=str(self.mtg_info.get("end_date") or ""),
            tdoc_meta_map=self._build_tdoc_meta_map(),
        )
        if dlg.exec_() == QDialog.Accepted and dlg.pinned_thread_ids:
            self._on_tracked_thread_changed()
            self._open_tracked_thread_dialog(dlg.pinned_thread_ids[-1])

    def _open_meeting_email_inbox(self):
        """Opens the unified Unread & Flagged Email Inbox cockpit across Tracked Threads and TDoc Emails."""
        existing = getattr(self, "_inbox_dialog", None)
        if existing is not None:
            try:
                existing.reload_from_dbs()
                existing.setWindowState(existing.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
                existing.raise_()
                existing.activateWindow()
                return
            except RuntimeError:
                self._inbox_dialog = None

        tdoc_ai_lookup = {}
        if hasattr(self, "model") and hasattr(self.model, "_data"):
            for r in self.model._data:
                if r.get("_is_tracked_thread"):
                    continue
                td = str(r.get("TDoc") or "").strip().upper()
                ai = str(r.get("Agenda Item") or "").strip()
                if td and ai:
                    tdoc_ai_lookup[td] = ai

        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        ai_sugg, wi_sugg, _tdoc_sugg, _sub_sugg, ai_to_sub_items = self._build_thread_autocomplete_suggestions()
        dlg = MeetingEmailInboxDialog(
            tracked_db=getattr(self, "tracked_threads_db", None),
            general_email_db=getattr(self, "general_email_db", None),
            meeting_id=self._get_meeting_id_label(),
            tdoc_ai_lookup=tdoc_ai_lookup,
            open_thread_callback=self._open_tracked_thread_dialog,
            open_tdoc_callback=self._open_tdoc_document_directly,
            parent=None,
            wg=wg,
            ai_suggestions=ai_sugg,
            wi_suggestions=wi_sugg,
            ai_to_sub_items=ai_to_sub_items,
            tdoc_meta_map=self._build_tdoc_meta_map(),
            agenda_map=getattr(self, "agenda_map", None),
            meeting_start=str(self.mtg_info.get("start_date", "") or ""),
            meeting_end=str(self.mtg_info.get("end_date", "") or ""),
        )
        dlg.data_changed.connect(self._refresh_email_counts)
        self._inbox_dialog = dlg
        dlg.finished.connect(lambda _=None: setattr(self, "_inbox_dialog", None))
        dlg.show()

    def _on_refresh_tracked_threads(self, silent: bool = False):
        if getattr(self, "_tracked_threads_sync_worker", None) and self._tracked_threads_sync_worker.isRunning():
            if not silent:
                notify_info(self, "Thread Sync Running", "Tracked email thread sync is already in progress.")
            return
        if not hasattr(self, "tracked_threads_db"):
            return
        meeting_id = self._get_meeting_id_label()
        threads = self.tracked_threads_db.get_threads_for_meeting(meeting_id)
        if not threads:
            if not silent:
                notify_info(self, "Tracked Threads", "No pinned email threads found for this meeting.")
            return

        self._sync_silent_mode = bool(silent)
        titles = [str(t.get("title") or t.get("thread_id") or "") for t in threads]
        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.start_busy("Syncing tracked threads...", total=len(titles), remaining=titles)
        if hasattr(self, "email_btn"):
            if not hasattr(self, "_default_email_btn_tooltip"):
                self._default_email_btn_tooltip = self.email_btn.toolTip()
            self.email_btn.setText(f"⏳ Emails (0/{len(titles)}) ▾")

        worker = TrackedThreadsRefreshWorker(self.tracked_threads_db.db_path, meeting_id, parent=self)
        worker.progress.connect(self._on_tracked_threads_sync_progress)
        worker.thread_updated.connect(lambda _tid, _n: (self._sync_tracked_threads_to_table(), self._refresh_inbox_dialog_if_open()))
        worker.finished_sync.connect(self._on_tracked_threads_sync_finished)
        self._tracked_threads_sync_worker = worker
        worker.start()

    def _on_cancel_tracked_threads_sync(self):
        if getattr(self, "_tracked_threads_sync_worker", None) and self._tracked_threads_sync_worker.isRunning():
            self._tracked_threads_sync_worker.request_cancel()

    def _on_tracked_threads_sync_progress(self, completed: int, total: int, current_title: str, remaining: list):
        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.update_progress(completed, total, current_title, remaining)
        if hasattr(self, "email_btn"):
            self.email_btn.setText(f"⏳ Emails ({completed}/{total}) ▾")
            rem_txt = "\n• ".join(remaining[:8]) if remaining else "(none)"
            self.email_btn.setToolTip(
                f"Syncing thread ({completed}/{total}): {current_title}\n\nRemaining:\n• {rem_txt}"
            )

    def _on_tracked_threads_sync_finished(self, ok: bool, updates: dict, err: str):
        silent = bool(getattr(self, "_sync_silent_mode", False))
        if hasattr(self, "email_btn") and hasattr(self, "_default_email_btn_tooltip"):
            self.email_btn.setToolTip(self._default_email_btn_tooltip)
        self._on_tracked_thread_changed()
        if not ok:
            if hasattr(self, "thread_sync_bar"):
                self.thread_sync_bar.show_error(err or "Tracked thread sync interrupted.")
            notify_toast(
                self,
                "Tracked Thread Sync Warning",
                err or "Tracked thread sync did not complete cleanly.",
                variant="warning",
            )
            return

        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.finish_ok()
        new_total = sum((updates or {}).values())
        if new_total > 0:
            notify_success(
                self,
                "Tracked Threads Refreshed",
                f"Found {new_total} new message(s) across {len(updates)} tracked thread(s).",
            )
            if silent:
                QToolTip.showText(
                    QCursor.pos(),
                    f"📧 {new_total} new reply/replies across {len(updates)} tracked email thread(s)!",
                    self,
                )
        elif not silent:
            notify_info(
                self,
                "Tracked Threads Up-to-Date",
                "No new replies found in configured Outlook folders.",
            )

    def _on_toggle_auto_refresh_threads(self, enabled: bool):
        if enabled:
            self.tracked_threads_timer.start()
        else:
            self.tracked_threads_timer.stop()

    def _refresh_email_counts(self):
        try:
            counts = self.general_email_db.get_email_counts_per_tdoc()
            self.model.set_email_counts(counts)
            self._sync_tracked_threads_to_table()
            if self.sender() is not getattr(self, "_inbox_dialog", None):
                self._refresh_inbox_dialog_if_open()
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

        def _pin_discussion_as_thread(info: dict):
            if not hasattr(self, "tracked_threads_db"):
                return
            agenda_map, valid_wis = self._get_pin_meeting_context()
            subj = str(info.get("subject") or f"Discussion on {base_tdoc}").strip()
            norm = str(info.get("normalized_subject") or subj).strip()
            tid = pin_thread_by_subject_or_cid(
                self.tracked_threads_db,
                title=norm or subj,
                normalized_subject=norm,
                meeting_id=self._get_meeting_id_label(),
                wg=wg,
                linked_tdocs=list(info.get("linked_tdocs") or family),
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=self._build_tdoc_meta_map(),
            )
            self._on_tracked_thread_changed()
            self._open_tracked_thread_dialog(tid)

        dialog = TDocEmailsDialog(
            base_tdoc,
            family,
            self.general_email_db.db_path,
            wg=wg,
            save_callback=_save_notes_proxy,
            parent=None,
            pin_thread_callback=_pin_discussion_as_thread,
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
        from modules.emails.core.tracked_threads_service import mark_thread_read_with_outlook
        synced_count = self.general_email_db.mark_all_read(sync_outlook=True)
        if hasattr(self, "tracked_threads_db"):
            for t in self.tracked_threads_db.get_threads_for_meeting(self._get_meeting_id_label()):
                synced_count += mark_thread_read_with_outlook(self.tracked_threads_db, t["thread_id"], sync_outlook=True)
        self._refresh_email_counts()
        self._refresh_comboboxes()
        QMessageBox.information(
            self,
            "Mark All Read",
            f"All meeting emails and tracked threads have been marked as read (synced {synced_count} Outlook item(s)).",
        )

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

    def _open_help_dialog(self):
        """Opens the comprehensive TDocs Cockpit Quick Guide & Shortcuts dialog."""
        from modules.meetings.ui.tdocs_help_dialog import TDocsHelpDialog
        dlg = TDocsHelpDialog(self, is_ran=getattr(self, "is_ran", False))
        dlg.exec_()

    def closeEvent(self, event):
        if hasattr(self, 'routing_timer'):
            self.routing_timer.stop()
        if hasattr(self, 'tracked_threads_timer'):
            self.tracked_threads_timer.stop()
        if getattr(self, '_tracked_threads_sync_worker', None) and self._tracked_threads_sync_worker.isRunning():
            try:
                self._tracked_threads_sync_worker.request_cancel()
                self._tracked_threads_sync_worker.wait(1500)
            except Exception:
                pass
        if hasattr(self, 'search_timer'):
            self.search_timer.stop()

        if getattr(self, 'email_window', None):
            try:
                self.email_window.close()
            except Exception:
                pass

        if getattr(self, '_session_plan_window', None):
            try:
                self._session_plan_window.close()
            except Exception:
                pass
            self._session_plan_window = None

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

        for dlg in list(getattr(self, '_tracked_thread_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_tracked_thread_dialogs'):
            self._tracked_thread_dialogs.clear()

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
            getattr(self, 'cacher_thread', None),
            getattr(self, 'cli_md_export_thread', None),
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
