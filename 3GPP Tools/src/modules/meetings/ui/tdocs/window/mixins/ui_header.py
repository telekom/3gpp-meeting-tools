# --- File: src/modules/meetings/ui/tdocs/window/mixins/ui_header.py ---
"""
Header/toolbar construction, resources menu, instant fetch/compare, routing indicator, venue clock.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import datetime
import logging
import re
import webbrowser
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QMenu,
    QApplication,
    QShortcut,
    QSizePolicy,
)
from core.network.network_state import NetworkState
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.world_clock_service import WorldClockService
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdocs_components import make_rich_tooltip
from modules.meetings.ui.meetings_menu import add_meeting_portal_actions
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.terminal_opener import open_terminal as _open_terminal
from core.ui.theme import C


class HeaderMixin:
    """Header/toolbar construction, resources menu, instant fetch/compare, routing indicator, venue clock."""

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

        self.email_btn.setMenu(self._build_email_menu())

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
