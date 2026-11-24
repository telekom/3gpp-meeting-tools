# --- File: src/modules/meetings/ui/tdocs/window/mixins/header_widgets.py ---
"""Toolbar widget, menu, shortcut, and FlowLayout assembly for `TDocsWindow`."""

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QShortcut,
    QSizePolicy,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.ui.world_clock_service import WorldClockService
from modules.meetings.ui.tdocs_components import make_rich_tooltip


class HeaderWidgetsMixin:
    """Toolbar construction, menus, shortcuts, and FlowLayout assembly."""

    def _build_refresh_menu(self) -> QMenu:
        refresh_menu = QMenu(self)
        refresh_menu.setToolTipsVisible(True)
        act = refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        act.setToolTip("Download and reload the latest official TDoc Excel list for this meeting.")
        act = refresh_menu.addAction("📋 Download Agenda CSV", self._download_agenda_csv)
        act.setToolTip("Download agenda.csv and refresh Agenda Item names/acronyms used by the TDoc view.")
        act = refresh_menu.addAction("📥 Cache All TDocs (Docs/)", self._cache_all_tdocs)
        act.setToolTip(
            "Download and cache all meeting TDocs (.zip) from the 3GPP Docs directory locally using parallel downloads."
        )
        act = refresh_menu.addAction("📥 Download Chairman's Notes", self._download_chair_notes)
        act.setToolTip(
            "Browse and download Chairman's Notes from /Inbox/Chair_Notes with pre-selected latest files per chair/track."
        )
        if self.is_ran:
            act = refresh_menu.addAction(
                "🔍 Scan Existing Chairman's Notes for Agreements",
                self._scan_existing_chair_notes,
            )
            act.setToolTip(
                "Scan local Chairman's Notes files (.docx) in Agenda/ChairNotes, "
                "extract current agreements, and merge into the meeting table."
            )
        act = refresh_menu.addAction(
            "📝 Import Chairman's Notes / Agenda (.docx / .doc / .zip)...",
            self._import_word_agenda_dialog,
        )
        act.setToolTip(
            "Import one or more local Word, ZIP, or HTML agenda/notes files and merge their TDoc remarks into the table."
        )
        act = refresh_menu.addAction("⚖️ Track Study Item && Meeting Agreements...", self._open_agreement_tracker)
        act.setToolTip(
            "Open the Cross-Meeting Study Item & Topic Agreement Tracker pre-filtered to this meeting's Working Group."
        )
        if self.is_sa2:
            act = refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)
            act.setToolTip("Fetch TdocsByAgenda.htm and merge chairman/secretary agenda metadata into the TDoc list.")
        if self.is_sa2_electronic:
            act = refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            act.setToolTip("Refresh the revision list for this electronic meeting without reloading the Excel list.")
            act = refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
            act.setToolTip("Refresh both the official TDoc Excel list and the available revision metadata.")
        return refresh_menu

    def _build_export_menu(self) -> QMenu:
        export_menu = QMenu(self)
        export_menu.setToolTipsVisible(True)

        export_specs = (
            (
                "📊 Export to Excel (.xlsx)...",
                self._open_excel_export_dialog,
                "Export TDoc metadata to a formatted Excel workbook. "
                "You can choose visible/all rows, columns, their order, and the destination file.",
            ),
            (
                "📦 Export TDocs + Excel...",
                self._export_tdocs_package,
                "Create a self-contained package under this meeting's Export folder. "
                "Exports selected TDoc rows, or all currently visible rows when nothing is selected, "
                "using cached documents where available and including an Excel manifest.",
            ),
            (
                "📄 Export TDocs to Markdown (.md)",
                self._export_markdown_visible,
                "Convert and export all currently visible/filtered TDocs to individual Markdown (.md) "
                "documents with an index manifest.",
            ),
            (
                "📝 Export Markdown Summary Reports",
                self._export_reports,
                "Generate Markdown summary reports for the meeting from the currently loaded TDoc data.",
            ),
            (
                "🧭 Export Pre-Meeting Report (.md)...",
                self._open_pre_meeting_export_dialog,
                "Generate a structured, per-release and per-WI pre-meeting Markdown briefing covering incoming LS, "
                "competing LS reply proposals, proposed CRs/pCRs, and company discussion papers.",
            ),
            (
                "📑 Export Post-Meeting Report (.md)...",
                self._open_post_meeting_export_dialog,
                "Generate a structured, per-release and per-WI post-meeting Markdown dossier covering "
                "incoming LS and agreed pCR/CR/LS.",
            ),
            (
                "📑 Export Navigable Dossier (HTML && PDF)...",
                self._export_dossier,
                "Compile selected or visible TDocs into an interactive HTML dossier and a bookmarked vector PDF "
                "with companion AI citation prompt.",
            ),
        )
        for label, slot, tip in export_specs:
            act = export_menu.addAction(label, slot)
            act.setToolTip(tip)
            act.setStatusTip(tip)

        self.cli_export_submenu = export_menu.addMenu("💬 Export && Open in CLI Agent")
        self.cli_export_submenu.setToolTip(
            "Pre-convert selected (or visible) TDocs to Markdown and launch Codex, Claude, or Gemini CLI."
        )
        self.cli_export_submenu.aboutToShow.connect(self._populate_export_cli_submenu)

        self.export_menu = export_menu
        self.export_menu_sep = export_menu.addSeparator()
        self.act_open_export_dir = export_menu.addAction("📂 Open Export Folder", self._open_export_folder)
        self.act_open_export_dir.setToolTip(
            "Open this meeting's local Export directory in Windows Explorer (Ctrl+Shift+E)."
        )
        self.act_open_export_dir.setStatusTip(self.act_open_export_dir.toolTip())
        self.act_open_export_dir.setShortcut(QKeySequence("Ctrl+Shift+E"))
        self.addAction(self.act_open_export_dir)
        self.export_menu.aboutToShow.connect(self._on_export_menu_about_to_show)
        return export_menu

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
        title_lbl.setToolTip(
            "Electronic Meeting (eMeeting)" if is_electronic else "In-Person Meeting (Face-to-Face)"
        )
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
        self.last_mod_lbl.setStyleSheet(
            f"font-size: 11px; color: {C.TEXT_MUTED}; margin-right: 8px; font-style: italic;"
        )

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.setToolTip(make_rich_tooltip(
            title="🔄 Refresh Meeting Data",
            description="Reload official TDoc list, download Agenda CSV, cache documents offline, or scan notes.",
            tip="Click arrow ▾ to view options including offline caching and notes scanning.",
        ))
        self.refresh_btn.setMenu(self._build_refresh_menu())

        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.folder_btn.setToolTip(make_rich_tooltip(
            title="🗂️ Resources & Directories",
            description=(
                "Quick access to local cache directories (Docs, Agenda, ChairNotes, SessionPlans), "
                "remote 3GPP FTP folders, and official 3GU Meeting Portal links."
            ),
            tip="You can also open a Windows Terminal directly inside the local meeting cache folder.",
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
            tip="Use this to view unedited official committee columns or run external Excel formulas.",
        ))
        self.excel_btn.clicked.connect(self._open_excel)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.export_btn.setToolTip(make_rich_tooltip(
            title="📤 Export Reports & Dossiers",
            shortcut="Ctrl+Shift+E",
            description="Export meeting TDoc metadata, document packages, LLM-ready content, or Markdown reports.",
            tip="Press Ctrl+Shift+E anytime to reveal the export folder in Windows File Explorer.",
        ))
        self.export_btn.setMenu(self._build_export_menu())

        self.session_plan_btn = QPushButton("🗓️ Session Plan")
        self.session_plan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.session_plan_btn.setToolTip(make_rich_tooltip(
            title="🗓️ Session Plan & Timetable",
            description=(
                "Interactive visual schedule viewer parsing session plan documents. "
                "Displays parallel tracks, room locations, and timings."
            ),
            tip="Automatically refreshes when new session plan documents are uploaded during the meeting.",
        ))
        self.session_plan_btn.clicked.connect(self._open_session_plan_window)

        self.btn_meeting_agreements = QPushButton("⚖️ Agreements")
        self.btn_meeting_agreements.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_meeting_agreements.setToolTip(make_rich_tooltip(
            title="⚖️ Chairman's Notes Agreements (RAN)",
            description="View, search, and export all Chairman's Notes agreements and working assumptions for this meeting.",
            tip="Double-clicking an agreement cell in the table will open this viewer focused on that agreement.",
        ))
        self.btn_meeting_agreements.clicked.connect(lambda: self._open_meeting_agreements_dialog())

        self.draft_cr_btn = QPushButton("📝 Draft CR / pCR")
        self.draft_cr_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.draft_cr_btn.setToolTip(make_rich_tooltip(
            title="📝 Draft CR / pCR Contribution",
            description=(
                "Compose an official Change Request (CR) or Pseudo-CR (pCR) contribution "
                "adhering strictly to MCC formatting rules."
            ),
            tip="Select a baseline spec or TDoc in the table to pre-populate revision headers and clauses.",
        ))
        self.draft_cr_btn.clicked.connect(lambda: self._open_draft_cr_dialog())

        self.stats_btn = QPushButton("📊 Statistics")
        self.stats_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.stats_btn.setToolTip(make_rich_tooltip(
            title="📊 Meeting Statistics & Analytics",
            description=(
                "Generate an interactive HTML analytics dashboard showing company contribution volumes, "
                "document types, and agenda distributions."
            ),
            tip="Click ⚙️ to configure custom company groupings or export parameters.",
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
            description="Customize company alias unification, contribution categorization, and output charts.",
        ))
        self.stats_cfg_btn.clicked.connect(self._open_stats_config)

        self.email_btn = QPushButton("📧 Emails ▾")
        self.email_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.email_btn.setToolTip(make_rich_tooltip(
            title="📧 Email Discussions & Reflector Sync",
            description="Sync and inspect Outlook emails linked to TDocs and revisions, strip quotes, and categorize stances.",
            tip="Discussion counts are displayed in the 'Emails' table column with stance breakdown badges.",
        ))
        self.email_btn.setMenu(self._build_email_menu())

        self.btn_help = QPushButton("❓ Help")
        self.btn_help.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_help.setToolTip(make_rich_tooltip(
            title="❓ Help & Shortcuts Guide",
            shortcut="F1",
            description=(
                "Open the comprehensive interactive user guide, table double-click cheatsheet, "
                "search syntax, and keyboard shortcuts."
            ),
            tip="Press F1 anywhere in the TDocs window to open instantly.",
        ))
        self.btn_help.clicked.connect(self._open_help_dialog)

        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self._open_help_dialog)

        self.count_lbl = QLabel(f"Showing {count} of {count} TDocs")
        self.count_lbl.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {C.TEXT_SECONDARY};")

        self.instant_fetch_input = QLineEdit()
        self.instant_fetch_input.setPlaceholderText("Instant Fetch...")
        self.instant_fetch_input.setToolTip(make_rich_tooltip(
            title="⚡ Instant TDoc Fetch",
            shortcut="Enter",
            description=(
                "Type any TDoc number (e.g. R2-2601234 or S2-2605678) and press Enter to "
                "download, unpack, and launch it immediately."
            ),
            tip="Bypasses table scrolling completely when documents are called up live in the meeting room.",
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
            description=(
                "Download the TDoc entered in Instant Fetch and immediately run a Track Changes "
                "Word comparison against your active Word draft."
            ),
            tip="Perfect for reviewing a newly posted revision against your local edits in seconds.",
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
