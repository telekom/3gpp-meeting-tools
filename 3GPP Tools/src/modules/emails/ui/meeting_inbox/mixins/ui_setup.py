"""UI construction mixin for MeetingEmailInboxDialog."""

from __future__ import annotations

from PyQt5.QtCore import Qt, QTimer, QUrl
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill, scale_px
from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS
from modules.meetings.ui.topic_pills import TagCloudWidget
from ..widgets import (
    _InboxStatusStripDelegate,
    _InboxTableWidget,
    _NotePlainTextEdit,
    get_compact_action_bar_qss,
)


class InboxUiSetupMixin:
    """Builds the toolbar, filter bar, inbox table, detail header, and reader panes."""

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Row 1: Responsive FlowLayout for State Pills + View Toggle + Search + Dropdown Filters + Reload
        top_widget = QWidget(self)
        top_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        top_sp.setHeightForWidth(True)
        top_widget.setSizePolicy(top_sp)
        top_row = FlowLayout(top_widget, margin=0, h_spacing=6, v_spacing=4)

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        pill_style = (
            f"QPushButton {{ padding: 4px 10px; border-radius: 13px; border: 1px solid {C.BORDER_STRONG}; "
            f"background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; font-size: 12px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_CANVAS}; }} "
            f"QPushButton:checked {{ background-color: {C.INFO_SOFT_STRONG}; border-color: {C.ACCENT_ALT}; "
            f"color: {C.INFO}; font-weight: bold; }}"
        )

        self.pill_unread = QPushButton("🔵 Unread Only (0)")
        self.pill_unread.setCheckable(True)
        self.pill_unread.setChecked(self._active_mode == "UNREAD")
        self.pill_unread.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_unread)
        self.pill_unread.clicked.connect(lambda: self._set_mode("UNREAD"))
        top_row.addWidget(self.pill_unread)

        self.pill_flagged = QPushButton("🚩 Flagged (0)")
        self.pill_flagged.setCheckable(True)
        self.pill_flagged.setChecked(self._active_mode == "FLAGGED")
        self.pill_flagged.setStyleSheet(
            f"QPushButton {{ padding: 4px 10px; border-radius: 13px; border: 1px solid {C.BORDER_STRONG}; "
            f"background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; font-size: 12px; }} "
            f"QPushButton:hover {{ background-color: {C.DANGER_SOFT}; }} "
            f"QPushButton:checked {{ background-color: {C.DANGER_SOFT_HOVER}; border-color: {C.DANGER_ALT}; "
            f"color: {C.DANGER}; font-weight: bold; }}"
        )
        self.mode_group.addButton(self.pill_flagged)
        self.pill_flagged.clicked.connect(lambda: self._set_mode("FLAGGED"))
        top_row.addWidget(self.pill_flagged)

        self.pill_recent = QPushButton("🕒 Last 24h (0)")
        self.pill_recent.setCheckable(True)
        self.pill_recent.setChecked(self._active_mode == "RECENT")
        self.pill_recent.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_recent)
        self.pill_recent.clicked.connect(lambda: self._set_mode("RECENT"))
        top_row.addWidget(self.pill_recent)

        self.pill_all = QPushButton("📋 All (0)")
        self.pill_all.setCheckable(True)
        self.pill_all.setChecked(self._active_mode == "ALL")
        self.pill_all.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_all)
        self.pill_all.clicked.connect(lambda: self._set_mode("ALL"))
        top_row.addWidget(self.pill_all)

        # View Mode Toggle: By Message vs. By Thread
        view_style = (
            f"QPushButton {{ padding: 4px 9px; border-radius: 5px; border: 1px solid {C.BORDER_STRONG}; "
            f"background-color: {C.BG_SURFACE}; color: {C.TEXT_BODY}; font-size: 12px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_CANVAS}; }} "
            f"QPushButton:checked {{ background-color: {C.WG_SA_BG}; border-color: {C.INDIGO_DARK}; "
            f"color: {C.INDIGO_DARK}; font-weight: bold; }}"
        )
        self.view_mode_group = QButtonGroup(self)
        self.view_mode_group.setExclusive(True)
        self.btn_view_messages = QPushButton("📨 By Message")
        self.btn_view_messages.setCheckable(True)
        self.btn_view_messages.setChecked(self._view_mode == "MESSAGES")
        self.btn_view_messages.setStyleSheet(view_style)
        self.btn_view_messages.setToolTip("Show individual email messages sorted newest first.")
        self.btn_view_messages.clicked.connect(lambda: self._set_view_mode("MESSAGES"))
        self.view_mode_group.addButton(self.btn_view_messages)
        top_row.addWidget(self.btn_view_messages)

        self.btn_view_threads = QPushButton("🧵 By Thread")
        self.btn_view_threads.setCheckable(True)
        self.btn_view_threads.setChecked(self._view_mode == "THREADS")
        self.btn_view_threads.setStyleSheet(view_style)
        self.btn_view_threads.setToolTip("Group emails by discussion thread for fast AI/WI assignment and thread triage.")
        self.btn_view_threads.clicked.connect(lambda: self._set_view_mode("THREADS"))
        self.view_mode_group.addButton(self.btn_view_threads)
        top_row.addWidget(self.btn_view_threads)

        self.search_edit = QLineEdit()
        self.search_edit.setMinimumWidth(190)
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setPlaceholderText("🔍 Search subject, sender, company, TDoc, or body...")
        if self._saved_search_text:
            self.search_edit.setText(self._saved_search_text)
        self.search_edit.textChanged.connect(self._apply_filters)
        top_row.addWidget(self.search_edit)

        self.company_combo = CheckableComboBox("Company")
        self.company_combo.setMinimumWidth(scale_px(135))
        self.company_combo.setMaximumWidth(scale_px(185))
        self.company_combo.setToolTip("Filter by one or more participating/sender companies.")
        self.company_combo.selectionChanged.connect(self._apply_filters)
        top_row.addWidget(self.company_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMinimumWidth(scale_px(130))
        self.ai_combo.setMaximumWidth(scale_px(210))
        self.ai_combo.setToolTip("Filter by one or more Agenda Items (or Unassigned).")
        self.ai_combo.selectionChanged.connect(self._apply_filters)
        top_row.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("Status")
        self.status_combo.setMinimumWidth(scale_px(120))
        self.status_combo.setMaximumWidth(scale_px(175))
        self.status_combo.setToolTip("Filter by thread/delegate status (My Status).")
        self.status_combo.selectionChanged.connect(self._apply_filters)
        top_row.addWidget(self.status_combo)

        self.source_combo = QComboBox()
        self.source_combo.addItem("All Sources", "")
        self.source_combo.addItem("⚠️ Needs Assignment (No AI / Unpinned)", "NEEDS_TRIAGE")
        self.source_combo.addItem("🧵 Tracked Threads Only", "THREAD")
        self.source_combo.addItem("📄 Unpinned Only", "UNPINNED")
        if self._saved_source_filter:
            src_idx = self.source_combo.findData(self._saved_source_filter)
            if src_idx >= 0:
                self.source_combo.setCurrentIndex(src_idx)
        self.source_combo.setToolTip("Filter by Tracked Thread vs. Unpinned TDoc Email, or isolate items needing assignment.")
        self.source_combo.currentIndexChanged.connect(self._apply_filters)
        top_row.addWidget(self.source_combo)

        self.refresh_btn = QPushButton("🔄 Reload")
        self.refresh_btn.setToolTip("Reload latest messages already saved in local meeting databases.")
        self.refresh_btn.clicked.connect(self.reload_from_dbs)
        top_row.addWidget(self.refresh_btn)

        self.sync_outlook_btn = QToolButton()
        self.sync_outlook_btn.setText("📡 Sync Outlook")
        self.sync_outlook_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.sync_outlook_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.sync_outlook_btn.setToolTip(
            "Pull new Tracked Thread replies, TDoc emails, and Outlook read/flag updates asynchronously."
        )
        self._sync_menu = QMenu(self.sync_outlook_btn)
        self._act_sync_selected = self._sync_menu.addAction("🔄 Sync Selected Thread")
        self._act_sync_selected.setEnabled(False)
        self._act_sync_selected.triggered.connect(self._on_sync_selected_clicked)
        self._act_sync_filtered = self._sync_menu.addAction("⚡ Sync Filtered Only")
        self._act_sync_filtered.triggered.connect(lambda: self._start_outlook_sync(scoped=True))
        self._act_sync_full = self._sync_menu.addAction("📡 Full Sync (All Threads && Folders)")
        self._act_sync_full.triggered.connect(lambda: self._start_outlook_sync(scoped=False))
        self.sync_outlook_btn.setMenu(self._sync_menu)
        self.sync_outlook_btn.clicked.connect(self._on_sync_outlook_clicked)
        self._default_sync_scoped = False
        top_row.addWidget(self.sync_outlook_btn)

        self.btn_layout_toggle = QPushButton("⬒ List on Top")
        self.btn_layout_toggle.setToolTip("Switch between Side-by-Side (desktop) and Top/Bottom (laptop) layout.")
        self.btn_layout_toggle.clicked.connect(self._toggle_layout_orientation)
        top_row.addWidget(self.btn_layout_toggle)

        self.help_pill_btn = create_help_tips_pill(
            [
                "Ctrl + Mouse Wheel in reading pane: zoom text larger/smaller (Ctrl+0 to reset; saved automatically)",
                "Ctrl + Click / Shift + Click rows: select multiple emails or threads to batch-assign AI, Sub-Topics, WI, Pin, Read, or Flag",
                "Click blue TDoc links (e.g. S2-xxxxxx, R3-xxxxxx) to open documents directly",
                "Side-by-Side / List on Top: switch desktop vs. laptop split layout",
            ],
            parent=top_widget,
            title="Email Inbox Tips",
        )
        top_row.addWidget(self.help_pill_btn)

        layout.addWidget(top_widget)

        # Row 2: Master-Detail Splitter
        self.splitter = QSplitter(Qt.Horizontal, self)
        self.splitter.setChildrenCollapsible(False)

        left_widget = QWidget()
        self._left_widget = left_widget
        left_widget.setMinimumWidth(240)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self.table = _InboxTableWidget(0, 6, left_widget)
        self.table.setHorizontalHeaderLabels([
            "State",
            "Source / Context",
            "AI",
            "Subject",
            "Sender",
            "Date",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setItemDelegateForColumn(0, _InboxStatusStripDelegate(self.table))
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setStretchLastSection(False)
        default_widths = [
            scale_px(76),
            scale_px(155),
            scale_px(80),
            scale_px(360),
            scale_px(150),
            scale_px(120),
        ]
        for col_i, col_w in enumerate(default_widths):
            hdr.resizeSection(col_i, col_w)
        self._col_base_widths = list(default_widths)
        prefix = self._meeting_settings_prefix()
        saved_hdr = self._settings.value(f"{prefix}/table_header_state")
        if saved_hdr:
            self._restoring_header = True
            try:
                if hdr.restoreState(saved_hdr):
                    self._col_base_widths = [max(40, hdr.sectionSize(c)) for c in range(6)]
                hdr.setSectionResizeMode(QHeaderView.Interactive)
            except Exception:
                pass
            finally:
                self._restoring_header = False
        hdr.sectionResized.connect(self._on_header_section_resized)
        self.table.viewportResized.connect(lambda: QTimer.singleShot(0, self._fit_columns_to_viewport))
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        left_layout.addWidget(self.table, 1)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        left_layout.addWidget(self.status_lbl)

        # Right Detail Pane
        right_widget = QWidget()
        self._right_widget = right_widget
        right_widget.setMinimumWidth(280)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 0, 0, 0)
        right_layout.setSpacing(4)

        meta_frame = QFrame(right_widget)
        meta_frame.setObjectName("InboxMetaFrame")
        meta_frame.setStyleSheet(
            f"#InboxMetaFrame {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        meta_layout = QVBoxLayout(meta_frame)
        meta_layout.setContentsMargins(8, 4, 8, 4)
        meta_layout.setSpacing(2)

        self.detail_subject_lbl = QLabel("Select an email to read")
        self.detail_subject_lbl.setWordWrap(True)
        self.detail_subject_lbl.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        meta_layout.addWidget(self.detail_subject_lbl)

        meta_info_row = QHBoxLayout()
        meta_info_row.setContentsMargins(0, 0, 0, 0)
        meta_info_row.setSpacing(8)
        self.detail_meta_lbl = QLabel("")
        self.detail_meta_lbl.setWordWrap(True)
        self.detail_meta_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        meta_info_row.addWidget(self.detail_meta_lbl)

        self.detail_links_lbl = QLabel("")
        self.detail_links_lbl.setWordWrap(True)
        self.detail_links_lbl.setTextFormat(Qt.RichText)
        self.detail_links_lbl.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.detail_links_lbl.setOpenExternalLinks(False)
        self.detail_links_lbl.linkActivated.connect(lambda href: self._on_anchor_clicked(QUrl(href)))
        self.detail_links_lbl.setStyleSheet(f"color: {C.INFO}; font-size: 11px;")
        meta_info_row.addWidget(self.detail_links_lbl, 1)
        meta_layout.addLayout(meta_info_row)

        # Inline Assignment Bar (Single compact row: AIs, Sub-Topics, WIs + Auto/Pin)
        self._show_assign_bar = bool(self._settings.value("show_assign_bar", True, type=bool))
        self.assign_frame = QFrame(meta_frame)
        self.assign_frame.setObjectName("InboxAssignBar")
        self.assign_frame.setStyleSheet(
            f"#InboxAssignBar {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 5px; }}"
        )
        assign_row = QHBoxLayout(self.assign_frame)
        assign_row.setContentsMargins(6, 2, 6, 2)
        assign_row.setSpacing(6)

        strict_ai = bool(self._ai_suggestions)
        strict_wi = bool(self._wi_suggestions)

        assign_row.addWidget(QLabel("<b>AI:</b>"))
        self.assign_ais_pills = TagCloudWidget(
            placeholder="AI #...",
            pill_colors={"bg": C.INFO_SOFT, "fg": C.INFO, "border": C.INFO_BORDER},
            strict_dropdown=strict_ai,
            dropdown_label="+ AI ▾",
            empty_dropdown_label="No AIs",
        )
        if self._ai_suggestions:
            self.assign_ais_pills.set_suggestions(
                self._ai_suggestions,
                value_extractor=lambda s: s.split(" — ", 1)[0].strip(),
            )
        self.assign_ais_pills.tagsChanged.connect(self._on_inline_links_changed)
        assign_row.addWidget(self.assign_ais_pills, 2)

        assign_row.addWidget(QLabel("<b>Sub:</b>"))
        self.assign_subs_pills = TagCloudWidget(
            placeholder="Sub-Topic...",
            pill_colors={"bg": C.TEAL_SURFACE, "fg": C.TEAL_DARK, "border": C.TEAL_BORDER},
            strict_dropdown=True,
            dropdown_label="+ Sub ▾",
            empty_dropdown_label="Assign AI",
        )
        self.assign_subs_pills.tagsChanged.connect(self._on_inline_links_changed)
        assign_row.addWidget(self.assign_subs_pills, 3)

        assign_row.addWidget(QLabel("<b>WI:</b>"))
        self.assign_wis_pills = TagCloudWidget(
            placeholder="WI...",
            pill_colors={"bg": C.WG_SA_BG, "fg": C.INDIGO_DARK, "border": C.WG_SA_BORDER},
            strict_dropdown=strict_wi,
            dropdown_label="+ WI ▾",
            empty_dropdown_label="No WIs",
        )
        if self._wi_suggestions:
            self.assign_wis_pills.set_suggestions(
                self._wi_suggestions,
                value_extractor=lambda s: s.split(" — ", 1)[0].strip(),
            )
        self.assign_wis_pills.tagsChanged.connect(self._on_inline_links_changed)
        assign_row.addWidget(self.assign_wis_pills, 2)

        self.btn_auto_assign = QPushButton("🪄 Auto")
        self.btn_auto_assign.setToolTip("Infer and apply AIs, Sub-Topics, and WIs from subject and linked TDocs.")
        self.btn_auto_assign.clicked.connect(self._on_auto_assign_clicked)
        assign_row.addWidget(self.btn_auto_assign)

        self.btn_pin_inline = QPushButton("📌 Pin")
        self.btn_pin_inline.setToolTip("Pin this email conversation as a Tracked Thread in-place.")
        self.btn_pin_inline.clicked.connect(self._on_pin_inline_clicked)
        assign_row.addWidget(self.btn_pin_inline)

        self.assign_frame.setVisible(self._show_assign_bar)
        meta_layout.addWidget(self.assign_frame)

        # Inline Always-Editable Comment Strip (auto-visible when a comment exists, toggled via 📝 Note)
        self.note_frame = QFrame(meta_frame)
        self.note_frame.setObjectName("InboxNoteBar")
        self.note_frame.setStyleSheet(
            f"#InboxNoteBar {{ background-color: {C.WARNING_SOFT}; border: 1px solid {C.WARNING_BORDER}; border-radius: 5px; }}"
        )
        note_row = QHBoxLayout(self.note_frame)
        note_row.setContentsMargins(6, 3, 6, 3)
        note_row.setSpacing(6)
        note_lbl = QLabel("<b>📝 Note:</b>")
        note_lbl.setStyleSheet(f"color: {C.WARNING};")
        note_row.addWidget(note_lbl)

        self.note_edit = _NotePlainTextEdit(self.note_frame)
        self.note_edit.setPlaceholderText("Add thread comment / note (auto-saved on blur or Ctrl+Enter)...")
        self.note_edit.setMaximumHeight(scale_px(46))
        self.note_edit.setMinimumHeight(scale_px(28))
        self.note_edit.setStyleSheet(
            f"QPlainTextEdit {{ background-color: {C.BG_CARD}; color: {C.TEXT_PRIMARY}; "
            f"border: 1px solid {C.WARNING_BORDER}; border-radius: 4px; padding: 2px 6px; }}"
        )
        self.note_edit.textChanged.connect(self._on_note_text_changed)
        self.note_edit.focusLost.connect(self._flush_pending_note_save)
        self.note_edit.saveRequested.connect(self._flush_pending_note_save)
        note_row.addWidget(self.note_edit, 1)

        self.note_saved_lbl = QLabel("")
        self.note_saved_lbl.setStyleSheet(f"color: {C.WARNING}; font-size: 10px;")
        note_row.addWidget(self.note_saved_lbl)

        self.note_frame.setVisible(False)
        meta_layout.addWidget(self.note_frame)
        right_layout.addWidget(meta_frame)

        # Compact Triage Action Bar (FlowLayout wraps cleanly when pane narrows)
        act_widget = QWidget(right_widget)
        act_widget.setObjectName("InboxActionBar")
        act_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        act_sp.setHeightForWidth(True)
        act_widget.setSizePolicy(act_sp)
        act_widget.setStyleSheet(get_compact_action_bar_qss("InboxActionBar"))
        act_row = FlowLayout(act_widget, margin=0, h_spacing=4, v_spacing=3)

        self.btn_mark_next = QPushButton("✔️ Read && Next")
        self.btn_mark_next.setToolTip("Mark the current message/thread Read and jump to the next unread item.")
        self.btn_mark_next.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_ALT}; color: {C.TEXT_ON_ACCENT}; font-weight: bold; "
            f"padding: 3px 10px; border-radius: 4px; border: none; }} "
            f"QPushButton:hover {{ background-color: {C.INFO}; }} "
            f"QPushButton:disabled {{ background-color: {C.BORDER_STRONG}; color: {C.TEXT_MUTED}; }}"
        )
        self.btn_mark_next.clicked.connect(self._on_mark_read_and_next)
        act_row.addWidget(self.btn_mark_next)

        self.btn_toggle_read = QPushButton("✓ Read")
        self.btn_toggle_read.setToolTip("Mark selected item(s) Read / Unread (syncs with Outlook).")
        self.btn_toggle_read.clicked.connect(self._on_toggle_read)
        act_row.addWidget(self.btn_toggle_read)

        self.btn_toggle_flag = QPushButton("🚩")
        self.btn_toggle_flag.setToolTip("Flag / Clear Flag in Outlook.")
        self.btn_toggle_flag.clicked.connect(self._on_toggle_flag)
        act_row.addWidget(self.btn_toggle_flag)

        self.btn_reply = QPushButton("↩️ Reply")
        self.btn_reply.setToolTip("Reply in Outlook.")
        self.btn_reply.clicked.connect(self._on_reply)
        act_row.addWidget(self.btn_reply)

        self.btn_open_outlook = QPushButton("📧 Outlook")
        self.btn_open_outlook.setToolTip("Open email in Outlook.")
        self.btn_open_outlook.clicked.connect(self._on_open_outlook)
        act_row.addWidget(self.btn_open_outlook)

        self.btn_open_context = QPushButton("📂 Thread")
        self.btn_open_context.setToolTip("Open full Tracked Thread or TDoc viewer.")
        self.btn_open_context.clicked.connect(self._on_open_context)
        act_row.addWidget(self.btn_open_context)

        self.btn_sync_selected = QPushButton("🔄 Sync")
        self.btn_sync_selected.setToolTip(
            "Sync only the currently selected email thread(s)/conversation(s) with Outlook asynchronously."
        )
        self.btn_sync_selected.clicked.connect(self._on_sync_selected_clicked)
        act_row.addWidget(self.btn_sync_selected)

        self.btn_merge_selected = QPushButton("🔗 Merge")
        self.btn_merge_selected.setToolTip("Merge selected split threads/conversations into a single tracked thread.")
        self.btn_merge_selected.clicked.connect(self._on_merge_selected_clicked)
        self.btn_merge_selected.setVisible(False)
        self.btn_merge_selected.setEnabled(False)
        act_row.addWidget(self.btn_merge_selected)

        self.detail_status_combo = QComboBox()
        self.detail_status_combo.addItems(MY_STATUS_OPTIONS)
        self.detail_status_combo.setMinimumWidth(scale_px(108))
        self.detail_status_combo.setToolTip("Set thread My Status (color bar shown in thread listing & TDocs table).")
        self.detail_status_combo.currentIndexChanged.connect(self._on_detail_status_changed)
        act_row.addWidget(self.detail_status_combo)

        self.btn_toggle_note = QPushButton("📝 Note")
        self.btn_toggle_note.setCheckable(True)
        self.btn_toggle_note.setToolTip("Show / edit thread comment (auto-shown whenever a comment is present).")
        self.btn_toggle_note.clicked.connect(self._on_toggle_note_clicked)
        act_row.addWidget(self.btn_toggle_note)

        self.btn_toggle_assign_bar = QPushButton("🏷️ AI / WI ▴" if self._show_assign_bar else "🏷️ AI / WI ▾")
        self.btn_toggle_assign_bar.setCheckable(True)
        self.btn_toggle_assign_bar.setChecked(self._show_assign_bar)
        self.btn_toggle_assign_bar.setToolTip("Show or hide the inline AI / Sub-Topic / WI assignment bar.")
        self.btn_toggle_assign_bar.clicked.connect(self._on_toggle_assign_bar)
        act_row.addWidget(self.btn_toggle_assign_bar)

        right_layout.addWidget(act_widget)

        self.preview_browser = QTextBrowser()
        self.preview_browser.setOpenLinks(False)
        apply_email_reader_document_style(self.preview_browser)
        self.preview_browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.preview_browser, self._settings, key="zoom_steps")
        right_layout.addWidget(self.preview_browser, 1)

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(right_widget)
        self._apply_layout_orientation(save_current=False)
        layout.addWidget(self.splitter, 1)
