"""UI construction mixin for MeetingEmailInboxDialog."""

from __future__ import annotations

from PyQt5.QtCore import Qt, QTimer, QUrl
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style
from modules.emails.ui.ai_summary import InlineEmailSummaryCard
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS
from modules.meetings.ui.topic_pills import TagCloudWidget
from ..widgets import (
    _InboxStatusStripDelegate,
    _InboxTableWidget,
    _NotePlainTextEdit,
    build_inbox_filter_bar,
    build_inbox_top_toolbar,
    get_compact_action_bar_qss,
    get_inbox_toolbar_button_qss,
)


class InboxUiSetupMixin:
    """Builds the toolbar, filter bar, inbox table, detail header, and reader panes."""

    def _clear_all_filters(self):
        """Resets search input, Hide Objected toggle, and all facet filter dropdowns in one click."""
        self._restoring_filters = True
        try:
            if hasattr(self, "search_edit"):
                self.search_edit.blockSignals(True)
                self.search_edit.clear()
                self.search_edit.blockSignals(False)
            self._hide_objected = False
            if hasattr(self, "btn_hide_objected"):
                self.btn_hide_objected.blockSignals(True)
                self.btn_hide_objected.setChecked(False)
                self.btn_hide_objected.blockSignals(False)
            for combo_attr in ("ai_combo", "company_combo", "status_combo"):
                combo = getattr(self, combo_attr, None)
                if combo is not None:
                    combo.blockSignals(True)
                    combo.setCheckedItems([])
                    combo.blockSignals(False)
            if hasattr(self, "source_combo"):
                self.source_combo.blockSignals(True)
                self.source_combo.setCurrentIndex(0)
                self.source_combo.blockSignals(False)
        finally:
            self._restoring_filters = False
        self._apply_filters()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 8)
        layout.setSpacing(5)

        layout.addWidget(build_inbox_top_toolbar(self))
        layout.addWidget(build_inbox_filter_bar(self))

        # Master-Detail Splitter
        self.splitter = QSplitter(Qt.Horizontal, self)
        self.splitter.setChildrenCollapsible(False)

        left_widget = QWidget()
        self._left_widget = left_widget
        left_widget.setMinimumWidth(240)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        # Slim Contextual Per-AI Briefing Strip (visible ONLY when exactly 1 concrete AI is filtered)
        self.ai_briefing_strip = QFrame(left_widget)
        self.ai_briefing_strip.setObjectName("InboxAiBriefingStrip")
        self.ai_briefing_strip.setStyleSheet(
            f"#InboxAiBriefingStrip {{ background-color: {C.INFO_SOFT}; border: 1px solid {C.INFO_BORDER}; border-radius: 5px; }}"
        )
        ai_strip_lay = QHBoxLayout(self.ai_briefing_strip)
        ai_strip_lay.setContentsMargins(8, 3, 8, 3)
        ai_strip_lay.setSpacing(6)
        self.ai_briefing_strip_lbl = QLabel("", self.ai_briefing_strip)
        self.ai_briefing_strip_lbl.setStyleSheet(f"color: {C.INFO_DARK}; font-size: 11px; font-weight: 600;")
        ai_strip_lay.addWidget(self.ai_briefing_strip_lbl, 1)
        self.btn_ai_strip_open = QPushButton("📖 Open", self.ai_briefing_strip)
        self.btn_ai_strip_open.setCursor(Qt.PointingHandCursor)
        self.btn_ai_strip_open.setStyleSheet(get_inbox_toolbar_button_qss(compact=True))
        self.btn_ai_strip_open.clicked.connect(lambda _=False: self._on_ai_strip_action(auto_run=False))
        ai_strip_lay.addWidget(self.btn_ai_strip_open)
        self.btn_ai_strip_run = QPushButton("⚡ Update", self.ai_briefing_strip)
        self.btn_ai_strip_run.setCursor(Qt.PointingHandCursor)
        self.btn_ai_strip_run.setStyleSheet(get_inbox_toolbar_button_qss(compact=True))
        self.btn_ai_strip_run.clicked.connect(lambda _=False: self._on_ai_strip_action(auto_run=True))
        ai_strip_lay.addWidget(self.btn_ai_strip_run)
        self.ai_briefing_strip.setVisible(False)
        left_layout.addWidget(self.ai_briefing_strip)

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

        # Compact Clickable In-Thread Message Jump Strip (auto-shown for multi-message Threads / TDocs)
        self.detail_jump_strip_lbl = QLabel("")
        self.detail_jump_strip_lbl.setWordWrap(True)
        self.detail_jump_strip_lbl.setTextFormat(Qt.RichText)
        self.detail_jump_strip_lbl.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.detail_jump_strip_lbl.setOpenExternalLinks(False)
        self.detail_jump_strip_lbl.linkActivated.connect(lambda href: self._on_anchor_clicked(QUrl(href)))
        self.detail_jump_strip_lbl.setStyleSheet(
            f"background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-radius: 4px; padding: 2px 6px; font-size: 11px;"
        )
        self.detail_jump_strip_lbl.setVisible(False)
        meta_layout.addWidget(self.detail_jump_strip_lbl)

        # Inline Assignment Bar (Single compact row: AIs, Sub-Topics, WIs + Auto/Pin; persisted in QSettings)
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
        self.btn_auto_assign.setCursor(Qt.PointingHandCursor)
        self.btn_auto_assign.setStyleSheet(get_inbox_toolbar_button_qss(compact=True))
        self.btn_auto_assign.setToolTip("Infer and apply AIs, Sub-Topics, and WIs from subject and linked TDocs.")
        self.btn_auto_assign.clicked.connect(self._on_auto_assign_clicked)
        assign_row.addWidget(self.btn_auto_assign)

        self.btn_pin_inline = QPushButton("📌 Pin")
        self.btn_pin_inline.setCursor(Qt.PointingHandCursor)
        self.btn_pin_inline.setStyleSheet(get_inbox_toolbar_button_qss(compact=True))
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

        self.btn_toggle_objected = QPushButton("🛑 Objected")
        self.btn_toggle_objected.setCheckable(True)
        self.btn_toggle_objected.setToolTip(
            "Mark this thread/TDoc as 🛑 Objected (Parked out of Priority Radar unless a newer revision r01/r02 arrives)."
        )
        self.btn_toggle_objected.clicked.connect(self._on_toggle_objected_clicked)
        act_row.addWidget(self.btn_toggle_objected)

        self.btn_reply = QPushButton("↩️ Reply")
        self.btn_reply.setToolTip("Reply in Outlook.")
        self.btn_reply.clicked.connect(self._on_reply)
        act_row.addWidget(self.btn_reply)

        self.btn_copy_start_template = QPushButton("📋 <<START>>")
        self.btn_copy_start_template.setToolTip(
            "Copy a pre-formatted 3GPP e-Meeting <<START>> ... <<END>> formal comment block to the clipboard."
        )
        self.btn_copy_start_template.clicked.connect(self._on_copy_start_template_clicked)
        act_row.addWidget(self.btn_copy_start_template)

        self.btn_open_outlook = QPushButton("📧 Outlook")
        self.btn_open_outlook.setToolTip("Open email in Outlook.")
        self.btn_open_outlook.clicked.connect(self._on_open_outlook)
        act_row.addWidget(self.btn_open_outlook)

        self.btn_open_context = QPushButton("📂 Thread")
        self.btn_open_context.setToolTip("Open full Tracked Thread or TDoc viewer (or latest cached /Drafts/ doc).")
        self.btn_open_context.clicked.connect(self._on_open_context)
        act_row.addWidget(self.btn_open_context)

        self.btn_jump_tdoc_table = QPushButton("🎯 Table")
        self.btn_jump_tdoc_table.setToolTip("Scroll to and highlight this TDoc in the TDocsWindow meeting table.")
        self.btn_jump_tdoc_table.clicked.connect(self._on_jump_tdoc_table_clicked)
        self.btn_jump_tdoc_table.setVisible(False)
        act_row.addWidget(self.btn_jump_tdoc_table)

        self.btn_llm_export = QPushButton("🤖 LLM && Export")
        self.btn_llm_export.setToolTip("Copy pure Markdown for LLMs or launch local AI summary.")
        self.btn_llm_export.clicked.connect(self._on_llm_export_clicked)
        act_row.addWidget(self.btn_llm_export)

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

        self.detail_reader_splitter = QSplitter(Qt.Vertical, right_widget)
        self.detail_reader_splitter.setChildrenCollapsible(False)

        self.summary_card = InlineEmailSummaryCard(self.detail_reader_splitter, settings_key="inbox_summary_expanded")
        self.summary_card.anchor_clicked.connect(self._on_anchor_clicked)
        self.summary_card.append_to_notes_requested.connect(self._on_append_summary_to_notes)
        self.summary_card.summary_saved.connect(self._on_inline_summary_saved)

        self.preview_browser = QTextBrowser(self.detail_reader_splitter)
        self.preview_browser.setOpenLinks(False)
        apply_email_reader_document_style(self.preview_browser)
        self.preview_browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.preview_browser, self._settings, key="zoom_steps")
        self.detail_reader_splitter.setSizes([190, 420])
        right_layout.addWidget(self.detail_reader_splitter, 1)

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(right_widget)
        self._apply_layout_orientation(save_current=False)
        layout.addWidget(self.splitter, 1)
