"""UI layout construction mixin for TDocEmailsDialog."""

import html
from typing import List

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill, scale_px
from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style
from modules.emails.ui.ai_summary import InlineEmailSummaryCard
from modules.emails.ui.meeting_email_inbox_dialog import (
    _NotePlainTextEdit,
    get_compact_action_bar_qss,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS


class TDocEmailsUiSetupMixin:
    """Builds the context card, inline note bar, action ribbon, filters, and reader."""

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Row 1: TDoc Context Card + Always-Editable Inline Comment Strip
        top_card = QFrame(self)
        top_card.setObjectName("TDocEmailHeader")
        top_card.setStyleSheet(
            f"#TDocEmailHeader {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        top_card_layout = QVBoxLayout(top_card)
        top_card_layout.setContentsMargins(8, 5, 8, 5)
        top_card_layout.setSpacing(3)

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(8)
        tdoc_title = str(self._tdoc_info.get("Title") or "").strip()
        heading_text = f"{self.target_tdoc} — {tdoc_title}" if tdoc_title else f"{self.target_tdoc}"
        self.lbl_tdoc_heading = QLabel(heading_text)
        self.lbl_tdoc_heading.setWordWrap(True)
        self.lbl_tdoc_heading.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        title_row.addWidget(self.lbl_tdoc_heading, 1)

        self.lbl_counts = QLabel("0 Total (0 Unread)")
        self.lbl_counts.setStyleSheet(f"font-weight: bold; color: {C.INFO}; font-size: 11px;")
        title_row.addWidget(self.lbl_counts)
        top_card_layout.addLayout(title_row)

        meta_bits: List[str] = []
        src = str(self._tdoc_info.get("Source") or "").strip()
        if src:
            src_short = (src[:58] + "…") if len(src) > 60 else src
            meta_bits.append(f"<b>Source:</b> {html.escape(src_short)}")
        ai_val = str(self._tdoc_info.get("Agenda Item") or "").strip()
        sub_val = str(self._tdoc_info.get("Sub-Item") or "").strip()
        if ai_val and ai_val not in ("-", "N/A"):
            ai_txt = f"<b>AI:</b> {html.escape(ai_val)}"
            if sub_val:
                ai_txt += f" ({html.escape(sub_val)})"
            meta_bits.append(ai_txt)
        wi_val = str(self._tdoc_info.get("Related WIs") or "").strip()
        if wi_val and wi_val not in ("-", "General"):
            meta_bits.append(f"<b>WI:</b> {html.escape(wi_val)}")
        t_status = str(self._tdoc_info.get("TDoc Status") or "").strip()
        if t_status:
            meta_bits.append(f"<b>Status:</b> {html.escape(t_status)}")

        chips_text = " ➔ ".join(
            f"<b>{html.escape(t)}</b>"
            if t == self.target_tdoc
            else f'<a href="tdoc:{html.escape(t)}" style="color:{C.INFO}; text-decoration:underline;">{html.escape(t)}</a>'
            for t in self.family_tdocs
        )
        meta_bits.append(f"<b>Family:</b> {chips_text}")

        self.lbl_tdoc_meta = QLabel(" &nbsp;·&nbsp; ".join(meta_bits))
        self.lbl_tdoc_meta.setWordWrap(True)
        self.lbl_tdoc_meta.setTextFormat(Qt.RichText)
        self.lbl_tdoc_meta.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.lbl_tdoc_meta.setOpenExternalLinks(False)
        self.lbl_tdoc_meta.linkActivated.connect(lambda href: self._on_anchor_clicked(QUrl(href)))
        self.lbl_tdoc_meta.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        if src and len(src) > 60:
            self.lbl_tdoc_meta.setToolTip(f"Source: {src}")
        top_card_layout.addWidget(self.lbl_tdoc_meta)

        # Always-Editable Comment Strip (auto-visible when My Notes is non-empty, toggled via 📝 Note)
        self.note_frame = QFrame(top_card)
        self.note_frame.setObjectName("TDocNoteBar")
        self.note_frame.setStyleSheet(
            f"#TDocNoteBar {{ background-color: {C.WARNING_SOFT}; border: 1px solid {C.WARNING_BORDER}; border-radius: 5px; }}"
        )
        note_row = QHBoxLayout(self.note_frame)
        note_row.setContentsMargins(6, 3, 6, 3)
        note_row.setSpacing(6)
        note_lbl = QLabel("<b>📝 Note:</b>")
        note_lbl.setStyleSheet(f"color: {C.WARNING};")
        note_row.addWidget(note_lbl)

        self.note_edit = _NotePlainTextEdit(self.note_frame)
        self.note_edit.setPlaceholderText("Add TDoc comment / note (auto-saved on blur or Ctrl+Enter)...")
        self.note_edit.setMaximumHeight(scale_px(44))
        self.note_edit.setMinimumHeight(scale_px(26))
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
        top_card_layout.addWidget(self.note_frame)
        layout.addWidget(top_card)

        # Row 2: Compact Action Bar (mirrors MeetingEmailInboxDialog's reading panel bar)
        toolbar_widget = QWidget(self)
        toolbar_widget.setObjectName("TDocActionBar")
        toolbar_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        toolbar_sp.setHeightForWidth(True)
        toolbar_widget.setSizePolicy(toolbar_sp)
        toolbar_widget.setStyleSheet(get_compact_action_bar_qss("TDocActionBar"))
        toolbar = FlowLayout(toolbar_widget, margin=0, h_spacing=4, v_spacing=3)

        self.next_unread_btn = QPushButton("🔵 Next Unread (0)")
        self.next_unread_btn.setToolTip("Jump to the next unread email in the discussion timeline.")
        self.next_unread_btn.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_ALT}; color: {C.TEXT_ON_ACCENT}; font-weight: bold; "
            f"padding: 3px 10px; border-radius: 4px; border: none; }} "
            f"QPushButton:hover {{ background-color: {C.INFO}; }} "
            f"QPushButton:disabled {{ background-color: {C.BORDER_STRONG}; color: {C.TEXT_MUTED}; font-weight: normal; }}"
        )
        self.next_unread_btn.clicked.connect(self._on_next_unread_clicked)
        toolbar.addWidget(self.next_unread_btn)

        self.btn_mark_all_read = QPushButton("✓ Read All")
        self.btn_mark_all_read.setToolTip("Mark all emails for this TDoc family as read in local DB and Outlook.")
        self.btn_mark_all_read.clicked.connect(self._mark_all_family_read)
        toolbar.addWidget(self.btn_mark_all_read)

        self.btn_reload_outlook = QPushButton("🔄 Sync")
        self.btn_reload_outlook.setToolTip(
            "Re-fetch these cached emails from Microsoft Outlook to populate HTML tables, attachments, and read/flag state."
        )
        self.btn_reload_outlook.clicked.connect(self._on_reload_all_from_outlook)
        toolbar.addWidget(self.btn_reload_outlook)

        self.btn_open_tdoc = QPushButton(f"📄 Open {self.target_tdoc}")
        self.btn_open_tdoc.setToolTip(f"Download/open {self.target_tdoc} document.")
        self.btn_open_tdoc.setVisible(self._open_tdoc_callback is not None)
        self.btn_open_tdoc.clicked.connect(self._on_open_tdoc_clicked)
        toolbar.addWidget(self.btn_open_tdoc)

        self.btn_pin_thread = QPushButton("📌 Pin Thread")
        self.btn_pin_thread.setToolTip("Promote this email discussion into a persistent Tracked Email Thread in the TDocs table.")
        self.btn_pin_thread.clicked.connect(self._on_pin_as_tracked_thread)
        self.btn_pin_thread.setVisible(self._pin_thread_callback is not None)
        toolbar.addWidget(self.btn_pin_thread)

        self.detail_status_combo = QComboBox()
        self.detail_status_combo.addItems(MY_STATUS_OPTIONS)
        self.detail_status_combo.setMinimumWidth(scale_px(108))
        self.detail_status_combo.setToolTip(f"Set My Status for {self.target_tdoc}.")
        self.detail_status_combo.currentIndexChanged.connect(self._on_tdoc_status_changed)
        toolbar.addWidget(self.detail_status_combo)

        self.btn_toggle_note = QPushButton("📝 Note")
        self.btn_toggle_note.setCheckable(True)
        self.btn_toggle_note.setToolTip("Show / edit TDoc comment (auto-shown whenever a comment is present).")
        self.btn_toggle_note.clicked.connect(self._on_toggle_note_clicked)
        toolbar.addWidget(self.btn_toggle_note)

        self.llm_export_btn = QPushButton("🤖 LLM && Export")
        self.llm_export_btn.setToolTip("AI Summarize Discussion (Ollama), Launch LLM CLI, or Copy/Save Markdown.")
        self.llm_export_menu = QMenu(self)
        self.llm_export_menu.aboutToShow.connect(self._populate_llm_export_menu)
        self.llm_export_btn.setMenu(self.llm_export_menu)
        toolbar.addWidget(self.llm_export_btn)

        docs_btn = QPushButton("📂 Docs")
        docs_menu = QMenu(self)
        act_cache_docs = docs_menu.addAction("📥 Cache All Linked Docs && Attachments")
        act_cache_docs.triggered.connect(self._on_cache_all_docs)
        act_open_folder = docs_menu.addAction("📂 Open Cached Docs Folder")
        act_open_folder.triggered.connect(self._on_open_cache_folder)
        docs_btn.setMenu(docs_menu)
        toolbar.addWidget(docs_btn)

        more_btn = QPushButton("⚙️ More")
        more_menu = QMenu(self)
        self.dedup_check = QAction("Deduplicate trailing quotes", self)
        self.dedup_check.setCheckable(True)
        self.dedup_check.setChecked(True)
        self.dedup_check.toggled.connect(self._render_messages)
        more_menu.addAction(self.dedup_check)

        self.show_meta_action = QAction("Show Metadata && Linked Docs Header", self)
        self.show_meta_action.setCheckable(True)
        self.show_meta_action.setChecked(self._settings.value("show_metadata_header", False, type=bool))
        self.show_meta_action.toggled.connect(self._on_show_metadata_toggled)
        more_menu.addAction(self.show_meta_action)

        self.export_chrono_check = QAction("Export Oldest First (for LLMs)", self)
        self.export_chrono_check.setCheckable(True)
        self.export_chrono_check.setChecked(self._settings.value("export_oldest_first", True, type=bool))
        self.export_chrono_check.toggled.connect(lambda v: self._settings.setValue("export_oldest_first", bool(v)))

        more_btn.setMenu(more_menu)
        toolbar.addWidget(more_btn)
        layout.addWidget(toolbar_widget)

        # Row 3: Search, Company, Read/Flag State, Display Order & Compact Scope Pills
        filter_bar = QWidget(self)
        filter_bar.setStyleSheet(".QWidget { background: transparent; border: none; }")
        filter_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        filter_sp.setHeightForWidth(True)
        filter_bar.setSizePolicy(filter_sp)
        filter_row = FlowLayout(filter_bar, margin=0, h_spacing=6, v_spacing=4)

        self.msg_search_edit = QLineEdit(filter_bar)
        self.msg_search_edit.setClearButtonEnabled(True)
        self.msg_search_edit.setMinimumWidth(scale_px(190))
        self.msg_search_edit.setMinimumHeight(scale_px(26))
        self.msg_search_edit.setPlaceholderText("🔍 Filter messages by keyword, TDoc, or sender...")
        self.msg_search_edit.textChanged.connect(self._render_messages)
        filter_row.addWidget(self.msg_search_edit)

        self.company_filter_combo = CheckableComboBox("Company")
        self.company_filter_combo.setMinimumWidth(scale_px(125))
        self.company_filter_combo.setMaximumWidth(scale_px(180))
        self.company_filter_combo.setMinimumHeight(scale_px(26))
        self.company_filter_combo.setToolTip("Filter messages by one or more sender companies.")
        self.company_filter_combo.selectionChanged.connect(self._render_messages)
        filter_row.addWidget(self.company_filter_combo)

        self.state_filter_combo = QComboBox(filter_bar)
        self.state_filter_combo.setMinimumHeight(scale_px(26))
        self.state_filter_combo.addItem("All Messages", "all")
        self.state_filter_combo.addItem("🔵 Unread Only", "unread")
        self.state_filter_combo.addItem("🚩 Flagged Only", "flagged")
        self.state_filter_combo.setToolTip("Filter messages by read or Outlook flag state.")
        self.state_filter_combo.currentIndexChanged.connect(self._render_messages)
        filter_row.addWidget(self.state_filter_combo)

        self.order_combo = QComboBox(filter_bar)
        self.order_combo.setMinimumHeight(scale_px(26))
        self.order_combo.addItem("⬇️ Newest", True)
        self.order_combo.addItem("⬆️ Oldest", False)
        saved_newest = self._settings.value("newest_first", True, type=bool)
        self.order_combo.setCurrentIndex(0 if saved_newest else 1)
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        filter_row.addWidget(self.order_combo)

        scope_pill_qss = (
            f"QPushButton {{ padding: 2px 7px; border-radius: 10px; border: 1px solid {C.BORDER_STRONG}; "
            f"background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; font-size: 11px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; }} "
            f"QPushButton:checked {{ background-color: {C.INFO_SOFT}; border-color: {C.ACCENT_ALT}; color: {C.INFO}; font-weight: bold; }}"
        )
        pill_h = scale_px(24)
        self.chk_family = QPushButton("🔗 Family", filter_bar)
        self.chk_family.setCheckable(True)
        self.chk_family.setChecked(True)
        self.chk_family.setMinimumHeight(pill_h)
        self.chk_family.setStyleSheet(scope_pill_qss)
        self.chk_family.setToolTip("Include emails mentioning any predecessor or revision in this TDoc family.")
        self.chk_family.toggled.connect(self._load_emails)
        filter_row.addWidget(self.chk_family)

        self.chk_include_quoted = QPushButton("💬 Quoted", filter_bar)
        self.chk_include_quoted.setCheckable(True)
        self.chk_include_quoted.setChecked(True)
        self.chk_include_quoted.setMinimumHeight(pill_h)
        self.chk_include_quoted.setStyleSheet(scope_pill_qss)
        self.chk_include_quoted.setToolTip("Include emails where the TDoc was only cited inside quoted reply history.")
        self.chk_include_quoted.toggled.connect(self._load_emails)
        filter_row.addWidget(self.chk_include_quoted)

        self.chk_show_ignored = QPushButton("🚫 Ignored", filter_bar)
        self.chk_show_ignored.setCheckable(True)
        self.chk_show_ignored.setChecked(False)
        self.chk_show_ignored.setMinimumHeight(pill_h)
        self.chk_show_ignored.setStyleSheet(scope_pill_qss)
        self.chk_show_ignored.setToolTip("Show emails previously marked as Ignored for this TDoc.")
        self.chk_show_ignored.toggled.connect(self._load_emails)
        filter_row.addWidget(self.chk_show_ignored)

        self.msg_count_lbl = QLabel("", filter_bar)
        self.msg_count_lbl.setMinimumHeight(pill_h)
        self.msg_count_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        filter_row.addWidget(self.msg_count_lbl)

        self.help_pill_btn = create_help_tips_pill(
            [
                "Ctrl + Mouse Wheel in reading pane: zoom text larger/smaller (Ctrl+0 to reset; saved automatically)",
                "🔵 Next Unread jumps through unread messages in the timeline",
                "Click TDoc links or [📎 ...] attachments to open documents directly",
            ],
            parent=self,
            title="TDoc Emails Tips",
        )
        filter_row.addWidget(self.help_pill_btn)
        layout.addWidget(filter_bar)

        # Initialize My Status & Comment bar from passed TDoc metadata
        self._populate_tdoc_status_and_note()

        # Row 4: Collapsible AI Summary Card + Continuous Timeline Browser
        self.reader_splitter = QSplitter(Qt.Vertical, self)
        self.reader_splitter.setChildrenCollapsible(False)

        self.summary_card = InlineEmailSummaryCard(self.reader_splitter, settings_key="tdoc_emails_summary_expanded")
        self.summary_card.anchor_clicked.connect(self._on_anchor_clicked)
        self.summary_card.append_to_notes_requested.connect(self._on_append_summary_to_notes)

        self.browser = QTextBrowser(self.reader_splitter)
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.browser, self._settings, key="zoom_steps")
        self.reader_splitter.setSizes([190, 480])
        layout.addWidget(self.reader_splitter, 1)
