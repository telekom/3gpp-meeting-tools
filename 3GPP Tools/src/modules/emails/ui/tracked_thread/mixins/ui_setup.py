"""UI layout and menu construction mixin for TrackedThreadDialog."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QFrame,
    QGridLayout,
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

from core.ai.cli_detector import detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import open_mcp_agent_setup_dialog
from core.ai.llm_runner import resolve_effective_model_spec
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill, scale_px
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style
from modules.emails.ui.ai_summary import InlineEmailSummaryCard, populate_summary_model_menu
from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.topic_pills import TagCloudWidget
from core.ui.theme import C


class TrackedThreadUiSetupMixin:
    """Toolbar, links panel, filter bar, reader pane, and export menu construction."""

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Row 1: Compact title, inline stats, and collapsible Links & Tags toggle
        header_row = QHBoxLayout()
        header_row.setSpacing(8)
        self.title_label = QLabel("Email Thread")
        self.title_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.title_label.setWordWrap(True)
        header_row.addWidget(self.title_label, 1)

        self.meta_label = QLabel("")
        self.meta_label.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        header_row.addWidget(self.meta_label)

        self.links_toggle_btn = QPushButton("🏷️ Links && Tags ▸")
        self.links_toggle_btn.setCheckable(True)
        saved_links_open = self._settings.value("links_panel_expanded", False, type=bool)
        self.links_toggle_btn.setChecked(saved_links_open)
        self.links_toggle_btn.setToolTip("Show or hide AIs, Sub-Topics, WIs, TDocs, and Personal Tags editors.")
        self.links_toggle_btn.toggled.connect(self._on_toggle_links_panel)
        header_row.addWidget(self.links_toggle_btn)
        layout.addLayout(header_row)

        # Collapsible compact 2-column Links & Tags panel
        self.links_panel = QFrame(self)
        self.links_panel.setObjectName("LinksPanel")
        self.links_panel.setStyleSheet(
            f"#LinksPanel {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        links_grid = QGridLayout(self.links_panel)
        links_grid.setContentsMargins(8, 6, 8, 6)
        links_grid.setHorizontalSpacing(10)
        links_grid.setVerticalSpacing(4)
        links_grid.setColumnStretch(1, 1)
        links_grid.setColumnStretch(3, 1)

        links_grid.addWidget(QLabel("<b>AIs:</b>"), 0, 0, Qt.AlignVCenter)
        self.ais_pills = TagCloudWidget(
            placeholder="AI # (e.g. 20.6.1)...",
            pill_colors={"bg": C.INFO_SOFT, "fg": C.INFO, "border": C.INFO_BORDER},
            strict_dropdown=True,
            dropdown_label="+ Add AI ▾",
            empty_dropdown_label="No meeting AIs",
        )
        self.ais_pills.set_suggestions(self._ai_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        self.ais_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.ais_pills, 0, 1)

        links_grid.addWidget(QLabel("<b>Sub-Topics:</b>"), 0, 2, Qt.AlignVCenter)
        self.sub_items_pills = TagCloudWidget(
            placeholder="AI Sub-Topic...",
            pill_colors={"bg": C.TEAL_SURFACE, "fg": C.TEAL_DARK, "border": C.TEAL_BORDER},
            strict_dropdown=True,
            dropdown_label="+ Add Sub-Topic ▾",
            empty_dropdown_label="Assign AI first",
        )
        self.sub_items_pills.set_suggestions([], value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        self.sub_items_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.sub_items_pills, 0, 3)

        links_grid.addWidget(QLabel("<b>WIs:</b>"), 1, 0, Qt.AlignVCenter)
        self.wis_pills = TagCloudWidget(
            placeholder="Work Item (e.g. FS_6G_ARC)...",
            pill_colors={"bg": C.WG_SA_BG, "fg": C.INDIGO_DARK, "border": C.WG_SA_BORDER},
            strict_dropdown=True,
            dropdown_label="+ Add WI ▾",
            empty_dropdown_label="No meeting WIs",
        )
        self.wis_pills.set_suggestions(self._wi_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        self.wis_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.wis_pills, 1, 1)

        links_grid.addWidget(QLabel("<b>TDocs:</b>"), 1, 2, Qt.AlignVCenter)
        self.tdocs_pills = TagCloudWidget(
            placeholder="TDoc # (auto-expands revs)...",
            pill_colors={"bg": C.WG_RAN_BG, "fg": C.EMERALD_DARK, "border": C.WG_RAN_BORDER},
        )
        self.tdocs_pills.set_suggestions(self._tdoc_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip().upper())
        self.tdocs_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.tdocs_pills, 1, 3)

        links_grid.addWidget(QLabel("<b>Tags:</b>"), 2, 0, Qt.AlignVCenter)
        self.tags_pills = TagCloudWidget(
            placeholder="Personal tag (e.g. Key Issue, Follow-up)...",
            pill_colors={"bg": C.WARNING_SOFT_ALT, "fg": C.WARNING_DARK, "border": C.WARNING_BORDER_STRONG},
        )
        self.tags_pills.set_suggestions(self._all_existing_tags)
        self.tags_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.tags_pills, 2, 1, 1, 3)

        self.links_panel.setVisible(saved_links_open)
        layout.addWidget(self.links_panel)

        # Row 2: Lean action toolbar with grouped menus
        toolbar_widget = QWidget(self)
        toolbar_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        toolbar_sp.setHeightForWidth(True)
        toolbar_widget.setSizePolicy(toolbar_sp)
        toolbar = FlowLayout(toolbar_widget, margin=0, h_spacing=6, v_spacing=4)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setToolTip("Sync new replies and Outlook read/flag states for this thread in the background.")
        self.refresh_btn.clicked.connect(self._on_refresh)
        toolbar.addWidget(self.refresh_btn)

        read_btn = QPushButton("✔️ Mark All Read")
        read_btn.setToolTip("Mark all messages in this thread as read (and sync to Outlook).")
        read_btn.clicked.connect(self._on_mark_read)
        toolbar.addWidget(read_btn)

        self.notes_btn = QPushButton("📝 Notes && Status...")
        self.notes_btn.setToolTip("Edit personal stance (My Status), Tags, and Delegate Notes for this email thread.")
        self.notes_btn.clicked.connect(self._on_edit_notes_and_status)
        toolbar.addWidget(self.notes_btn)

        self.llm_export_btn = QPushButton("🤖 LLM && Export")
        self.llm_export_btn.setToolTip("Launch an external LLM CLI (prompt staged on clipboard) or export Markdown.")
        self.llm_export_menu = QMenu(self)
        self.llm_export_menu.aboutToShow.connect(self._populate_llm_export_menu)
        self.llm_export_btn.setMenu(self.llm_export_menu)
        toolbar.addWidget(self.llm_export_btn)

        docs_btn = QPushButton("📂 Docs")
        docs_menu = QMenu(self)
        act_cache_docs = docs_menu.addAction("📥 Cache All Linked Docs && Attachments")
        act_cache_docs.triggered.connect(self._on_cache_all_docs)
        act_open_folder = docs_menu.addAction("📂 Open Thread Cache Folder")
        act_open_folder.triggered.connect(self._on_open_thread_folder)
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

        more_menu.addSeparator()
        act_reload_all = more_menu.addAction("🔄 Reload All Cached Emails (Tables && Links)")
        act_reload_all.setToolTip("Re-fetch cached Outlook messages by EntryID to update HTML tables, links, and read/flag states.")
        act_reload_all.triggered.connect(self._on_reload_all_cached)
        act_reload_msg = more_menu.addAction("📥 Reload / Merge .msg File(s)...")
        act_reload_msg.setToolTip("Load a local Outlook .msg file into this thread (updates matching messages or adds replies).")
        act_reload_msg.triggered.connect(self._on_reload_from_msg_file)
        more_menu.addSeparator()
        act_merge = more_menu.addAction("🔗 Merge Split Threads...")
        act_merge.setToolTip("Search and merge one or more split threads from this meeting into this thread.")
        act_merge.triggered.connect(self._on_merge_thread)
        act_unpin = more_menu.addAction("📌 Unpin Email Thread")
        act_unpin.triggered.connect(self._on_delete_thread)
        more_btn.setMenu(more_menu)
        toolbar.addWidget(more_btn)

        layout.addWidget(toolbar_widget)

        self.sync_bar = AsyncSyncProgressBar(self)
        self.sync_bar.cancel_requested.connect(self._on_cancel_refresh)
        layout.addWidget(self.sync_bar)

        # Row 3: Search, Company/State Filters & Display Order
        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        self.msg_search_edit = QLineEdit()
        self.msg_search_edit.setClearButtonEnabled(True)
        self.msg_search_edit.setPlaceholderText("🔍 Filter messages by keyword, TDoc, or sender...")
        self.msg_search_edit.textChanged.connect(self._render_messages)
        filter_row.addWidget(self.msg_search_edit, 1)

        self.company_filter_combo = CheckableComboBox("Company")
        self.company_filter_combo.setMinimumWidth(scale_px(130))
        self.company_filter_combo.setMaximumWidth(scale_px(190))
        self.company_filter_combo.setToolTip("Filter messages by one or more sender companies.")
        self.company_filter_combo.selectionChanged.connect(self._render_messages)
        filter_row.addWidget(self.company_filter_combo)

        self.state_filter_combo = QComboBox()
        self.state_filter_combo.addItem("All Messages", "all")
        self.state_filter_combo.addItem("🔵 Unread Only", "unread")
        self.state_filter_combo.addItem("🚩 Flagged Only", "flagged")
        self.state_filter_combo.setToolTip("Filter messages by read or Outlook flag state.")
        self.state_filter_combo.currentIndexChanged.connect(self._render_messages)
        filter_row.addWidget(self.state_filter_combo)

        self.order_combo = QComboBox()
        self.order_combo.setToolTip("Choose display order of messages within this thread.")
        self.order_combo.addItem("⬇️ Newest First", True)
        self.order_combo.addItem("⬆️ Oldest First", False)
        saved_newest = self._settings.value("newest_first", True, type=bool)
        self.order_combo.setCurrentIndex(0 if saved_newest else 1)
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        filter_row.addWidget(self.order_combo)

        self.next_unread_btn = QPushButton("🔵 Next Unread (0)")
        self.next_unread_btn.setToolTip("Jump to the next unread email in the thread timeline.")
        self.next_unread_btn.setStyleSheet(
            f"QPushButton {{ padding: 3px 9px; border-radius: 4px; border: 1px solid {C.INFO_BORDER_MEDIUM}; "
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.INFO_SOFT_STRONG}; }} "
            f"QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; font-weight: normal; }}"
        )
        self.next_unread_btn.clicked.connect(self._on_next_unread_clicked)
        filter_row.addWidget(self.next_unread_btn)

        self.msg_count_lbl = QLabel("")
        self.msg_count_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        filter_row.addWidget(self.msg_count_lbl)

        self.help_pill_btn = create_help_tips_pill(
            [
                "Ctrl + Mouse Wheel in reading pane: zoom text larger/smaller (Ctrl+0 to reset; saved automatically)",
                "🔵 Next Unread jumps through unread replies in timeline order",
                "Click TDoc links or [📎 ...] attachments to open documents directly",
            ],
            parent=self,
            title="Thread Viewer Tips",
        )
        filter_row.addWidget(self.help_pill_btn)
        layout.addLayout(filter_row)

        # Collapsible AI Summary Card + Timeline browser in vertical splitter
        self.reader_splitter = QSplitter(Qt.Vertical, self)
        self.reader_splitter.setChildrenCollapsible(False)

        self.summary_card = InlineEmailSummaryCard(self.reader_splitter, settings_key="tracked_thread_summary_expanded")
        self.summary_card.anchor_clicked.connect(self._on_anchor_clicked)
        self.summary_card.append_to_notes_requested.connect(self._on_append_summary_to_notes)

        self.browser = QTextBrowser(self.reader_splitter)
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.browser, self._settings, key="zoom_steps")
        self.reader_splitter.setSizes([190, 480])
        layout.addWidget(self.reader_splitter, 1)

    def _on_order_changed(self):
        self._settings.setValue("newest_first", bool(self.order_combo.currentData()))
        self._render_messages()

    def _on_show_metadata_toggled(self, checked: bool):
        self._settings.setValue("show_metadata_header", bool(checked))
        self._render_messages()

    def _on_toggle_links_panel(self, expanded: bool):
        self._settings.setValue("links_panel_expanded", bool(expanded))
        if hasattr(self, "links_panel"):
            self.links_panel.setVisible(bool(expanded))
        self._update_links_toggle_label(self.db.get_thread(self.thread_id) or {})

    def _update_links_toggle_label(self, thread: dict):
        if not hasattr(self, "links_toggle_btn"):
            return
        total_items = sum(
            len(thread.get(k) or [])
            for k in ("linked_ais", "linked_sub_items", "linked_wis", "linked_tdocs", "user_tags")
        )
        arrow = "▾" if self.links_toggle_btn.isChecked() else "▸"
        suffix = f" ({total_items})" if total_items > 0 else ""
        self.links_toggle_btn.setText(f"🏷️ Links && Tags{suffix} {arrow}")
        parts = []
        for label, key in (
            ("AIs", "linked_ais"),
            ("Sub-Topics", "linked_sub_items"),
            ("WIs", "linked_wis"),
            ("TDocs", "linked_tdocs"),
            ("Tags", "user_tags"),
        ):
            vals = thread.get(key) or []
            if vals:
                parts.append(f"{label}: {', '.join(vals)}")
        self.links_toggle_btn.setToolTip("\n".join(parts) if parts else "Show or hide AIs, Sub-Topics, WIs, TDocs, and Personal Tags editors.")

    def _populate_llm_export_menu(self):
        menu = self.llm_export_menu
        menu.clear()
        menu.setToolTipsVisible(True)

        if hasattr(self, "summary_card"):
            spec = self.summary_card._active_spec()
            fresh = self.summary_card.current_freshness()
            if fresh and fresh.has_summary and fresh.new_count > 0:
                sum_lbl = f"⚡ Update Summary (+{fresh.new_count}) [{spec.short_badge()}]"
            elif fresh and fresh.has_summary:
                sum_lbl = f"🔄 Refresh Summary [{spec.short_badge()}]"
            else:
                sum_lbl = f"✨ Summarize Thread [{spec.short_badge()}]"
            act_sum = menu.addAction(sum_lbl)
            act_sum.triggered.connect(lambda _=False: self.summary_card.trigger_summarize(mode="auto"))
            if fresh and fresh.has_summary:
                act_reb = menu.addAction("🔄 Full Rebuild Summary")
                act_reb.triggered.connect(lambda _=False: self.summary_card.trigger_summarize(mode="full"))
            model_sub = menu.addMenu("⚙️ Summary Model")
            populate_summary_model_menu(
                model_sub,
                role="leaf",
                active_spec=spec,
                on_selected=lambda s: setattr(self.summary_card, "_model_override", s),
                parent=self,
            )
            menu.addSeparator()

        try:
            statuses = detect_cli_agents()
        except Exception:
            statuses = {}
        installed = [st for st in statuses.values() if st.installed]
        if installed:
            for st in installed:
                badge = format_active_model_badge(st.agent_id)
                act = menu.addAction(f"{st.icon} Launch {st.label} [{badge}]")
                act.setToolTip(
                    f"Prepare thread.md + converted .docx drafts and open {st.label} idle "
                    "(prompt copied to clipboard for Ctrl+V — zero auto token burn)."
                )
                act.triggered.connect(lambda _=False, aid=st.agent_id: self._on_launch_cli(aid))
        else:
            no_cli = menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
            no_cli.setEnabled(False)

        menu.addSeparator()
        act_copy = menu.addAction("📋 Copy Thread Markdown")
        act_copy.triggered.connect(self._on_copy_markdown)
        act_save = menu.addAction("💾 Save Thread .md...")
        act_save.triggered.connect(self._on_save_markdown)
        if hasattr(self, "export_chrono_check"):
            menu.addSeparator()
            menu.addAction(self.export_chrono_check)
        menu.addSeparator()
        act_cfg = menu.addAction("⚙️ Configure CLI Agents && MCP...")
        act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(self))

