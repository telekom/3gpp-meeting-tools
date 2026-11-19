# --- File: src/modules/emails/ui/tracked_thread_dialog.py ---
from datetime import datetime, timedelta
from pathlib import Path
import re
import subprocess
from typing import Callable, Dict, List, Optional
import urllib.parse

from PyQt5.QtCore import QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtGui import QColor, QDesktopServices
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ai.cli_detector import detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import (
    launch_cli_in_terminal,
    open_mcp_agent_setup_dialog,
    write_cli_context_file,
)
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import install_reader_zoom, zoom_text_widget
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    find_unread_heading_blocks,
    is_cacheable_document_url,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    cache_all_thread_documents_and_attachments,
    cache_and_open_thread_document_url,
    classify_discovered_threads,
    load_meeting_discovery_config,
    mark_thread_read_with_outlook,
    open_outlook_attachment,
    open_outlook_message,
    pin_thread_by_subject_or_cid,
    prepare_thread_llm_workspace,
    reload_outlook_message,
    reply_to_outlook_message,
    save_meeting_discovery_config,
    set_thread_message_flag_with_outlook,
    set_thread_message_read_with_outlook,
)
from modules.emails.core.tracked_threads_workers import (
    OutlookFolderScanWorker,
    SingleThreadRefreshWorker,
)
from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar
from modules.meetings.ui.topic_pills import TagCloudWidget


class TrackedThreadDialog(QDialog):
    """Integrated viewer for a tracked email thread with Markdown export and LLM CLI launch."""

    thread_changed = pyqtSignal()

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        thread_id: str,
        parent: Optional[QWidget] = None,
        ai_suggestions: Optional[List[str]] = None,
        wi_suggestions: Optional[List[str]] = None,
        tdoc_suggestions: Optional[List[str]] = None,
        tdoc_family_expander: Optional[Callable[[List[str]], List[str]]] = None,
        sub_item_suggestions: Optional[List[str]] = None,
        all_existing_tags: Optional[List[str]] = None,
        annotations_save_callback: Optional[Callable[[str, str, str, str], None]] = None,
        ai_to_sub_items: Optional[Dict[str, List[str]]] = None,
    ):
        super().__init__(
            parent,
            Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint,
        )
        self.db = db
        self.thread_id = thread_id
        self._ai_suggestions = ai_suggestions or []
        self._wi_suggestions = wi_suggestions or []
        self._tdoc_suggestions = tdoc_suggestions or []
        self._sub_item_suggestions = sub_item_suggestions or []
        self._ai_to_sub_items: Dict[str, List[str]] = self._normalize_ai_sub_map(
            ai_to_sub_items, self._sub_item_suggestions
        )
        self._all_existing_tags = all_existing_tags or []
        self._annotations_save_callback = annotations_save_callback
        self._tdoc_family_expander = tdoc_family_expander
        self._loading = False
        self._refresh_worker: Optional[SingleThreadRefreshWorker] = None
        self._unread_blocks: List[object] = []
        self._unread_cursor: int = -1
        self._initial_scroll_done: bool = False
        self._settings = QSettings("3GPP_Tools", "TrackedThreadDialog")
        self.setWindowTitle("Tracked Email Thread")
        self.setMinimumSize(520, 420)
        self.resize(1020, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass
        self._build_ui()
        self._load_data()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._initial_scroll_done:
            QTimer.singleShot(50, self._scroll_to_initial_unread)

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
        self.meta_label.setStyleSheet("color: #64748B; font-size: 11px;")
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
            "#LinksPanel { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }"
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
            pill_colors={"bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
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
            pill_colors={"bg": "#F0FDFA", "fg": "#0F766E", "border": "#99F6E4"},
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
            pill_colors={"bg": "#EEF2FF", "fg": "#4338CA", "border": "#C7D2FE"},
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
            pill_colors={"bg": "#ECFDF5", "fg": "#047857", "border": "#A7F3D0"},
        )
        self.tdocs_pills.set_suggestions(self._tdoc_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip().upper())
        self.tdocs_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.tdocs_pills, 1, 3)

        links_grid.addWidget(QLabel("<b>Tags:</b>"), 2, 0, Qt.AlignVCenter)
        self.tags_pills = TagCloudWidget(
            placeholder="Personal tag (e.g. Key Issue, Follow-up)...",
            pill_colors={"bg": "#FEF3C7", "fg": "#92400E", "border": "#FCD34D"},
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
        act_merge = more_menu.addAction("🔗 Merge Split Thread...")
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

        self.company_filter_combo = QComboBox()
        self.company_filter_combo.addItem("All Companies", "")
        self.company_filter_combo.currentIndexChanged.connect(self._render_messages)
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
            "QPushButton { padding: 3px 9px; border-radius: 4px; border: 1px solid #93C5FD; "
            "background-color: #EFF6FF; color: #1D4ED8; font-weight: bold; } "
            "QPushButton:hover { background-color: #DBEAFE; } "
            "QPushButton:disabled { background-color: #F8FAFC; color: #94A3B8; border-color: #E2E8F0; font-weight: normal; }"
        )
        self.next_unread_btn.clicked.connect(self._on_next_unread_clicked)
        filter_row.addWidget(self.next_unread_btn)

        self.msg_count_lbl = QLabel("")
        self.msg_count_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
        filter_row.addWidget(self.msg_count_lbl)
        layout.addLayout(filter_row)

        # Timeline browser (supports Ctrl+Wheel zoom and Ctrl+0 reset)
        self.browser = QTextBrowser()
        self.browser.setOpenLinks(False)
        self.browser.setToolTip("Tip: Hold Ctrl + Mouse Wheel to zoom text size (Ctrl+0 to reset).")
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        install_reader_zoom()
        saved_zoom = self._settings.value("zoom_steps", 0, type=int)
        if saved_zoom:
            zoom_text_widget(self.browser, int(saved_zoom))
        layout.addWidget(self.browser, 1)

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

    def closeEvent(self, event):
        try:
            if getattr(self, "_refresh_worker", None) and self._refresh_worker.isRunning():
                self._refresh_worker.request_cancel()
                self._refresh_worker.wait(1500)
        except Exception:
            pass
        try:
            self._settings.setValue("geometry", self.saveGeometry())
            if hasattr(self, "browser"):
                self._settings.setValue("zoom_steps", int(self.browser.property("_reader_zoom_steps") or 0))
        except Exception:
            pass
        super().closeEvent(event)

    def _populate_company_filter(self, messages: list):
        cur_val = self.company_filter_combo.currentData() or ""
        companies = sorted({
            str(m.get("company") or "").strip()
            for m in messages
            if str(m.get("company") or "").strip() and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
        }, key=str.lower)
        self.company_filter_combo.blockSignals(True)
        self.company_filter_combo.clear()
        self.company_filter_combo.addItem("All Companies", "")
        for c in companies:
            self.company_filter_combo.addItem(c, c)
        idx = self.company_filter_combo.findData(cur_val)
        self.company_filter_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.company_filter_combo.blockSignals(False)

    def _load_data(self):
        thread = self.db.get_thread(self.thread_id)
        if not thread:
            return
        self._loading = True
        try:
            self.title_label.setText(f"📧 {thread.get('title', 'Email Thread')}")
            self.meta_label.setText(
                f"{thread.get('meeting_id') or 'N/A'} | "
                f"{thread.get('message_count', 0)} msgs | "
                f"Unread: {thread.get('unread_count', 0)} | "
                f"Flagged: {thread.get('flagged_count', 0)}"
            )
            self.meta_label.setToolTip(f"Last synced: {thread.get('last_synced_at') or 'Never'}")
            self._update_links_toggle_label(thread)
            self.ais_pills.setText(", ".join(thread.get("linked_ais") or []))
            self.sub_items_pills.setText(", ".join(thread.get("linked_sub_items") or []))
            self.wis_pills.setText(", ".join(thread.get("linked_wis") or []))
            self.tdocs_pills.setText(", ".join(thread.get("linked_tdocs") or []))
            self.tags_pills.setText(", ".join(thread.get("user_tags") or []))
            self._refresh_sub_items_dropdown(thread.get("linked_ais") or [])
            status = str(thread.get("user_status") or "⚪ Neutral").strip()
            has_notes = bool(str(thread.get("user_notes") or "").strip())
            if hasattr(self, "notes_btn"):
                if status and status != "⚪ Neutral":
                    self.notes_btn.setText(f"{status} · 📝 Notes")
                elif has_notes:
                    self.notes_btn.setText("📓 Notes && Status...")
                else:
                    self.notes_btn.setText("📝 Notes && Status...")
            self._populate_company_filter(self.db.get_thread_messages(self.thread_id))
        finally:
            self._loading = False
        self._render_messages()

    def _render_messages(self):
        thread = self.db.get_thread(self.thread_id)
        all_messages = self.db.get_thread_messages(self.thread_id)
        if not thread:
            return

        query = (self.msg_search_edit.text() if hasattr(self, "msg_search_edit") else "").strip().lower()
        comp_filter = (self.company_filter_combo.currentData() if hasattr(self, "company_filter_combo") else "") or ""
        state_filter = (self.state_filter_combo.currentData() if hasattr(self, "state_filter_combo") else "all") or "all"

        filtered = []
        for m in all_messages:
            if state_filter == "unread" and int(m.get("is_read", 0) or 0) != 0:
                continue
            if state_filter == "flagged" and int(m.get("flag_status", 0) or 0) != 2:
                continue
            if comp_filter and (m.get("company") or "").strip().lower() != comp_filter.lower():
                continue
            if query:
                haystack = " ".join([
                    str(m.get("sender_name") or ""),
                    str(m.get("sender_email") or ""),
                    str(m.get("company") or ""),
                    str(m.get("subject") or ""),
                    str(m.get("body_text") or ""),
                ]).lower()
                if query not in haystack:
                    continue
            filtered.append(m)

        if hasattr(self, "msg_count_lbl"):
            if len(filtered) == len(all_messages):
                self.msg_count_lbl.setText(f"{len(all_messages)} message(s)")
            else:
                self.msg_count_lbl.setText(f"Showing {len(filtered)} of {len(all_messages)}")

        newest_first = bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else True
        include_meta = self.show_meta_action.isChecked() if hasattr(self, "show_meta_action") else False
        md = export_thread_to_markdown(
            thread,
            filtered,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=True,
            newest_first=newest_first,
            include_metadata=include_meta,
        )
        self.browser.setMarkdown(md)
        self._unread_blocks = find_unread_heading_blocks(self.browser)
        self._unread_cursor = -1
        self._update_next_unread_button()

    def _update_next_unread_button(self):
        if not hasattr(self, "next_unread_btn"):
            return
        cnt = len(self._unread_blocks)
        self.next_unread_btn.setText(f"🔵 Next Unread ({cnt})")
        self.next_unread_btn.setEnabled(cnt > 0)
        if cnt <= 0:
            self.next_unread_btn.setToolTip("No unread emails in the current view.")
        elif self._unread_cursor >= 0:
            self.next_unread_btn.setToolTip(
                f"Jump to next unread email ({self._unread_cursor + 1}/{cnt})."
            )
        else:
            self.next_unread_btn.setToolTip(f"Jump to the first of {cnt} unread email(s).")

    def _scroll_to_initial_unread(self):
        if self._initial_scroll_done:
            return
        self._initial_scroll_done = True
        blocks = find_unread_heading_blocks(self.browser)
        self._unread_blocks = blocks
        if blocks:
            self._unread_cursor = 0
            scroll_browser_to_block(self.browser, blocks[0])
            self._update_next_unread_button()

    def _on_next_unread_clicked(self):
        blocks = find_unread_heading_blocks(self.browser)
        self._unread_blocks = blocks
        if not blocks:
            self._update_next_unread_button()
            return
        self._unread_cursor = (self._unread_cursor + 1) % len(blocks)
        scroll_browser_to_block(self.browser, blocks[self._unread_cursor])
        self._update_next_unread_button()

    def _get_thread_cache_dir(self) -> Path:
        d = self.db.db_path.parent / "ThreadAttachments" / self.thread_id
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _on_cache_all_docs(self):
        cached = cache_all_thread_documents_and_attachments(self.db, self.thread_id, self._get_thread_cache_dir())
        QMessageBox.information(
            self,
            "Thread Documents Cached",
            f"Cached {len(cached)} document(s)/attachment(s) into:\n{self._get_thread_cache_dir()}",
        )

    def _on_open_thread_folder(self):
        import os
        folder = self._get_thread_cache_dir()
        if hasattr(os, "startfile"):
            os.startfile(str(folder.resolve()))

    def _on_anchor_clicked(self, url: QUrl):
        scheme = url.scheme().lower()
        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if host == "msg" and path_parts:
                if not open_outlook_message(path_parts[0]):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply" and path_parts:
                if not reply_to_outlook_message(path_parts[0], reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
            elif host == "reload" and path_parts:
                msg_id = path_parts[0]
                sb = self.browser.verticalScrollBar()
                prev_scroll = sb.value()
                if reload_outlook_message(self.db, self.thread_id, msg_id):
                    self._load_data()
                    sb.setValue(min(prev_scroll, sb.maximum()))
                    self.thread_changed.emit()
                    notify_success(self, "Email Reloaded", "Message body, tables, and links re-loaded from Outlook.")
                else:
                    notify_toast(self, "Reload Failed", "Could not re-fetch this message from Outlook (it may have been moved or deleted).", variant="warning")
            elif host in ("read", "unread", "flag", "unflag") and path_parts:
                msg_id = path_parts[0]
                sb = self.browser.verticalScrollBar()
                prev_scroll = sb.value()
                if host in ("read", "unread"):
                    set_thread_message_read_with_outlook(self.db, msg_id, is_read=(host == "read"), sync_outlook=True)
                else:
                    set_thread_message_flag_with_outlook(self.db, msg_id, is_flagged=(host == "flag"), sync_outlook=True)
                self._load_data()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.thread_changed.emit()
            elif host == "copy" and path_parts:
                target_id = path_parts[0]
                msgs = self.db.get_thread_messages(self.thread_id)
                msg = next((m for m in msgs if str(m.get("id") or "") == target_id), None)
                if msg:
                    body = (msg.get("clean_body") if self.dedup_check.isChecked() and msg.get("clean_body") else msg.get("body_text")) or ""
                    clip_txt = f"[{msg.get('date_sent', '')}] {msg.get('sender_name', '')} ({msg.get('company', '')}):\n{body.strip()}"
                    QApplication.clipboard().setText(clip_txt)
            elif host == "att" and len(path_parts) >= 2:
                out_dir = self._get_thread_cache_dir()
                if not open_outlook_attachment(path_parts[0], path_parts[1], save_dir=out_dir):
                    QMessageBox.warning(self, "Attachment", f"Could not open attachment '{path_parts[1]}' from Outlook.")
            return
        url_str = url.toString()
        if is_cacheable_document_url(url_str):
            cached_path = cache_and_open_thread_document_url(url_str, self._get_thread_cache_dir(), open_after=True)
            if cached_path:
                return
        QDesktopServices.openUrl(url)

    @staticmethod
    def _normalize_ai_sub_map(
        explicit_map: Optional[Dict[str, List[str]]],
        sub_item_suggestions: List[str],
    ) -> Dict[str, List[str]]:
        out: Dict[str, List[str]] = {}
        for ai_key, subs in (explicit_map or {}).items():
            ak = str(ai_key or "").strip()
            if not ak:
                continue
            bucket = out.setdefault(ak, [])
            for s in (subs or []):
                sc = str(s or "").strip()
                if sc and sc not in bucket:
                    bucket.append(sc)
        if not explicit_map:
            for item in (sub_item_suggestions or []):
                txt = str(item or "").strip()
                if " — AI " in txt:
                    sub_part, ai_part = txt.rsplit(" — AI ", 1)
                    ak = ai_part.strip()
                    sc = sub_part.strip()
                    if ak and sc:
                        bucket = out.setdefault(ak, [])
                        if sc not in bucket:
                            bucket.append(sc)
        return out

    def _get_valid_sub_items_for_ais(self, ais: List[str]) -> List[str]:
        by_lower = {k.lower(): v for k, v in self._ai_to_sub_items.items()}
        valid: List[str] = []
        seen_lower: set = set()
        for ai in (ais or []):
            for sub in by_lower.get(str(ai).strip().lower(), []):
                if sub.lower() not in seen_lower:
                    seen_lower.add(sub.lower())
                    valid.append(sub)
        return valid

    def _refresh_sub_items_dropdown(self, ais: List[str]) -> List[str]:
        valid_subs = self._get_valid_sub_items_for_ais(ais)
        if not hasattr(self, "sub_items_pills"):
            return valid_subs
        if not ais:
            self.sub_items_pills.set_empty_dropdown_label("Assign AI first")
        elif not valid_subs:
            self.sub_items_pills.set_empty_dropdown_label("No sub-topics for AI")
        self.sub_items_pills.set_suggestions(valid_subs, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        return valid_subs

    def _on_links_changed(self):
        if self._loading:
            return
        ais = self.ais_pills.tags()
        valid_subs = self._refresh_sub_items_dropdown(ais)
        valid_subs_lower = {s.lower() for s in valid_subs}
        sub_items = self.sub_items_pills.tags()
        # Automatically prune sub-topics whose parent AI is no longer assigned
        pruned_subs = [s for s in sub_items if s.lower() in valid_subs_lower]
        if pruned_subs != sub_items:
            self._loading = True
            try:
                self.sub_items_pills.setText(", ".join(pruned_subs))
            finally:
                self._loading = False
            sub_items = pruned_subs

        wis = self.wis_pills.tags()
        tdocs = self.tdocs_pills.tags()
        user_tags = self.tags_pills.tags()
        if self._tdoc_family_expander:
            expanded = self._tdoc_family_expander(tdocs)
            if [t.upper() for t in expanded] != [t.upper() for t in tdocs]:
                self._loading = True
                try:
                    self.tdocs_pills.setText(", ".join(expanded))
                finally:
                    self._loading = False
                tdocs = expanded
        self.db.update_thread_links(
            self.thread_id,
            linked_ais=ais,
            linked_wis=wis,
            linked_tdocs=tdocs,
            linked_sub_items=sub_items,
        )
        self.db.update_thread_annotations(self.thread_id, user_tags=user_tags)
        t = self.db.get_thread(self.thread_id) or {}
        self._update_links_toggle_label(t)
        if self._annotations_save_callback:
            self._annotations_save_callback(
                self.thread_id,
                str(t.get("user_status") or "⚪ Neutral"),
                str(t.get("user_notes") or ""),
                ", ".join(user_tags),
            )
        self.thread_changed.emit()
        self._render_messages()

    def _on_edit_notes_and_status(self):
        from modules.meetings.ui.tdocs_dialogs import InteractiveNotesDialog

        thread = self.db.get_thread(self.thread_id) or {}
        row_data = {
            "TDoc": self.thread_id,
            "Title": thread.get("title", ""),
            "Secretary Remarks": f"Tracked Email Thread ({thread.get('message_count', 0)} messages)",
            "My Status": thread.get("user_status") or "⚪ Neutral",
            "My Notes": thread.get("user_notes") or "",
            "My Tags": ", ".join(thread.get("user_tags") or []),
            "_is_tracked_thread": True,
        }

        def _save(tid: str, status: str, notes: str, tags_str: str = ""):
            tag_list = [x.strip() for x in str(tags_str or "").replace(";", ",").split(",") if x.strip()]
            self.db.update_thread_annotations(tid, user_status=status, user_notes=notes, user_tags=tag_list)
            if self._annotations_save_callback:
                self._annotations_save_callback(tid, status, notes, tags_str)
            self._load_data()
            self.thread_changed.emit()

        all_tags = sorted(set(self._all_existing_tags) | set(self.tags_pills.tags()), key=str.lower)
        InteractiveNotesDialog(self, self.thread_id, row_data, _save, all_existing_tags=all_tags).exec_()

    def _on_refresh(self):
        self._start_thread_worker(reload_cached=False)

    def _on_reload_all_cached(self):
        self._start_thread_worker(reload_cached=True)

    def _on_reload_from_msg_file(self):
        from modules.emails.core.tracked_threads_service import parse_msg_file

        paths, _ = QFileDialog.getOpenFileNames(self, "Select Outlook .msg File(s) to Reload/Merge", "", "Outlook Message (*.msg)")
        if not paths:
            return
        loaded = []
        for p in paths:
            parsed = parse_msg_file(Path(p))
            if parsed:
                loaded.append(parsed)
        if loaded:
            self.db.upsert_messages(self.thread_id, loaded)
            self._load_data()
            self.thread_changed.emit()
            notify_success(self, "Loaded .msg File(s)", f"Updated {len(loaded)} message(s) in this thread.")
        else:
            notify_toast(self, "Load Failed", "Could not parse selected .msg file(s).", variant="warning")

    def _start_thread_worker(self, reload_cached: bool = False):
        if self._refresh_worker and self._refresh_worker.isRunning():
            return
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setText("⏳ Reloading..." if reload_cached else "⏳ Syncing...")
        if reload_cached:
            msgs = [m for m in self.db.get_thread_messages(self.thread_id) if m.get("id") and not str(m.get("id")).startswith("MSG-")]
            self.sync_bar.start_busy("Reloading cached Outlook emails...", total=len(msgs))
        else:
            thread = self.db.get_thread(self.thread_id) or {}
            folders = thread.get("outlook_folders") or []
            self.sync_bar.start_busy("Scanning Outlook folders...", total=len(folders), remaining=folders)

        worker = SingleThreadRefreshWorker(self.db.db_path, self.thread_id, reload_cached=reload_cached, parent=self)
        worker.progress.connect(self.sync_bar.update_progress)
        worker.finished_sync.connect(lambda ok, cnt, err: self._on_refresh_finished(ok, cnt, err, reload_cached=reload_cached))
        self._refresh_worker = worker
        worker.start()

    def _on_cancel_refresh(self):
        if self._refresh_worker and self._refresh_worker.isRunning():
            self._refresh_worker.request_cancel()

    def _on_refresh_finished(self, ok: bool, count: int, err: str, reload_cached: bool = False):
        self.refresh_btn.setEnabled(True)
        self.refresh_btn.setText("🔄 Refresh")
        self._load_data()
        self.thread_changed.emit()
        if ok:
            self.sync_bar.finish_ok()
            if reload_cached:
                notify_success(self, "Emails Reloaded", f"Re-parsed {count} cached Outlook message(s).")
            elif count > 0:
                notify_success(self, "Thread Refreshed", f"{count} new message(s) synced.")
            else:
                notify_info(self, "Thread Up-to-Date", "No new replies found in Outlook.")
        else:
            self.sync_bar.show_error(err or "Thread sync did not finish cleanly.")
            notify_toast(self, "Thread Sync", err or "Sync warning", variant="warning")

    def _on_mark_read(self):
        mark_thread_read_with_outlook(self.db, self.thread_id, sync_outlook=True)
        self._load_data()
        self.thread_changed.emit()

    def _on_merge_thread(self):
        current = self.db.get_thread(self.thread_id)
        if not current:
            return
        candidates = [
            t for t in self.db.get_threads_for_meeting(current.get("meeting_id", ""))
            if t["thread_id"] != self.thread_id
        ]
        if not candidates:
            QMessageBox.information(self, "No Other Threads", "No other tracked threads in this meeting to merge.")
            return
        items = [f"{t['thread_id']} — {t['title']} ({t.get('message_count', 0)} msgs)" for t in candidates]
        choice, ok = QInputDialog.getItem(
            self,
            "Merge Split Thread",
            "Select another tracked thread to merge into this thread:",
            items,
            0,
            False,
        )
        if ok and choice:
            src_id = choice.split(" — ", 1)[0].strip()
            self.db.merge_threads(self.thread_id, [src_id])
            self._load_data()
            self.thread_changed.emit()

    def _on_delete_thread(self):
        if QMessageBox.question(
            self,
            "Unpin Email Thread",
            "Remove this tracked email thread from the meeting?",
            QMessageBox.Yes | QMessageBox.No,
        ) == QMessageBox.Yes:
            self.db.delete_thread(self.thread_id)
            self.thread_changed.emit()
            self.accept()

    def _get_export_newest_first(self) -> bool:
        if hasattr(self, "export_chrono_check") and self.export_chrono_check.isChecked():
            return False
        return bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else False

    def _get_markdown(self) -> str:
        thread = self.db.get_thread(self.thread_id) or {}
        messages = self.db.get_thread_messages(self.thread_id)
        return export_thread_to_markdown(
            thread,
            messages,
            deduplicate_quotes=self.dedup_check.isChecked(),
            newest_first=self._get_export_newest_first(),
        )

    def _on_copy_markdown(self):
        QApplication.clipboard().setText(self._get_markdown())
        QMessageBox.information(self, "Copied", "Thread Markdown copied to clipboard.")

    def _on_save_markdown(self):
        thread = self.db.get_thread(self.thread_id) or {}
        safe_name = "".join(c for c in thread.get("title", "thread") if c.isalnum() or c in " -_").strip()[:60]
        path, _ = QFileDialog.getSaveFileName(self, "Save Thread Markdown", f"{safe_name or 'thread'}.md", "Markdown (*.md)")
        if path:
            Path(path).write_text(self._get_markdown(), encoding="utf-8")

    def _on_launch_cli(self, agent_id: Optional[str] = None):
        cli = (agent_id or "codex").strip().lower()
        prep = prepare_thread_llm_workspace(
            self.db,
            self.thread_id,
            self._get_thread_cache_dir(),
            deduplicate_quotes=self.dedup_check.isChecked(),
            newest_first=self._get_export_newest_first(),
            cache_linked_docs=True,
        )
        md_path = prep["thread_md"]
        ws_dir = prep["workspace_dir"]
        doc_mds = prep["doc_mds"]
        extra_docs = ""
        if doc_mds:
            names = ", ".join(p.name for p in doc_mds[:10])
            extra_docs = f" Companion converted draft documents are also in {ws_dir} ({names})."
        prompt = (
            f"Analyze the 3GPP email thread in {md_path} and summarize the views exchanged by company, "
            f"key disagreements, and open points.{extra_docs}"
        )
        thread = self.db.get_thread(self.thread_id) or {}
        write_cli_context_file(
            ws_dir,
            task_mode="custom",
            custom_question=prompt,
            markdown_files=[md_path] + list(doc_mds),
            briefing_title=f"Tracked Email Thread: {thread.get('title', self.thread_id)}",
            user_notes_override=str(thread.get("user_notes") or ""),
        )
        launch_cli_in_terminal(
            cli,
            ws_dir,
            bootstrap_prompt=prompt,
            parent=self,
            auto_execute=False,
        )


def _infer_meeting_tag(meeting_id: str, wg: str) -> str:
    raw = str(meeting_id or "").strip()
    if not raw:
        return str(wg or "").strip()
    if "#" in raw:
        return raw
    m = re.match(r"^([A-Za-z0-9]+)\s+(\d+.*)$", raw)
    if m:
        return f"{m.group(1)}#{m.group(2)}"
    if wg and raw.isdigit():
        return f"{wg}#{raw}"
    return raw


class PinFromOutlookDialog(QDialog):
    """Browse, classify, ignore, and pin untracked email threads from configured WG Outlook folders."""

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        meeting_id: str,
        wg: str,
        parent: Optional[QWidget] = None,
        agenda_map: Optional[Dict[str, object]] = None,
        valid_wis: Optional[List[str]] = None,
        meeting_dir: Optional[Path] = None,
        general_email_db=None,
        valid_tdocs: Optional[set] = None,
        mtg_start_date: str = "",
        mtg_end_date: str = "",
        auto_scan: bool = True,
    ):
        super().__init__(parent)
        self.db = db
        self.meeting_id = meeting_id
        self.wg = wg
        self._agenda_map = agenda_map
        self._valid_wis = valid_wis
        self._meeting_dir = Path(meeting_dir) if meeting_dir else None
        self._general_email_db = general_email_db
        self._valid_tdocs = set(valid_tdocs) if valid_tdocs else set()
        self.pinned_thread_ids: List[str] = []
        self._raw_threads: List[dict] = []
        self._threads: List[dict] = []
        self._scan_worker: Optional[OutlookFolderScanWorker] = None
        self._loading_config = False

        default_tag = _infer_meeting_tag(meeting_id, wg)
        self._config = load_meeting_discovery_config(
            self._meeting_dir,
            default_tag=default_tag,
            default_start=str(mtg_start_date or "").strip()[:10],
            default_end=str(mtg_end_date or "").strip()[:10],
        )

        self.setWindowTitle(f"Discover & Pin Email Threads — Configured {wg} Outlook Folders")
        self.resize(1120, 680)
        self._build_ui()
        if auto_scan:
            self._load_threads()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # Row 1: Search & Coverage Filter & Rescan
        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        top_row.addWidget(QLabel("<b>Coverage:</b>"))
        self.coverage_combo = QComboBox()
        self.coverage_combo.addItem("🔍 Untracked & Body-Only (🔴 + 🟡)", "UNTRACKED_AND_BODY")
        self.coverage_combo.addItem("🔴 Strictly Untracked Only (No TDoc)", "STRICT_UNTRACKED")
        self.coverage_combo.addItem("🟡 All Unpinned (Includes Subject TDoc)", "ALL_UNPINNED")
        self.coverage_combo.addItem("📋 All Discovered (Includes Pinned)", "ALL")
        self.coverage_combo.addItem("🚫 Ignored Threads", "IGNORED")
        cov_val = str(self._config.get("coverage_filter") or "UNTRACKED_AND_BODY").upper()
        cov_idx = max(0, self.coverage_combo.findData(cov_val))
        self.coverage_combo.setCurrentIndex(cov_idx)
        self.coverage_combo.currentIndexChanged.connect(self._on_config_ui_changed)
        top_row.addWidget(self.coverage_combo)

        self.search_edit = QLineEdit()
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setPlaceholderText("🔍 Search subject, sender, TDoc, AI, or WI...")
        self.search_edit.textChanged.connect(self._apply_filter)
        top_row.addWidget(self.search_edit, 1)

        self.reload_btn = QPushButton("🔄 Rescan Folders")
        self.reload_btn.clicked.connect(self._load_threads)
        top_row.addWidget(self.reload_btn)
        layout.addLayout(top_row)

        # Row 2: Configurable Subject Tag & Meeting Date Window (persisted per meeting)
        cfg_frame = QFrame(self)
        cfg_frame.setObjectName("DiscoveryCfgBar")
        cfg_frame.setStyleSheet(
            "#DiscoveryCfgBar { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }"
        )
        cfg_row = QHBoxLayout(cfg_frame)
        cfg_row.setContentsMargins(8, 5, 8, 5)
        cfg_row.setSpacing(8)

        self.require_tag_check = QCheckBox("Require Subject Tag:")
        self.require_tag_check.setChecked(bool(self._config.get("require_subject_tag", False)))
        self.require_tag_check.setToolTip(
            "Only show threads whose Outlook subject contains this substring (configurable & saved per meeting)."
        )
        self.require_tag_check.toggled.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.require_tag_check)

        self.tag_edit = QLineEdit(str(self._config.get("subject_tag") or ""))
        self.tag_edit.setPlaceholderText("e.g. SA2#178")
        self.tag_edit.setMaximumWidth(150)
        self.tag_edit.setClearButtonEnabled(True)
        self.tag_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.tag_edit)

        self.use_date_check = QCheckBox("Date Window:")
        self.use_date_check.setChecked(bool(self._config.get("use_date_filter", False)))
        self.use_date_check.setToolTip("Restrict results to the meeting date window ± buffer days.")
        self.use_date_check.toggled.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.use_date_check)

        self.start_date_edit = QLineEdit(str(self._config.get("start_date") or ""))
        self.start_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.start_date_edit.setMaximumWidth(95)
        self.start_date_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.start_date_edit)

        cfg_row.addWidget(QLabel("to"))

        self.end_date_edit = QLineEdit(str(self._config.get("end_date") or ""))
        self.end_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.end_date_edit.setMaximumWidth(95)
        self.end_date_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.end_date_edit)

        cfg_row.addWidget(QLabel("−"))
        self.days_before_spin = QSpinBox()
        self.days_before_spin.setRange(0, 90)
        self.days_before_spin.setValue(int(self._config.get("days_before", 14) or 0))
        self.days_before_spin.setSuffix("d")
        self.days_before_spin.setToolTip("Days before start date to include.")
        self.days_before_spin.valueChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.days_before_spin)

        cfg_row.addWidget(QLabel("+"))
        self.days_after_spin = QSpinBox()
        self.days_after_spin.setRange(0, 90)
        self.days_after_spin.setValue(int(self._config.get("days_after", 7) or 0))
        self.days_after_spin.setSuffix("d")
        self.days_after_spin.setToolTip("Days after end date to include.")
        self.days_after_spin.valueChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.days_after_spin)

        cfg_row.addStretch()
        layout.addWidget(cfg_frame)

        self.scan_bar = AsyncSyncProgressBar(self)
        self.scan_bar.cancel_requested.connect(self._on_cancel_scan)
        layout.addWidget(self.scan_bar)

        self.summary_label = QLabel("")
        self.summary_label.setStyleSheet("color: #475569; font-size: 11px;")
        layout.addWidget(self.summary_label)

        # Multi-column discovery table
        self.table = QTableWidget(0, 6, self)
        self.table.setHorizontalHeaderLabels([
            "Coverage",
            "Msgs (Unread)",
            "Subject / Thread Title",
            "Suggested AI / WI / TDocs",
            "Last Sender",
            "Last Reply",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.itemSelectionChanged.connect(self._update_action_buttons)
        layout.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        self.ignore_btn = QPushButton("🚫 Ignore Selected")
        self.ignore_btn.setToolTip("Hide selected thread(s) from future untracked discovery for this meeting.")
        self.ignore_btn.clicked.connect(self._on_toggle_ignore_selected)
        btn_row.addWidget(self.ignore_btn)

        btn_row.addStretch()
        self.pin_btn = QPushButton("📌 Pin Selected Thread(s)")
        self.pin_btn.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; font-weight: bold; "
            "padding: 6px 14px; border-radius: 5px; border: none; } "
            "QPushButton:hover { background-color: #1D4ED8; } "
            "QPushButton:disabled { background-color: #CBD5E1; color: #64748B; }"
        )
        self.pin_btn.clicked.connect(self._on_pin_selected)
        btn_row.addWidget(self.pin_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)
        self._update_action_buttons()

    def _on_config_ui_changed(self, *_args):
        if self._loading_config:
            return
        self._config["coverage_filter"] = self.coverage_combo.currentData() or "UNTRACKED_AND_BODY"
        self._config["require_subject_tag"] = bool(self.require_tag_check.isChecked())
        self._config["subject_tag"] = self.tag_edit.text().strip()
        self._config["use_date_filter"] = bool(self.use_date_check.isChecked())
        self._config["start_date"] = self.start_date_edit.text().strip()
        self._config["end_date"] = self.end_date_edit.text().strip()
        self._config["days_before"] = int(self.days_before_spin.value())
        self._config["days_after"] = int(self.days_after_spin.value())
        save_meeting_discovery_config(self._meeting_dir, self._config)
        self._update_action_buttons()
        self._apply_filter()

    def _on_cancel_scan(self):
        if self._scan_worker and self._scan_worker.isRunning():
            self._scan_worker.request_cancel()

    def _stop_scan_worker(self):
        if getattr(self, "_scan_worker", None) and self._scan_worker.isRunning():
            self._scan_worker.request_cancel()
            self._scan_worker.wait(1500)

    def reject(self):
        self._stop_scan_worker()
        super().reject()

    def closeEvent(self, event):
        self._stop_scan_worker()
        super().closeEvent(event)

    def _load_threads(self):
        if self._scan_worker and self._scan_worker.isRunning():
            return
        self.reload_btn.setEnabled(False)
        self.reload_btn.setText("⏳ Scanning...")
        self.scan_bar.start_busy(f"Scanning {self.wg} Outlook folders & Sent Items...")

        worker = OutlookFolderScanWorker(self.wg, parent=self)
        worker.progress.connect(self.scan_bar.update_progress)
        worker.finished_scan.connect(self._on_scan_finished)
        self._scan_worker = worker
        worker.start()

    def _reclassify_threads(self):
        ignored_set = set(self._config.get("ignored_thread_keys") or [])
        self._threads = classify_discovered_threads(
            self._raw_threads,
            self.db,
            self._general_email_db,
            meeting_id=self.meeting_id,
            valid_tdocs=self._valid_tdocs,
            agenda_map=self._agenda_map,
            valid_wis=self._valid_wis,
            ignored_keys=ignored_set,
        )

    def _on_scan_finished(self, ok: bool, threads: list, err: str):
        self.reload_btn.setEnabled(True)
        self.reload_btn.setText("🔄 Rescan Folders")
        self._raw_threads = list(threads or [])
        self._reclassify_threads()
        self._apply_filter()
        if ok:
            self.scan_bar.finish_ok()
        else:
            self.scan_bar.show_error(err or "Folder scan did not finish cleanly.")

    @staticmethod
    def _parse_iso_date(val: str) -> Optional[datetime]:
        txt = str(val or "").strip()[:10]
        if not txt:
            return None
        try:
            return datetime.strptime(txt, "%Y-%m-%d")
        except Exception:
            return None

    def _matches_tag_and_date(self, t: dict) -> bool:
        if self.require_tag_check.isChecked():
            tag = self.tag_edit.text().strip().lower()
            if tag:
                subj_blob = f"{t.get('raw_subject', '')} {t.get('title', '')}".lower()
                if tag not in subj_blob:
                    return False

        if self.use_date_check.isChecked():
            last_dt = self._parse_iso_date(t.get("last_date", ""))
            start_dt = self._parse_iso_date(self.start_date_edit.text())
            end_dt = self._parse_iso_date(self.end_date_edit.text())
            if last_dt is not None:
                if start_dt is not None:
                    min_dt = start_dt - timedelta(days=int(self.days_before_spin.value()))
                    if last_dt < min_dt:
                        return False
                if end_dt is not None:
                    max_dt = end_dt + timedelta(days=int(self.days_after_spin.value()))
                    if last_dt > max_dt:
                        return False
        return True

    def _apply_filter(self):
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        q = self.search_edit.text().strip().lower()

        scoped: List[dict] = []
        counts = {"STRICT_UNTRACKED": 0, "BODY_TDOC": 0, "SUBJECT_TDOC": 0, "PINNED": 0, "IGNORED": 0}
        for t in self._threads:
            if not self._matches_tag_and_date(t):
                continue
            is_ignored = bool(t.get("is_ignored"))
            cov = str(t.get("coverage") or "STRICT_UNTRACKED")
            if is_ignored and cov != "PINNED":
                counts["IGNORED"] += 1
            elif cov in counts:
                counts[cov] += 1

            if cov_mode == "IGNORED":
                if not is_ignored:
                    continue
            else:
                if is_ignored and cov_mode != "ALL":
                    continue
                if cov_mode == "UNTRACKED_AND_BODY" and cov not in ("STRICT_UNTRACKED", "BODY_TDOC"):
                    continue
                if cov_mode == "STRICT_UNTRACKED" and cov != "STRICT_UNTRACKED":
                    continue
                if cov_mode == "ALL_UNPINNED" and cov == "PINNED":
                    continue

            if q:
                hay = " ".join([
                    str(t.get("title") or ""),
                    str(t.get("last_sender") or ""),
                    " ".join(t.get("suggested_tdocs") or []),
                    " ".join(t.get("suggested_ais") or []),
                    " ".join(t.get("suggested_wis") or []),
                ]).lower()
                if q not in hay:
                    continue
            scoped.append(t)

        self.table.setRowCount(len(scoped))
        for r, t in enumerate(scoped):
            cov = str(t.get("coverage") or "")
            cov_item = QTableWidgetItem(str(t.get("coverage_label") or cov))
            cov_item.setData(Qt.UserRole, t)
            if cov == "STRICT_UNTRACKED" and not t.get("is_ignored"):
                cov_item.setForeground(QColor("#B91C1C"))
            elif cov == "BODY_TDOC" and not t.get("is_ignored"):
                cov_item.setForeground(QColor("#B45309"))
            elif cov == "SUBJECT_TDOC":
                cov_item.setForeground(QColor("#15803D"))
            elif cov == "PINNED":
                cov_item.setForeground(QColor("#1D4ED8"))
            else:
                cov_item.setForeground(QColor("#64748B"))

            tot_cnt = int(t.get("count", 1) or 0)
            unread_cnt = int(t.get("unread_count", 0) or 0)
            msg_str = f"{tot_cnt} ({unread_cnt} new)" if unread_cnt > 0 else f"{tot_cnt}"
            msg_item = QTableWidgetItem(msg_str)
            msg_item.setTextAlignment(Qt.AlignCenter)
            if unread_cnt > 0:
                font = msg_item.font()
                font.setBold(True)
                msg_item.setFont(font)
                msg_item.setForeground(QColor("#1D4ED8"))

            title_item = QTableWidgetItem(str(t.get("title") or ""))
            if t.get("raw_subject") and t.get("raw_subject") != t.get("title"):
                title_item.setToolTip(str(t.get("raw_subject")))

            sugg_parts = []
            if t.get("suggested_ais"):
                sugg_parts.append("AI: " + ", ".join(t["suggested_ais"][:3]))
            if t.get("suggested_wis"):
                sugg_parts.append("WI: " + ", ".join(t["suggested_wis"][:2]))
            if t.get("suggested_tdocs"):
                sugg_parts.append("TDoc: " + ", ".join(t["suggested_tdocs"][:3]))
            sugg_item = QTableWidgetItem(" · ".join(sugg_parts) if sugg_parts else "—")

            sender_item = QTableWidgetItem(str(t.get("last_sender") or ""))
            date_item = QTableWidgetItem(str(t.get("last_date") or ""))

            self.table.setItem(r, 0, cov_item)
            self.table.setItem(r, 1, msg_item)
            self.table.setItem(r, 2, title_item)
            self.table.setItem(r, 3, sugg_item)
            self.table.setItem(r, 4, sender_item)
            self.table.setItem(r, 5, date_item)

        self.summary_label.setText(
            f"Showing {len(scoped)} thread(s)  •  "
            f"🔴 {counts['STRICT_UNTRACKED']} no TDoc  |  "
            f"🟡 {counts['BODY_TDOC']} body-only  |  "
            f"🟢 {counts['SUBJECT_TDOC']} subject TDoc  |  "
            f"📌 {counts['PINNED']} pinned  |  "
            f"🚫 {counts['IGNORED']} ignored"
        )
        self._update_action_buttons()

    def _selected_threads(self) -> List[dict]:
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}) if self.table.selectionModel() else []
        out = []
        for r in rows:
            item = self.table.item(r, 0)
            data = item.data(Qt.UserRole) if item else None
            if isinstance(data, dict):
                out.append(data)
        return out

    def _update_action_buttons(self):
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        selected = self._selected_threads()
        if cov_mode == "IGNORED":
            self.ignore_btn.setText("↩️ Un-ignore Selected")
            self.ignore_btn.setToolTip("Restore selected thread(s) back into untracked discovery.")
        else:
            self.ignore_btn.setText("🚫 Ignore Selected")
            self.ignore_btn.setToolTip("Hide selected thread(s) from future untracked discovery for this meeting.")
        self.ignore_btn.setEnabled(bool(selected))
        self.pin_btn.setEnabled(bool(selected))

    def _on_toggle_ignore_selected(self):
        selected = self._selected_threads()
        if not selected:
            return
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        ignored = set(self._config.get("ignored_thread_keys") or [])
        for t in selected:
            nkey = str(t.get("norm_key") or "").strip().lower()
            if not nkey:
                continue
            if cov_mode == "IGNORED":
                ignored.discard(nkey)
            else:
                ignored.add(nkey)
        self._config["ignored_thread_keys"] = sorted(ignored)
        save_meeting_discovery_config(self._meeting_dir, self._config)
        self._reclassify_threads()
        self._apply_filter()

    def _on_pin_selected(self):
        selected = self._selected_threads()
        if not selected:
            QMessageBox.information(self, "No Thread Selected", "Please select at least one email thread to pin.")
            return
        ignored = set(self._config.get("ignored_thread_keys") or [])
        changed_ignored = False
        for t in selected:
            nkey = str(t.get("norm_key") or "").strip().lower()
            if nkey and nkey in ignored:
                ignored.discard(nkey)
                changed_ignored = True
            tid = pin_thread_by_subject_or_cid(
                self.db,
                title=t["title"],
                normalized_subject=t["normalized_subject"],
                conversation_id=t.get("conversation_id", ""),
                meeting_id=self.meeting_id,
                wg=self.wg,
                agenda_map=self._agenda_map,
                valid_wis=self._valid_wis,
            )
            self.pinned_thread_ids.append(tid)
        if changed_ignored:
            self._config["ignored_thread_keys"] = sorted(ignored)
            save_meeting_discovery_config(self._meeting_dir, self._config)
        self.accept()

