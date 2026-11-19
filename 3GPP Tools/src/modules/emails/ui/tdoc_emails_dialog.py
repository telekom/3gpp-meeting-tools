# --- File: src/modules/emails/ui/tdoc_emails_dialog.py ---
import html
import logging
import os
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
import urllib.parse
import webbrowser

from PyQt5.QtCore import QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
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
from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill, scale_px
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.general_email_db import (
    GeneralEmailDatabase,
    reload_general_email_from_outlook,
    reload_tdoc_emails_from_outlook,
)
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.thread_markdown_exporter import (
    apply_email_reader_document_style,
    export_thread_to_markdown,
    extract_thread_document_urls,
    find_unread_heading_blocks,
    is_cacheable_document_url,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_service import (
    cache_and_open_thread_document_url,
    normalize_subject,
    open_outlook_attachment,
    open_outlook_message,
    reply_to_outlook_message,
)
from modules.emails.ui.email_summary_dialog import EmailSummaryDialog

logger = logging.getLogger(__name__)


class StandaloneEmailReaderWindow(QWidget):
    """Independent, fully resizable top-level window for reading emails on laptops/secondary monitors."""

    tdoc_selected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        self.setWindowFlags(Qt.Window)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.setWindowTitle("📖 Email Message Viewer")
        self.resize(850, 650)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        top_bar = QHBoxLayout()
        self.lbl_title = QLabel("<b>Message Details</b>")
        self.lbl_title.setStyleSheet("font-size: 13px; color: #005A9E;")
        top_bar.addWidget(self.lbl_title)
        top_bar.addStretch()

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.close)
        top_bar.addWidget(btn_close)
        layout.addLayout(top_bar)

        self.browser = QTextBrowser()
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setStyleSheet("background: white; border: 1px solid #CCC; border-radius: 4px; padding: 12px;")
        self._reader_settings = QSettings("3GPP_Tools", "StandaloneEmailReaderWindow")
        bind_reader_zoom_settings(self.browser, self._reader_settings, key="zoom_steps")
        layout.addWidget(self.browser)

    def update_content(self, html_content: str, subject: str = ""):
        if subject:
            self.setWindowTitle(f"📖 {subject}")
            self.lbl_title.setText(f"<b>{html.escape(subject[:80])}</b>")
        self.browser.setHtml(html_content)

    def _on_anchor_clicked(self, url: QUrl):
        url_str = url.toString()
        if url_str.startswith("tdoc:"):
            target_tdoc = url_str.split(":", 1)[1].strip()
            self.tdoc_selected.emit(target_tdoc)
        elif url_str.startswith("mailto:") or url_str.startswith("http://") or url_str.startswith("https://"):
            webbrowser.open(url_str)


class TDocEmailsDialog(QDialog):
    """Unified continuous Markdown timeline viewer for emails linked to a 3GPP TDoc family."""

    data_changed = pyqtSignal()
    tdoc_selected = pyqtSignal(str)

    def __init__(
        self,
        target_tdoc: str,
        family_tdocs: list,
        db_path: Path,
        wg: str = "SA2",
        save_callback=None,
        parent=None,
        pin_thread_callback: Optional[Callable[[dict], None]] = None,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.target_tdoc = str(target_tdoc or "").strip().upper()
        self.family_tdocs = [str(t).strip().upper() for t in (family_tdocs or []) if str(t).strip()]
        if self.target_tdoc and self.target_tdoc not in self.family_tdocs:
            self.family_tdocs.insert(0, self.target_tdoc)

        self.wg = str(wg or "SA2").strip().upper()
        self.save_callback = save_callback
        self._pin_thread_callback = pin_thread_callback
        self.db = GeneralEmailDatabase(db_path)
        self._summary_dialog = None
        self.emails: List[dict] = []
        self._filtered_emails: List[dict] = []
        self._unread_blocks: List[Any] = []
        self._unread_cursor: int = -1
        self._initial_scroll_done: bool = False
        self._loading = False
        self._settings = QSettings("3GPP_Tools", "TDocEmailsDialog")

        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setWindowTitle(f"📧 Related Emails: {self.target_tdoc}")
        self.setMinimumSize(620, 460)
        self.resize(1100, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        self._init_ui()
        self._load_emails()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._initial_scroll_done:
            QTimer.singleShot(50, self._scroll_to_initial_unread)

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Row 1: Compact Family Breadcrumb & TDoc Match Scope Controls
        top_card = QFrame(self)
        top_card.setObjectName("TDocEmailHeader")
        top_card.setStyleSheet(
            "#TDocEmailHeader { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }"
        )
        top_layout = QHBoxLayout(top_card)
        top_layout.setContentsMargins(10, 6, 10, 6)
        top_layout.setSpacing(10)

        chips_text = " ➔ ".join(
            f"<b>{html.escape(t)}</b>" if t == self.target_tdoc else html.escape(t)
            for t in self.family_tdocs
        )
        lbl_chips = QLabel(f"<b>Document Family:</b> {chips_text}")
        lbl_chips.setWordWrap(True)
        top_layout.addWidget(lbl_chips, 1)

        self.chk_family = QCheckBox("Show Family Revisions")
        self.chk_family.setChecked(True)
        self.chk_family.setToolTip("Include emails mentioning any predecessor or revision in this TDoc family.")
        self.chk_family.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_family)

        self.chk_include_quoted = QCheckBox("Include Quoted Matches")
        self.chk_include_quoted.setToolTip("Uncheck to hide emails where the TDoc was only cited inside quoted reply history.")
        self.chk_include_quoted.setChecked(True)
        self.chk_include_quoted.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_include_quoted)

        self.chk_show_ignored = QCheckBox("Show Ignored")
        self.chk_show_ignored.setToolTip("Show emails previously marked as Ignored for this TDoc.")
        self.chk_show_ignored.setChecked(False)
        self.chk_show_ignored.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_show_ignored)

        self.lbl_counts = QLabel("0 Total (0 Unread)")
        self.lbl_counts.setStyleSheet("font-weight: bold; color: #1D4ED8;")
        top_layout.addWidget(self.lbl_counts)
        layout.addWidget(top_card)

        # Row 2: Lean Action Toolbar (mirrors TrackedThreadDialog)
        toolbar_widget = QWidget(self)
        toolbar_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        toolbar_sp.setHeightForWidth(True)
        toolbar_widget.setSizePolicy(toolbar_sp)
        toolbar = FlowLayout(toolbar_widget, margin=0, h_spacing=6, v_spacing=4)

        self.btn_reload_outlook = QPushButton("🔄 Reload from Outlook")
        self.btn_reload_outlook.setToolTip(
            "Re-fetch these cached emails from Microsoft Outlook to populate HTML tables, attachments, and read/flag state."
        )
        self.btn_reload_outlook.clicked.connect(self._on_reload_all_from_outlook)
        toolbar.addWidget(self.btn_reload_outlook)

        self.btn_mark_all_read = QPushButton("✔️ Mark All Read")
        self.btn_mark_all_read.setToolTip("Mark all emails for this TDoc family as read in local DB and Outlook.")
        self.btn_mark_all_read.clicked.connect(self._mark_all_family_read)
        toolbar.addWidget(self.btn_mark_all_read)

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

        self.btn_pin_thread = QPushButton("📌 Pin as Tracked Thread")
        self.btn_pin_thread.setToolTip("Promote this email discussion into a persistent Tracked Email Thread in the TDocs table.")
        self.btn_pin_thread.clicked.connect(self._on_pin_as_tracked_thread)
        self.btn_pin_thread.setVisible(self._pin_thread_callback is not None)
        toolbar.addWidget(self.btn_pin_thread)

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

        # Row 3: Search, Company, Read/Flag State, Display Order & Next Unread
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
        self.order_combo.addItem("⬇️ Newest First", True)
        self.order_combo.addItem("⬆️ Oldest First", False)
        saved_newest = self._settings.value("newest_first", True, type=bool)
        self.order_combo.setCurrentIndex(0 if saved_newest else 1)
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        filter_row.addWidget(self.order_combo)

        self.next_unread_btn = QPushButton("🔵 Next Unread (0)")
        self.next_unread_btn.setToolTip("Jump to the next unread email in the discussion timeline.")
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
        layout.addLayout(filter_row)

        # Row 4: Continuous Timeline Browser
        self.browser = QTextBrowser()
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.browser, self._settings, key="zoom_steps")
        layout.addWidget(self.browser, 1)

    @staticmethod
    def _compress_blank_lines(text: str, max_consecutive_empty: int = 1) -> str:
        if not text:
            return ""
        normalized = text.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
        lines = normalized.split("\n")
        cleaned_lines = []
        empty_count = 0
        for line in lines:
            if not line.strip():
                empty_count += 1
                if empty_count <= max_consecutive_empty:
                    cleaned_lines.append("")
            else:
                empty_count = 0
                cleaned_lines.append(line.rstrip())
        return "\n".join(cleaned_lines).strip()

    def _on_order_changed(self):
        self._settings.setValue("newest_first", bool(self.order_combo.currentData()))
        self._render_messages()

    def _on_show_metadata_toggled(self, checked: bool):
        self._settings.setValue("show_metadata_header", bool(checked))
        self._render_messages()

    def _build_synthetic_thread_dict(self) -> dict:
        scope = self.family_tdocs if self.chk_family.isChecked() else [self.target_tdoc]
        base_title = normalize_subject(self.emails[0].get("subject", "")) if self.emails else ""
        return {
            "thread_id": f"TDOC-{self.target_tdoc}",
            "title": f"Emails Referencing {self.target_tdoc}" + (f": {base_title}" if base_title else ""),
            "normalized_subject": base_title.lower(),
            "wg": self.wg,
            "linked_tdocs": list(scope),
        }

    def _populate_company_filter(self, messages: List[dict]):
        companies = sorted(
            {
                str(m.get("company") or "").strip()
                for m in messages
                if str(m.get("company") or "").strip()
                and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
            },
            key=str.lower,
        )
        self.company_filter_combo.blockSignals(True)
        self.company_filter_combo.updateItems(companies)
        self.company_filter_combo.blockSignals(False)

    def _load_emails(self):
        self._loading = True
        try:
            query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
            raw_emails = self.db.get_emails_for_tdocs(query_set, show_ignored=self.chk_show_ignored.isChecked())
            if not self.chk_include_quoted.isChecked():
                self.emails = [e for e in raw_emails if e.get("match_location") != "Quoted"]
            else:
                self.emails = raw_emails

            # Normalize chronological ascending order internally; display order respects order_combo
            self.emails.sort(key=lambda m: str(m.get("date_received") or m.get("date_sent") or ""))
            self._populate_company_filter(self.emails)
            self._update_unread_count_label()
        finally:
            self._loading = False
        self._render_messages()

    def _update_unread_count_label(self):
        unread = sum(1 for e in self.emails if int(e.get("is_read", 0) or 0) == 0 and int(e.get("is_ignored", 0) or 0) == 0)
        flagged = sum(1 for e in self.emails if int(e.get("flag_status", 0) or 0) == 2 and int(e.get("is_ignored", 0) or 0) == 0)
        flag_part = f", {flagged} Flagged" if flagged > 0 else ""
        self.lbl_counts.setText(f"{len(self.emails)} Total ({unread} Unread{flag_part})")

    def _render_messages(self):
        if self._loading:
            return
        query = self.msg_search_edit.text().strip().lower()
        filter_comp = self.company_filter_combo.isFiltered()
        selected_comps = (
            {str(c).strip().lower() for c in self.company_filter_combo.getCheckedItems() if str(c).strip()}
            if filter_comp
            else set()
        )
        state_filter = (self.state_filter_combo.currentData() or "all").strip().lower()

        filtered: List[dict] = []
        for m in self.emails:
            if state_filter == "unread" and int(m.get("is_read", 0) or 0) != 0:
                continue
            if state_filter == "flagged" and int(m.get("flag_status", 0) or 0) != 2:
                continue
            if filter_comp and str(m.get("company") or "").strip().lower() not in selected_comps:
                continue
            if query:
                hay = " ".join([
                    str(m.get("sender_name") or ""),
                    str(m.get("sender_email") or ""),
                    str(m.get("company") or ""),
                    str(m.get("subject") or ""),
                    str(m.get("matched_tdoc") or ""),
                    str(m.get("body_text") or ""),
                ]).lower()
                if query not in hay:
                    continue
            filtered.append(m)

        self._filtered_emails = filtered
        if len(filtered) == len(self.emails):
            self.msg_count_lbl.setText(f"{len(self.emails)} message(s)")
        else:
            self.msg_count_lbl.setText(f"Showing {len(filtered)} of {len(self.emails)}")

        newest_first = bool(self.order_combo.currentData())
        include_meta = self.show_meta_action.isChecked()
        thread_meta = self._build_synthetic_thread_dict()
        md = export_thread_to_markdown(
            thread_meta,
            filtered,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=True,
            newest_first=newest_first,
            include_metadata=include_meta,
            include_ignore_action=True,
            enable_tdoc_links=True,
        )
        self.browser.setMarkdown(md)
        self._unread_blocks = find_unread_heading_blocks(self.browser)
        self._unread_cursor = -1
        self._update_next_unread_button()

    def _update_next_unread_button(self):
        cnt = len(self._unread_blocks)
        self.next_unread_btn.setText(f"🔵 Next Unread ({cnt})")
        self.next_unread_btn.setEnabled(cnt > 0)
        if cnt <= 0:
            self.next_unread_btn.setToolTip("No unread emails in the current view.")
        elif self._unread_cursor >= 0:
            self.next_unread_btn.setToolTip(f"Jump to next unread email ({self._unread_cursor + 1}/{cnt}).")
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

    def _get_cache_dir(self) -> Path:
        d = self.db.db_path.parent / "TDocEmailCache" / self.target_tdoc
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _on_reload_all_from_outlook(self):
        eids = [str(e.get("id") or "") for e in self.emails if e.get("id") and not str(e.get("id")).startswith("MSG-")]
        if not eids:
            notify_info(self, "Reload Emails", "No Outlook-backed emails found to reload.")
            return
        sb = self.browser.verticalScrollBar()
        prev_scroll = sb.value()
        updated = reload_tdoc_emails_from_outlook(self.db, eids)
        self._load_emails()
        sb.setValue(min(prev_scroll, sb.maximum()))
        self.data_changed.emit()
        if updated > 0:
            notify_success(self, "Emails Reloaded", f"Re-parsed {updated} Outlook email(s) (tables, links & attachments).")
        else:
            notify_toast(
                self,
                "Reload Warning",
                "Could not reload emails from Outlook (check Outlook connection or moved items).",
                variant="warning",
            )

    def _mark_all_family_read(self):
        query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
        self.db.set_tdocs_read_status(query_set, True, sync_outlook=True)
        self.data_changed.emit()
        self._load_emails()

    def _on_pin_as_tracked_thread(self):
        if not self._pin_thread_callback or not self.emails:
            return
        latest = self.emails[-1]
        self._pin_thread_callback(
            {
                "subject": latest.get("subject") or f"Discussion on {self.target_tdoc}",
                "normalized_subject": normalize_subject(latest.get("subject") or ""),
                "linked_tdocs": list(self.family_tdocs),
            }
        )

    def _on_cache_all_docs(self):
        cache_dir = self._get_cache_dir()
        cached: List[Path] = []
        for doc_url in extract_thread_document_urls(self.emails):
            p = cache_and_open_thread_document_url(doc_url, cache_dir, open_after=False)
            if p:
                cached.append(p)
        ns = OutlookClient.get_namespace()
        if ns:
            for m in self.emails:
                eid = str(m.get("id") or "").strip()
                if not eid or eid.startswith("MSG-") or not m.get("attachments"):
                    continue
                try:
                    item = ns.GetItemFromID(eid)
                    atts = getattr(item, "Attachments", None)
                    if not atts:
                        continue
                    for i in range(1, atts.Count + 1):
                        att = atts.Item(i)
                        fname = getattr(att, "FileName", "") or ""
                        if not fname or fname.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                            continue
                        safe_name = "".join(c for c in fname if c not in '<>:"/\\|?*').strip()
                        if safe_name:
                            out_path = cache_dir / safe_name
                            if not out_path.exists():
                                att.SaveAsFile(str(out_path.resolve()))
                            cached.append(out_path)
                except Exception:
                    continue
        QMessageBox.information(
            self,
            "Discussion Documents Cached",
            f"Cached {len(cached)} document(s)/attachment(s) into:\n{cache_dir}",
        )

    def _on_open_cache_folder(self):
        folder = self._get_cache_dir()
        if hasattr(os, "startfile"):
            os.startfile(str(folder.resolve()))

    def _populate_llm_export_menu(self):
        menu = self.llm_export_menu
        menu.clear()
        menu.setToolTipsVisible(True)

        act_ai_sum = menu.addAction("🤖 AI Summarize Discussion (Ollama)...")
        act_ai_sum.setToolTip("Generate a structured stance summary using local Ollama and optionally append it to TDoc My Notes.")
        act_ai_sum.triggered.connect(self._on_ai_summarize_clicked)

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
                    f"Prepare discussion Markdown + cached docs and open {st.label} idle "
                    "(prompt copied to clipboard for Ctrl+V — zero auto token burn)."
                )
                act.triggered.connect(lambda _=False, aid=st.agent_id: self._on_launch_cli(aid))
        else:
            no_cli = menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
            no_cli.setEnabled(False)

        menu.addSeparator()
        act_copy = menu.addAction("📋 Copy Discussion Markdown")
        act_copy.triggered.connect(self._on_copy_markdown)
        act_save = menu.addAction("💾 Save Discussion .md...")
        act_save.triggered.connect(self._on_save_markdown)
        menu.addSeparator()
        menu.addAction(self.export_chrono_check)
        menu.addSeparator()
        act_cfg = menu.addAction("⚙️ Configure CLI Agents && MCP...")
        act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(self))

    def _get_export_newest_first(self) -> bool:
        if hasattr(self, "export_chrono_check") and self.export_chrono_check.isChecked():
            return False
        return bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else False

    def _get_markdown(self) -> str:
        return export_thread_to_markdown(
            self._build_synthetic_thread_dict(),
            self.emails,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=False,
            newest_first=self._get_export_newest_first(),
            include_metadata=True,
        )

    def _on_copy_markdown(self):
        QApplication.clipboard().setText(self._get_markdown())
        notify_success(self, "Copied", "Discussion Markdown copied to clipboard.")

    def _on_save_markdown(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc Emails Markdown",
            f"{self.target_tdoc}_emails.md",
            "Markdown (*.md)",
        )
        if path:
            Path(path).write_text(self._get_markdown(), encoding="utf-8")

    def _on_launch_cli(self, agent_id: Optional[str] = None):
        cli = (agent_id or "codex").strip().lower()
        cache_dir = self._get_cache_dir()
        ws_dir = cache_dir / "llm_workspace"
        ws_dir.mkdir(parents=True, exist_ok=True)
        md_path = ws_dir / f"{self.target_tdoc}_emails.md"
        md_path.write_text(self._get_markdown(), encoding="utf-8")

        doc_mds: List[Path] = []
        for existing in cache_dir.iterdir():
            if existing.is_file() and existing.suffix.lower() in (".docx", ".docm"):
                out_md = ws_dir / f"{existing.stem}.md"
                if not out_md.exists():
                    try:
                        from modules.meetings.core.tdoc_markdown_converter import convert_document_to_markdown
                        convert_document_to_markdown(str(existing), str(out_md))
                    except Exception:
                        pass
                if out_md.exists():
                    doc_mds.append(out_md)

        extra_docs = ""
        if doc_mds:
            extra_docs = f" Companion converted documents are also in {ws_dir} ({', '.join(p.name for p in doc_mds[:10])})."
        prompt = (
            f"Analyze the 3GPP email discussion referencing {self.target_tdoc} in {md_path} and summarize "
            f"company positions, key technical objections, revisions, and open points.{extra_docs}"
        )
        write_cli_context_file(
            ws_dir,
            task_mode="custom",
            custom_question=prompt,
            markdown_files=[md_path] + doc_mds,
            briefing_title=f"TDoc Related Emails: {self.target_tdoc}",
        )
        launch_cli_in_terminal(
            cli,
            ws_dir,
            bootstrap_prompt=prompt,
            parent=self,
            auto_execute=False,
        )

    def _on_ai_summarize_clicked(self):
        target_emails = list(self._filtered_emails or self.emails)
        if not target_emails:
            QMessageBox.warning(self, "No Emails", "There are no emails to summarize for this document.")
            return
        if self._summary_dialog and self._summary_dialog.isVisible():
            self._summary_dialog.setWindowState(
                self._summary_dialog.windowState() & ~Qt.WindowMinimized | Qt.WindowActive
            )
            self._summary_dialog.raise_()
            self._summary_dialog.activateWindow()
            return
        self._summary_dialog = EmailSummaryDialog(
            tdoc_id=self.target_tdoc,
            family_tdocs=self.family_tdocs,
            emails_data=target_emails,
            save_callback=self.save_callback,
            parent=None,
        )
        self._summary_dialog.show()

    def _on_anchor_clicked(self, url: QUrl):
        scheme = url.scheme().lower()
        if scheme == "tdoc":
            target_tdoc = url.path().strip() or url.toString().split(":", 1)[1].strip()
            if target_tdoc:
                self.tdoc_selected.emit(target_tdoc.upper())
            return

        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if not path_parts:
                return
            msg_id = path_parts[0]
            sb = self.browser.verticalScrollBar()
            prev_scroll = sb.value()

            if host == "msg":
                if not open_outlook_message(msg_id):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply":
                if not reply_to_outlook_message(msg_id, reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
            elif host == "reload":
                if reload_general_email_from_outlook(self.db, msg_id):
                    self._load_emails()
                    sb.setValue(min(prev_scroll, sb.maximum()))
                    self.data_changed.emit()
                    notify_success(self, "Email Reloaded", "Message body, tables, links, and attachments re-loaded from Outlook.")
                else:
                    notify_toast(self, "Reload Failed", "Could not re-fetch this message from Outlook.", variant="warning")
            elif host in ("read", "unread"):
                self.db.set_emails_read_status([msg_id], is_read=(host == "read"), sync_outlook=True)
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host in ("flag", "unflag"):
                self.db.set_emails_flag_status([msg_id], is_flagged=(host == "flag"), sync_outlook=True)
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host in ("ignore", "unignore"):
                self.db.set_emails_ignored_status([msg_id], is_ignored=(host == "ignore"))
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host == "copy":
                msg = next((m for m in self.emails if str(m.get("id") or "") == msg_id), None)
                if msg:
                    body = (msg.get("clean_body") if self.dedup_check.isChecked() and msg.get("clean_body") else msg.get("body_text")) or ""
                    clip_txt = f"[{msg.get('date_received', '')}] {msg.get('sender_name', '')} ({msg.get('company', '')}):\n{body.strip()}"
                    QApplication.clipboard().setText(clip_txt)
                    notify_info(self, "Copied", "Message copied to clipboard.")
            elif host == "att" and len(path_parts) >= 2:
                att_name = "/".join(path_parts[1:])
                if not open_outlook_attachment(msg_id, att_name, save_dir=self._get_cache_dir()):
                    QMessageBox.warning(self, "Attachment", f"Could not open attachment '{att_name}' from Outlook.")
            return

        url_str = url.toString()
        if is_cacheable_document_url(url_str):
            cached_path = cache_and_open_thread_document_url(url_str, self._get_cache_dir(), open_after=True)
            if cached_path:
                notify_info(self, "3GPP Document Opened", f"Cached and opened {cached_path.name}.")
                return
            notify_toast(
                self,
                "Direct Download Unavailable",
                "Could not download automatically — opening link in browser.",
                variant="warning",
            )
        QDesktopServices.openUrl(url)

    def closeEvent(self, event):
        try:
            self._settings.setValue("geometry", self.saveGeometry())
            if hasattr(self, "browser"):
                self._settings.setValue("zoom_steps", int(self.browser.property("_reader_zoom_steps") or 0))
        except Exception:
            pass
        if getattr(self, "_summary_dialog", None):
            try:
                self._summary_dialog.close()
            except Exception:
                pass
        super().closeEvent(event)
