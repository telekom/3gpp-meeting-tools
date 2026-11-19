# --- File: src/modules/emails/ui/meeting_email_inbox_dialog.py ---
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
import urllib.parse

from PyQt5.QtCore import QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QColor, QDesktopServices
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import install_reader_zoom, zoom_text_widget
from core.ui.toast import notify_info, notify_toast
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.thread_markdown_exporter import (
    escape_non_html_angles,
    inject_html_tables_into_body,
    is_cacheable_document_url,
    linkify_markdown_urls,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    cache_and_open_thread_document_url,
    open_outlook_attachment,
    open_outlook_message,
    reply_to_outlook_message,
)


class MeetingEmailInboxDialog(QDialog):
    """Unified cockpit aggregating unread, flagged, and recent emails across Tracked Threads and Indexed TDoc emails."""

    data_changed = pyqtSignal()

    def __init__(
        self,
        tracked_db: Optional[TrackedThreadsDatabase],
        general_email_db: Optional[Any],
        meeting_id: str = "",
        tdoc_ai_lookup: Optional[Dict[str, str]] = None,
        open_thread_callback: Optional[Callable[[str], None]] = None,
        open_tdoc_callback: Optional[Callable[[str], None]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.tracked_db = tracked_db
        self.general_email_db = general_email_db
        self.meeting_id = str(meeting_id or "").strip()
        self.tdoc_ai_lookup = {str(k).upper(): str(v) for k, v in (tdoc_ai_lookup or {}).items()}
        self._open_thread_callback = open_thread_callback
        self._open_tdoc_callback = open_tdoc_callback

        self._all_items: List[dict] = []
        self._visible_items: List[dict] = []
        self._items_by_id: Dict[str, dict] = {}
        self._current_item: Optional[dict] = None
        self._active_mode = "UNREAD"
        self._settings = QSettings("3GPP_Tools", "MeetingEmailInboxDialog")

        title_suffix = f" — {self.meeting_id}" if self.meeting_id else ""
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        self.setWindowTitle(f"📬 Unread & Flagged Email Inbox{title_suffix}")
        self.setMinimumSize(520, 420)
        self.resize(1220, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        self._build_ui()
        self.reload_from_dbs()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # Row 1: Responsive FlowLayout for State Pills + Search + Company/AI dropdowns + Refresh
        top_widget = QWidget(self)
        top_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        top_sp.setHeightForWidth(True)
        top_widget.setSizePolicy(top_sp)
        top_row = FlowLayout(top_widget, margin=0, h_spacing=6, v_spacing=4)

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        pill_style = (
            "QPushButton { padding: 4px 10px; border-radius: 13px; border: 1px solid #CBD5E1; "
            "background-color: #FFFFFF; color: #334155; font-size: 12px; } "
            "QPushButton:hover { background-color: #F1F5F9; } "
            "QPushButton:checked { background-color: #DBEAFE; border-color: #2563EB; "
            "color: #1D4ED8; font-weight: bold; }"
        )

        self.pill_unread = QPushButton("🔵 Unread Only (0)")
        self.pill_unread.setCheckable(True)
        self.pill_unread.setChecked(True)
        self.pill_unread.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_unread)
        self.pill_unread.clicked.connect(lambda: self._set_mode("UNREAD"))
        top_row.addWidget(self.pill_unread)

        self.pill_flagged = QPushButton("🚩 Flagged (0)")
        self.pill_flagged.setCheckable(True)
        self.pill_flagged.setStyleSheet(
            "QPushButton { padding: 4px 10px; border-radius: 13px; border: 1px solid #CBD5E1; "
            "background-color: #FFFFFF; color: #334155; font-size: 12px; } "
            "QPushButton:hover { background-color: #FEF2F2; } "
            "QPushButton:checked { background-color: #FEE2E2; border-color: #DC2626; "
            "color: #B91C1C; font-weight: bold; }"
        )
        self.mode_group.addButton(self.pill_flagged)
        self.pill_flagged.clicked.connect(lambda: self._set_mode("FLAGGED"))
        top_row.addWidget(self.pill_flagged)

        self.pill_recent = QPushButton("🕒 Last 24h (0)")
        self.pill_recent.setCheckable(True)
        self.pill_recent.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_recent)
        self.pill_recent.clicked.connect(lambda: self._set_mode("RECENT"))
        top_row.addWidget(self.pill_recent)

        self.pill_all = QPushButton("📋 All (0)")
        self.pill_all.setCheckable(True)
        self.pill_all.setStyleSheet(pill_style)
        self.mode_group.addButton(self.pill_all)
        self.pill_all.clicked.connect(lambda: self._set_mode("ALL"))
        top_row.addWidget(self.pill_all)

        self.search_edit = QLineEdit()
        self.search_edit.setMinimumWidth(200)
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setPlaceholderText("🔍 Search subject, sender, company, TDoc, or body...")
        self.search_edit.textChanged.connect(self._apply_filters)
        top_row.addWidget(self.search_edit)

        self.company_combo = QComboBox()
        self.company_combo.addItem("All Companies", "")
        self.company_combo.currentIndexChanged.connect(self._apply_filters)
        top_row.addWidget(self.company_combo)

        self.ai_combo = QComboBox()
        self.ai_combo.addItem("All AIs", "")
        self.ai_combo.currentIndexChanged.connect(self._apply_filters)
        top_row.addWidget(self.ai_combo)

        self.refresh_btn = QPushButton("🔄 Reload")
        self.refresh_btn.setToolTip("Reload latest synced messages from local meeting databases.")
        self.refresh_btn.clicked.connect(self.reload_from_dbs)
        top_row.addWidget(self.refresh_btn)

        layout.addWidget(top_widget)

        # Row 2: Master-Detail Splitter
        self.splitter = QSplitter(Qt.Horizontal, self)
        self.splitter.setChildrenCollapsible(False)

        left_widget = QWidget()
        left_widget.setMinimumWidth(240)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self.table = QTableWidget(0, 6, left_widget)
        self.table.setHorizontalHeaderLabels([
            "State",
            "Source / Context",
            "AI",
            "Subject",
            "Sender / Company",
            "Date",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.resizeSection(0, 52)
        hdr.resizeSection(1, 165)
        hdr.resizeSection(2, 68)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        hdr.resizeSection(4, 140)
        hdr.resizeSection(5, 115)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        left_layout.addWidget(self.table, 1)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
        left_layout.addWidget(self.status_lbl)

        # Right Detail Pane
        right_widget = QWidget()
        right_widget.setMinimumWidth(260)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 0, 0, 0)
        right_layout.setSpacing(6)

        meta_frame = QFrame(right_widget)
        meta_frame.setObjectName("InboxMetaFrame")
        meta_frame.setStyleSheet(
            "#InboxMetaFrame { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }"
        )
        meta_layout = QVBoxLayout(meta_frame)
        meta_layout.setContentsMargins(10, 8, 10, 8)
        meta_layout.setSpacing(3)

        self.detail_subject_lbl = QLabel("Select an email to read")
        self.detail_subject_lbl.setWordWrap(True)
        self.detail_subject_lbl.setStyleSheet("font-size: 13px; font-weight: bold; color: #0F172A;")
        meta_layout.addWidget(self.detail_subject_lbl)

        self.detail_meta_lbl = QLabel("")
        self.detail_meta_lbl.setWordWrap(True)
        self.detail_meta_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        meta_layout.addWidget(self.detail_meta_lbl)

        self.detail_links_lbl = QLabel("")
        self.detail_links_lbl.setWordWrap(True)
        self.detail_links_lbl.setStyleSheet("color: #1D4ED8; font-size: 11px;")
        meta_layout.addWidget(self.detail_links_lbl)

        right_layout.addWidget(meta_frame)

        # Triage Action Bar (FlowLayout wraps cleanly when pane narrows)
        act_widget = QWidget(right_widget)
        act_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        act_sp.setHeightForWidth(True)
        act_widget.setSizePolicy(act_sp)
        act_row = FlowLayout(act_widget, margin=0, h_spacing=6, v_spacing=4)

        self.btn_mark_next = QPushButton("✔️ Mark Read && Next Unread")
        self.btn_mark_next.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; font-weight: bold; "
            "padding: 5px 11px; border-radius: 5px; border: none; } "
            "QPushButton:hover { background-color: #1D4ED8; } "
            "QPushButton:disabled { background-color: #CBD5E1; color: #64748B; }"
        )
        self.btn_mark_next.clicked.connect(self._on_mark_read_and_next)
        act_row.addWidget(self.btn_mark_next)

        self.btn_toggle_read = QPushButton("✓ Mark Read")
        self.btn_toggle_read.clicked.connect(self._on_toggle_read)
        act_row.addWidget(self.btn_toggle_read)

        self.btn_toggle_flag = QPushButton("🚩 Flag")
        self.btn_toggle_flag.clicked.connect(self._on_toggle_flag)
        act_row.addWidget(self.btn_toggle_flag)

        self.btn_reply = QPushButton("↩️ Reply")
        self.btn_reply.clicked.connect(self._on_reply)
        act_row.addWidget(self.btn_reply)

        self.btn_open_outlook = QPushButton("📧 Outlook")
        self.btn_open_outlook.clicked.connect(self._on_open_outlook)
        act_row.addWidget(self.btn_open_outlook)

        self.btn_open_context = QPushButton("📂 Open Thread")
        self.btn_open_context.clicked.connect(self._on_open_context)
        act_row.addWidget(self.btn_open_context)

        right_layout.addWidget(act_widget)

        self.preview_browser = QTextBrowser()
        self.preview_browser.setOpenLinks(False)
        self.preview_browser.setToolTip("Tip: Hold Ctrl + Mouse Wheel to zoom text size (Ctrl+0 to reset).")
        self.preview_browser.anchorClicked.connect(self._on_anchor_clicked)
        install_reader_zoom()
        saved_zoom = self._settings.value("zoom_steps", 0, type=int)
        if saved_zoom:
            zoom_text_widget(self.preview_browser, int(saved_zoom))
        right_layout.addWidget(self.preview_browser, 1)

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([580, 620])
        saved_split = self._settings.value("splitter_state")
        if saved_split:
            try:
                self.splitter.restoreState(saved_split)
            except Exception:
                pass
        layout.addWidget(self.splitter, 1)

    def closeEvent(self, event):
        try:
            self._settings.setValue("geometry", self.saveGeometry())
            if hasattr(self, "splitter"):
                self._settings.setValue("splitter_state", self.splitter.saveState())
            if hasattr(self, "preview_browser"):
                self._settings.setValue("zoom_steps", int(self.preview_browser.property("_reader_zoom_steps") or 0))
        except Exception:
            pass
        super().closeEvent(event)

    def _set_mode(self, mode: str):
        self._active_mode = str(mode or "UNREAD").upper()
        self._apply_filters()

    def reload_from_dbs(self, preserve_id: Optional[str] = None):
        """Loads and deduplicates emails from TrackedThreadsDatabase and GeneralEmailDatabase."""
        by_id: Dict[str, dict] = {}

        if self.tracked_db is not None:
            threads = self.tracked_db.get_threads_for_meeting(self.meeting_id)
            for th in threads:
                tid = str(th.get("thread_id") or "")
                th_title = str(th.get("title") or "").strip()
                th_ais = [str(a).strip() for a in (th.get("linked_ais") or []) if str(a).strip()]
                th_tdocs = [str(t).strip().upper() for t in (th.get("linked_tdocs") or []) if str(t).strip()]
                for m in self.tracked_db.get_thread_messages(tid):
                    eid = str(m.get("id") or "").strip()
                    if not eid:
                        continue
                    by_id[eid] = {
                        "id": eid,
                        "source_kind": "thread",
                        "thread_id": tid,
                        "thread_title": th_title,
                        "subject": str(m.get("subject") or th_title or "(No Subject)").strip(),
                        "sender_name": str(m.get("sender_name") or "").strip(),
                        "sender_email": str(m.get("sender_email") or "").strip(),
                        "company": str(m.get("company") or "").strip(),
                        "date_str": str(m.get("date_sent") or "").strip(),
                        "body_text": str(m.get("clean_body") or m.get("body_text") or ""),
                        "raw_body_text": str(m.get("body_text") or ""),
                        "html_body": "",
                        "is_read": int(m.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(m.get("flag_status", 0) or 0) == 2,
                        "attachments": list(m.get("attachments") or []),
                        "linked_ais": list(th_ais),
                        "linked_tdocs": list(th_tdocs),
                        "in_tracked_db": True,
                        "in_general_db": False,
                    }

        if self.general_email_db is not None and hasattr(self.general_email_db, "get_all_emails_with_tdocs"):
            for row in self.general_email_db.get_all_emails_with_tdocs(show_ignored=False):
                eid = str(row.get("id") or "").strip()
                if not eid:
                    continue
                matched_tdocs = [t.strip().upper() for t in str(row.get("matched_tdocs") or "").split(",") if t.strip()]
                inferred_ais = []
                for td in matched_tdocs:
                    ai = self.tdoc_ai_lookup.get(td, "")
                    if ai and ai not in inferred_ais:
                        inferred_ais.append(ai)

                if eid in by_id:
                    item = by_id[eid]
                    item["in_general_db"] = True
                    for td in matched_tdocs:
                        if td not in item["linked_tdocs"]:
                            item["linked_tdocs"].append(td)
                    for ai in inferred_ais:
                        if ai not in item["linked_ais"]:
                            item["linked_ais"].append(ai)
                    if row.get("html_body") and not item.get("html_body"):
                        item["html_body"] = str(row["html_body"])
                    comp_val = str(row.get("company") or row.get("sender_company") or "").strip()
                    if not item.get("company") and comp_val:
                        item["company"] = comp_val
                else:
                    by_id[eid] = {
                        "id": eid,
                        "source_kind": "tdoc",
                        "thread_id": "",
                        "thread_title": "",
                        "subject": str(row.get("subject") or "(No Subject)").strip(),
                        "sender_name": str(row.get("sender_name") or "").strip(),
                        "sender_email": str(row.get("sender_email") or "").strip(),
                        "company": str(row.get("company") or row.get("sender_company") or "").strip(),
                        "date_str": str(row.get("date_received") or "").strip(),
                        "body_text": str(row.get("body_text") or ""),
                        "raw_body_text": str(row.get("body_text") or ""),
                        "html_body": str(row.get("html_body") or ""),
                        "is_read": int(row.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(row.get("flag_status", 0) or 0) == 2,
                        "attachments": [],
                        "linked_ais": inferred_ais,
                        "linked_tdocs": matched_tdocs,
                        "in_tracked_db": False,
                        "in_general_db": True,
                    }

        self._all_items = sorted(by_id.values(), key=lambda x: x.get("date_str", ""), reverse=True)
        self._items_by_id = {it["id"]: it for it in self._all_items}

        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters(preserve_id=preserve_id)

    def _populate_filter_combos(self):
        cur_comp = self.company_combo.currentData() or ""
        cur_ai = self.ai_combo.currentData() or ""

        companies = sorted({it["company"] for it in self._all_items if it.get("company")}, key=str.lower)
        ais = sorted({ai for it in self._all_items for ai in (it.get("linked_ais") or []) if ai})

        self.company_combo.blockSignals(True)
        self.company_combo.clear()
        self.company_combo.addItem("All Companies", "")
        for c in companies:
            self.company_combo.addItem(c, c)
        c_idx = self.company_combo.findData(cur_comp)
        self.company_combo.setCurrentIndex(max(0, c_idx))
        self.company_combo.blockSignals(False)

        self.ai_combo.blockSignals(True)
        self.ai_combo.clear()
        self.ai_combo.addItem("All AIs", "")
        for ai in ais:
            self.ai_combo.addItem(ai, ai)
        a_idx = self.ai_combo.findData(cur_ai)
        self.ai_combo.setCurrentIndex(max(0, a_idx))
        self.ai_combo.blockSignals(False)

    @staticmethod
    def _recent_cutoff_str() -> str:
        return (datetime.now() - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M")

    def _update_pill_counts(self):
        cutoff = self._recent_cutoff_str()
        unread_cnt = sum(1 for it in self._all_items if not it["is_read"])
        flagged_cnt = sum(1 for it in self._all_items if it["is_flagged"])
        recent_cnt = sum(1 for it in self._all_items if str(it.get("date_str") or "") >= cutoff)
        total_cnt = len(self._all_items)

        self.pill_unread.setText(f"🔵 Unread Only ({unread_cnt})")
        self.pill_flagged.setText(f"🚩 Flagged ({flagged_cnt})")
        self.pill_recent.setText(f"🕒 Last 24h ({recent_cnt})")
        self.pill_all.setText(f"📋 All ({total_cnt})")

    def _apply_filters(self, *_args, preserve_id: Optional[str] = None, select_row: Optional[int] = None):
        target_id = preserve_id if preserve_id is not None else (self._current_item["id"] if self._current_item else None)
        mode = self._active_mode
        cutoff = self._recent_cutoff_str()
        q = self.search_edit.text().strip().lower()
        comp_filter = (self.company_combo.currentData() or "").strip().lower()
        ai_filter = (self.ai_combo.currentData() or "").strip()

        visible: List[dict] = []
        for it in self._all_items:
            if mode == "UNREAD" and it["is_read"]:
                continue
            if mode == "FLAGGED" and not it["is_flagged"]:
                continue
            if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
                continue
            if comp_filter and str(it.get("company") or "").strip().lower() != comp_filter:
                continue
            if ai_filter and ai_filter not in (it.get("linked_ais") or []):
                continue
            if q:
                hay = " ".join([
                    str(it.get("subject") or ""),
                    str(it.get("thread_title") or ""),
                    str(it.get("sender_name") or ""),
                    str(it.get("company") or ""),
                    " ".join(it.get("linked_tdocs") or []),
                    " ".join(it.get("linked_ais") or []),
                    str(it.get("body_text") or "")[:1200],
                ]).lower()
                if q not in hay:
                    continue
            visible.append(it)

        self._visible_items = visible
        self.table.blockSignals(True)
        self.table.setRowCount(len(visible))
        row_to_select = -1
        for r, it in enumerate(visible):
            state_parts = []
            if not it["is_read"]:
                state_parts.append("🔵")
            if it["is_flagged"]:
                state_parts.append("🚩")
            if not state_parts:
                state_parts.append("✓")
            state_item = QTableWidgetItem(" ".join(state_parts))
            state_item.setData(Qt.UserRole, it["id"])
            state_item.setTextAlignment(Qt.AlignCenter)

            if it["in_tracked_db"] and it.get("thread_title"):
                ctx = f"🧵 {it['thread_title'][:34]}"
                if len(it["thread_title"]) > 34:
                    ctx += "…"
            elif it.get("linked_tdocs"):
                ctx = "📄 " + ", ".join(it["linked_tdocs"][:2])
            else:
                ctx = "📧 Email"
            ctx_item = QTableWidgetItem(ctx)

            ai_item = QTableWidgetItem(", ".join(it.get("linked_ais") or []) or "—")

            subj_item = QTableWidgetItem(str(it.get("subject") or ""))
            if not it["is_read"]:
                f = subj_item.font()
                f.setBold(True)
                subj_item.setFont(f)

            sender_disp = it["sender_name"] or it["sender_email"] or "Unknown"
            if it.get("company"):
                sender_disp = f"{sender_disp} ({it['company']})"
            sender_item = QTableWidgetItem(sender_disp)
            date_item = QTableWidgetItem(str(it.get("date_str") or "")[:16])

            for cell in (state_item, ctx_item, ai_item, subj_item, sender_item, date_item):
                if not it["is_read"]:
                    cell.setBackground(QColor("#F8FAFF"))

            self.table.setItem(r, 0, state_item)
            self.table.setItem(r, 1, ctx_item)
            self.table.setItem(r, 2, ai_item)
            self.table.setItem(r, 3, subj_item)
            self.table.setItem(r, 4, sender_item)
            self.table.setItem(r, 5, date_item)

            if target_id and it["id"] == target_id:
                row_to_select = r

        if visible:
            if select_row is not None and 0 <= select_row < len(visible):
                row_to_select = select_row
            elif select_row is not None and select_row >= len(visible):
                row_to_select = len(visible) - 1
            elif row_to_select < 0:
                row_to_select = 0
            self.table.selectRow(row_to_select)
        self.table.blockSignals(False)

        self.status_lbl.setText(f"Showing {len(visible)} of {len(self._all_items)} indexed email(s)")
        self._on_selection_changed()

    def _on_selection_changed(self):
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}) if self.table.selectionModel() else []
        if not rows or rows[0] < 0 or rows[0] >= len(self._visible_items):
            self._current_item = None
            self._render_empty_detail()
            return
        self._current_item = self._visible_items[rows[0]]
        self._render_detail(self._current_item)

    def _render_empty_detail(self):
        self.detail_subject_lbl.setText("No matching emails in this filter")
        self.detail_meta_lbl.setText("")
        self.detail_links_lbl.setText("")
        self.preview_browser.setMarkdown("*Select another filter tab or adjust search criteria.*")
        for btn in (self.btn_mark_next, self.btn_toggle_read, self.btn_toggle_flag, self.btn_reply, self.btn_open_outlook, self.btn_open_context):
            btn.setEnabled(False)

    def _render_detail(self, it: dict):
        self.detail_subject_lbl.setText(it.get("subject") or "(No Subject)")
        sender = it.get("sender_name") or it.get("sender_email") or "Unknown"
        comp = f" ({it['company']})" if it.get("company") else ""
        badges = []
        if not it["is_read"]:
            badges.append("🔵 UNREAD")
        if it["is_flagged"]:
            badges.append("🚩 FLAGGED")
        badge_str = f"  ·  {' · '.join(badges)}" if badges else "  ·  ✓ Read"
        self.detail_meta_lbl.setText(f"From: {sender}{comp}  |  Date: {it.get('date_str') or 'Unknown'}{badge_str}")

        links = []
        if it.get("thread_title"):
            links.append(f"🧵 Thread: {it['thread_title']}")
        if it.get("linked_ais"):
            links.append("AIs: " + ", ".join(it["linked_ais"]))
        if it.get("linked_tdocs"):
            links.append("TDocs: " + ", ".join(it["linked_tdocs"][:6]))
        self.detail_links_lbl.setText("  |  ".join(links))
        self.detail_links_lbl.setVisible(bool(links))

        can_outlook = not str(it.get("id") or "").startswith("MSG-")
        has_context = bool(
            (it.get("thread_id") and self._open_thread_callback)
            or (it.get("linked_tdocs") and self._open_tdoc_callback)
        )
        self.btn_mark_next.setEnabled(True)
        self.btn_toggle_read.setEnabled(True)
        self.btn_toggle_read.setText("✓ Mark Read" if not it["is_read"] else "🔵 Mark Unread")
        self.btn_toggle_flag.setEnabled(True)
        self.btn_toggle_flag.setText("🏳️ Unflag" if it["is_flagged"] else "🚩 Flag")
        self.btn_reply.setEnabled(can_outlook)
        self.btn_open_outlook.setEnabled(can_outlook)
        self.btn_open_context.setEnabled(has_context)
        if it.get("thread_id") and self._open_thread_callback:
            self.btn_open_context.setText("📂 Open Full Thread")
        elif it.get("linked_tdocs") and self._open_tdoc_callback:
            self.btn_open_context.setText(f"📄 Open {it['linked_tdocs'][0]}")
        else:
            self.btn_open_context.setText("📂 Open Context")

        lines: List[str] = []
        atts = it.get("attachments") or []
        if atts and can_outlook:
            enc_id = urllib.parse.quote(it["id"], safe="")
            att_links = [
                f"[📎 {a}](outlook://att/{enc_id}/{urllib.parse.quote(str(a), safe='')})"
                for a in atts
            ]
            lines.append(f"**Attachments:** {' · '.join(att_links)}\n")

        body = str(it.get("body_text") or "")
        if it.get("html_body"):
            body = inject_html_tables_into_body(body, str(it["html_body"]))
        body = linkify_markdown_urls(body.strip())
        body = escape_non_html_angles(body)
        lines.append(body if body else "*(Empty message body)*")
        self.preview_browser.setMarkdown("\n\n".join(lines))

    def _mutate_read(self, it: dict, is_read: bool):
        eid = it["id"]
        if self.tracked_db is not None:
            self.tracked_db.set_message_read_status(eid, is_read)
        if self.general_email_db is not None and hasattr(self.general_email_db, "set_emails_read_status"):
            self.general_email_db.set_emails_read_status([eid], is_read, sync_outlook=False)
        if not eid.startswith("MSG-"):
            OutlookClient.set_messages_read_status([eid], is_read=is_read)
        it["is_read"] = bool(is_read)
        self._update_pill_counts()
        self.data_changed.emit()

    def _mutate_flag(self, it: dict, is_flagged: bool):
        eid = it["id"]
        if self.tracked_db is not None:
            self.tracked_db.set_message_flag_status(eid, is_flagged)
        if self.general_email_db is not None and hasattr(self.general_email_db, "set_emails_flag_status"):
            self.general_email_db.set_emails_flag_status([eid], is_flagged, sync_outlook=False)
        if not eid.startswith("MSG-"):
            OutlookClient.set_messages_flag_status([eid], is_flagged=is_flagged)
        it["is_flagged"] = bool(is_flagged)
        self._update_pill_counts()
        self.data_changed.emit()

    def _on_mark_read_and_next(self):
        if not self._current_item:
            return
        cur_row = self.table.currentRow()
        it = self._current_item
        if not it["is_read"]:
            self._mutate_read(it, True)
            if self._active_mode == "UNREAD":
                self._apply_filters(preserve_id="", select_row=max(0, cur_row))
                return
        # Move to next unread row in current filtered view
        if self._visible_items:
            for offset in range(1, len(self._visible_items) + 1):
                idx = (max(0, cur_row) + offset) % len(self._visible_items)
                if not self._visible_items[idx]["is_read"]:
                    self._apply_filters(preserve_id=self._visible_items[idx]["id"])
                    return
            next_row = min(max(0, cur_row) + 1, len(self._visible_items) - 1)
            self._apply_filters(preserve_id=self._visible_items[next_row]["id"])

    def _on_toggle_read(self):
        if not self._current_item:
            return
        it = self._current_item
        cur_row = self.table.currentRow()
        self._mutate_read(it, not it["is_read"])
        self._apply_filters(
            preserve_id=it["id"],
            select_row=cur_row if self._active_mode == "UNREAD" else None,
        )

    def _on_toggle_flag(self):
        if not self._current_item:
            return
        it = self._current_item
        cur_row = self.table.currentRow()
        self._mutate_flag(it, not it["is_flagged"])
        self._apply_filters(
            preserve_id=it["id"],
            select_row=cur_row if self._active_mode == "FLAGGED" else None,
        )

    def _on_reply(self):
        if not self._current_item:
            return
        if not reply_to_outlook_message(self._current_item["id"], reply_all=True):
            QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")

    def _on_open_outlook(self):
        if not self._current_item:
            return
        if not open_outlook_message(self._current_item["id"]):
            QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")

    def _on_open_context(self):
        if not self._current_item:
            return
        it = self._current_item
        if it.get("thread_id") and self._open_thread_callback:
            self._open_thread_callback(it["thread_id"])
            self.reload_from_dbs(preserve_id=it["id"])
        elif it.get("linked_tdocs") and self._open_tdoc_callback:
            self._open_tdoc_callback(it["linked_tdocs"][0])
            self.reload_from_dbs(preserve_id=it["id"])

    def _on_anchor_clicked(self, url: QUrl):
        scheme = url.scheme().lower()
        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if host == "att" and len(path_parts) >= 2:
                msg_id = path_parts[0]
                att_name = "/".join(path_parts[1:])
                cache_dir = Path(self.tracked_db.db_path).parent / "InboxAttachments" if self.tracked_db else Path.home() / ".3gpp_tools_inbox"
                cache_dir.mkdir(parents=True, exist_ok=True)
                ok = open_outlook_attachment(msg_id, att_name, cache_dir)
                if not ok:
                    QMessageBox.warning(self, "Attachment", f"Could not extract attachment '{att_name}' from Outlook.")
            return

        if scheme in ("http", "https"):
            raw_url = url.toString()
            if is_cacheable_document_url(raw_url) and self.tracked_db is not None:
                cache_dir = Path(self.tracked_db.db_path).parent / "InboxAttachments"
                cache_dir.mkdir(parents=True, exist_ok=True)
                local_path = cache_and_open_thread_document_url(raw_url, cache_dir, open_after=True)
                if local_path:
                    notify_info(self, "3GPP Document Opened", f"Cached and opened {local_path.name}.")
                    return
                notify_toast(
                    self,
                    "Direct Download Unavailable",
                    "Could not download automatically — opening link in browser.",
                    variant="warning",
                )
            QDesktopServices.openUrl(url)
