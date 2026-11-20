# --- File: src/modules/emails/ui/meeting_email_inbox_dialog.py ---
from datetime import datetime, timedelta
import html
import json
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional
import urllib.parse

from PyQt5.QtCore import QItemSelectionModel, QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtGui import QColor, QDesktopServices, QPainter
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
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill, install_reader_zoom, scale_px, zoom_text_widget
from core.ui.theme import C
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.thread_markdown_exporter import (
    _format_trailing_quoted_history,
    apply_email_reader_document_style,
    clean_tdoc_id,
    escape_non_html_angles,
    export_thread_to_markdown,
    find_unread_heading_blocks,
    format_compact_timestamp,
    format_sender_display,
    inject_html_tables_into_body,
    is_cacheable_document_url,
    linkify_markdown_urls,
    linkify_tdoc_tokens,
    partition_email_attachments,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    auto_suggest_links,
    cache_and_open_thread_document_url,
    normalize_subject,
    open_outlook_attachment,
    open_outlook_message,
    pin_thread_by_subject_or_cid,
    reload_outlook_message,
    reply_to_outlook_message,
)
from modules.meetings.ui.topic_pills import TagCloudWidget


_STATUS_COLOR_MAP = {
    "🔵 My TDoc": C.ACCENT_ALT,
    "🟢 Support": C.SUCCESS,
    "🔴 Object": C.DANGER_ALT,
    "🟡 Monitor": C.WARNING_ACCENT,
}

_STATUS_BADGE_BG_MAP = {
    "🔵 My TDoc": C.INFO_SOFT,
    "🟢 Support": C.SUCCESS_SOFT,
    "🔴 Object": C.DANGER_SOFT,
    "🟡 Monitor": C.WARNING_SOFT,
}


def _status_color_hex(status: str) -> str:
    return _STATUS_COLOR_MAP.get(str(status or "⚪ Neutral").strip(), "")


def get_compact_action_bar_qss(container_name: str = "InboxActionBar") -> str:
    """Shared compact button styling for email thread/TDoc reading panels."""
    return (
        f"#{container_name} QPushButton {{ padding: 3px 8px; border-radius: 4px; "
        f"border: 1px solid {C.BORDER_STRONG}; background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; }} "
        f"#{container_name} QPushButton:hover {{ background-color: {C.BG_HOVER}; }} "
        f"#{container_name} QPushButton:checked {{ background-color: {C.INFO_SOFT}; border-color: {C.ACCENT_ALT}; color: {C.INFO}; font-weight: bold; }} "
        f"#{container_name} QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}"
    )


def get_status_combo_qss(status_val: str) -> str:
    """Shared compact tint QSS for My Status selector combos."""
    st = str(status_val or "⚪ Neutral").strip() or "⚪ Neutral"
    accent = _STATUS_COLOR_MAP.get(st, C.BORDER_STRONG)
    bg = _STATUS_BADGE_BG_MAP.get(st, C.BG_CARD)
    return (
        f"QComboBox {{ background-color: {bg}; color: {C.TEXT_PRIMARY}; "
        f"border: 1px solid {accent}; border-radius: 4px; padding: 2px 6px; "
        f"font-weight: {'bold' if st != '⚪ Neutral' else 'normal'}; }}"
    )


class _InboxStatusStripDelegate(QStyledItemDelegate):
    """Paints a crisp left-edge status color stripe on the State column without widening the table."""

    def paint(self, painter: QPainter, option, index):
        super().paint(painter, option, index)
        color_hex = str(index.data(Qt.UserRole + 1) or "").strip()
        if not color_hex:
            return
        painter.save()
        bar_w = scale_px(5, minimum=4)
        rect = option.rect
        painter.fillRect(rect.left() + 1, rect.top() + 2, bar_w, max(4, rect.height() - 4), QColor(color_hex))
        painter.restore()


class _NotePlainTextEdit(QPlainTextEdit):
    """Compact inline note editor that emits save signals on blur or Ctrl+Enter."""

    focusLost = pyqtSignal()
    saveRequested = pyqtSignal()

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.focusLost.emit()

    def keyPressEvent(self, event):
        if (event.modifiers() & Qt.ControlModifier) and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.saveRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class _InboxTableWidget(QTableWidget):
    """QTableWidget with ExtendedSelection, deterministic selectRow(), and viewport resize signal."""

    viewportResized = pyqtSignal()

    def selectRow(self, row: int):
        sel_model = self.selectionModel()
        model = self.model()
        if sel_model is None or model is None or row < 0 or row >= self.rowCount():
            return
        idx = model.index(row, 0)
        sel_model.setCurrentIndex(idx, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.viewportResized.emit()


class MeetingEmailInboxDialog(QDialog):
    """Unified cockpit for triaging unread, flagged, unassigned, and unpinned emails/threads."""

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
        *,
        wg: str = "SA2",
        ai_suggestions: Optional[List[str]] = None,
        wi_suggestions: Optional[List[str]] = None,
        ai_to_sub_items: Optional[Dict[str, List[str]]] = None,
        tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
        agenda_map: Optional[Dict[str, Any]] = None,
        settings: Optional[QSettings] = None,
        meeting_start: str = "",
        meeting_end: str = "",
    ):
        super().__init__(parent)
        self.tracked_db = tracked_db
        self.general_email_db = general_email_db
        self.meeting_id = str(meeting_id or "").strip()
        self.wg = str(wg or "SA2").strip().upper()
        self._meeting_start = str(meeting_start or "").strip()
        self._meeting_end = str(meeting_end or "").strip()
        self._sync_worker = None
        self.tdoc_ai_lookup = {str(k).upper(): str(v) for k, v in (tdoc_ai_lookup or {}).items()}
        self._tdoc_meta_map = tdoc_meta_map or {}
        self._agenda_map = agenda_map
        self._ai_suggestions = list(ai_suggestions or [])
        self._wi_suggestions = list(wi_suggestions or [])
        self._valid_wis = [s.split(" — ", 1)[0].strip() for s in self._wi_suggestions if str(s).strip()]
        self._ai_to_sub_items: Dict[str, List[str]] = {
            str(k).strip(): list(dict.fromkeys(str(v).strip() for v in vals if str(v).strip()))
            for k, vals in (ai_to_sub_items or {}).items()
            if str(k).strip()
        }
        self._ai_label_map: Dict[str, str] = {}
        for item in self._ai_suggestions:
            code = str(item).split(" — ", 1)[0].strip()
            if code and code not in self._ai_label_map:
                self._ai_label_map[code] = str(item)

        self._open_thread_callback = open_thread_callback
        self._open_tdoc_callback = open_tdoc_callback

        self._all_items: List[dict] = []
        self._grouped_items: List[dict] = []
        self._visible_items: List[dict] = []
        self._items_by_id: Dict[str, dict] = {}
        self._selected_items: List[dict] = []
        self._current_item: Optional[dict] = None
        self._multi_base_common: Dict[str, List[str]] = {"ais": [], "subs": [], "wis": []}
        self._settings = settings if settings is not None else QSettings("3GPP_Tools", "MeetingEmailInboxDialog")
        self._restoring_filters = True
        self._active_mode = "UNREAD"
        self._view_mode = "MESSAGES"
        self._saved_source_filter = ""
        self._saved_search_text = ""
        self._pending_company_filter: Optional[List[str]] = None
        self._pending_ai_filter: Optional[List[str]] = None
        self._pending_status_filter: Optional[List[str]] = None
        self._load_filter_settings()
        self._layout_orientation = "V" if self._settings.value("layout_orientation", "H") == "V" else "H"
        self._updating_assign_bar = False
        self._updating_status_bar = False
        self._updating_note_bar = False
        self._restoring_header = False
        self._note_target_item: Optional[dict] = None
        self._note_dirty = False
        self._note_save_timer = QTimer(self)
        self._note_save_timer.setSingleShot(True)
        self._note_save_timer.setInterval(450)
        self._note_save_timer.timeout.connect(self._flush_pending_note_save)

        title_suffix = f" — {self.meeting_id}" if self.meeting_id else ""
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        self.setWindowTitle(f"📬 Unread & Flagged Email Inbox{title_suffix}")
        self.setMinimumSize(520, 420)
        self.resize(1260, 780)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        try:
            self._build_ui()
            self.reload_from_dbs()
        finally:
            self._restoring_filters = False

    def _meeting_settings_prefix(self) -> str:
        raw = self.meeting_id or "_global"
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", raw).strip("_") or "_global"
        return f"meetings/{safe}"

    @staticmethod
    def _parse_json_str_list(raw_val: Any) -> List[str]:
        if isinstance(raw_val, list):
            return [str(x).strip() for x in raw_val if str(x).strip()]
        text = str(raw_val or "").strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
        return []

    def _load_filter_settings(self):
        prefix = self._meeting_settings_prefix()
        mode = str(self._settings.value(f"{prefix}/active_mode", "UNREAD") or "UNREAD").strip().upper()
        self._active_mode = mode if mode in ("UNREAD", "FLAGGED", "RECENT", "ALL") else "UNREAD"

        raw_vm = self._settings.value(f"{prefix}/view_mode", None)
        if raw_vm is None:
            raw_vm = self._settings.value("view_mode", "MESSAGES")
        self._view_mode = "THREADS" if str(raw_vm or "").strip().upper() == "THREADS" else "MESSAGES"

        src = str(self._settings.value(f"{prefix}/source_filter", "") or "").strip().upper()
        self._saved_source_filter = src if src in ("", "NEEDS_TRIAGE", "THREAD", "UNPINNED") else ""
        self._saved_search_text = str(self._settings.value(f"{prefix}/search_text", "") or "")

        comp_filtered = self._settings.value(f"{prefix}/company_filtered", False, type=bool)
        comp_list = self._parse_json_str_list(self._settings.value(f"{prefix}/company_checked", "[]"))
        self._pending_company_filter = comp_list if (comp_filtered and comp_list) else None

        ai_filtered = self._settings.value(f"{prefix}/ai_filtered", False, type=bool)
        ai_list = self._parse_json_str_list(self._settings.value(f"{prefix}/ai_checked", "[]"))
        self._pending_ai_filter = ai_list if (ai_filtered and ai_list) else None

        status_filtered = self._settings.value(f"{prefix}/status_filtered", False, type=bool)
        status_list = self._parse_json_str_list(self._settings.value(f"{prefix}/status_checked", "[]"))
        self._pending_status_filter = status_list if (status_filtered and status_list) else None

    def _save_filter_settings(self):
        if getattr(self, "_restoring_filters", False):
            return
        try:
            prefix = self._meeting_settings_prefix()
            self._settings.setValue(f"{prefix}/active_mode", self._active_mode)
            self._settings.setValue(f"{prefix}/view_mode", self._view_mode)
            if hasattr(self, "source_combo"):
                self._settings.setValue(f"{prefix}/source_filter", str(self.source_combo.currentData() or ""))
            if hasattr(self, "search_edit"):
                self._settings.setValue(f"{prefix}/search_text", self.search_edit.text())
            if hasattr(self, "company_combo"):
                comp_filtered = self.company_combo.isFiltered()
                comp_checked = (
                    [str(c) for c in self.company_combo.getCheckedItems() if str(c).strip()]
                    if comp_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/company_filtered", bool(comp_filtered and comp_checked))
                self._settings.setValue(f"{prefix}/company_checked", json.dumps(comp_checked))
            if hasattr(self, "ai_combo"):
                ai_filtered = self.ai_combo.isFiltered()
                ai_checked = (
                    [str(a) for a in self.ai_combo.getCheckedItems() if str(a).strip()]
                    if ai_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/ai_filtered", bool(ai_filtered and ai_checked))
                self._settings.setValue(f"{prefix}/ai_checked", json.dumps(ai_checked))
            if hasattr(self, "status_combo"):
                st_filtered = self.status_combo.isFiltered()
                st_checked = (
                    [str(s) for s in self.status_combo.getCheckedItems() if str(s).strip()]
                    if st_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/status_filtered", bool(st_filtered and st_checked))
                self._settings.setValue(f"{prefix}/status_checked", json.dumps(st_checked))
            self._settings.sync()
        except Exception:
            pass

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
            "QPushButton { padding: 4px 10px; border-radius: 13px; border: 1px solid #CBD5E1; "
            "background-color: #FFFFFF; color: #334155; font-size: 12px; } "
            "QPushButton:hover { background-color: #F1F5F9; } "
            "QPushButton:checked { background-color: #DBEAFE; border-color: #2563EB; "
            "color: #1D4ED8; font-weight: bold; }"
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
            "QPushButton { padding: 4px 9px; border-radius: 5px; border: 1px solid #CBD5E1; "
            "background-color: #F8FAFC; color: #334155; font-size: 12px; } "
            "QPushButton:hover { background-color: #F1F5F9; } "
            "QPushButton:checked { background-color: #EEF2FF; border-color: #6366F1; "
            "color: #4338CA; font-weight: bold; }"
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
        self.status_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
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
            "#InboxMetaFrame { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }"
        )
        meta_layout = QVBoxLayout(meta_frame)
        meta_layout.setContentsMargins(8, 4, 8, 4)
        meta_layout.setSpacing(2)

        self.detail_subject_lbl = QLabel("Select an email to read")
        self.detail_subject_lbl.setWordWrap(True)
        self.detail_subject_lbl.setStyleSheet("font-size: 13px; font-weight: bold; color: #0F172A;")
        meta_layout.addWidget(self.detail_subject_lbl)

        meta_info_row = QHBoxLayout()
        meta_info_row.setContentsMargins(0, 0, 0, 0)
        meta_info_row.setSpacing(8)
        self.detail_meta_lbl = QLabel("")
        self.detail_meta_lbl.setWordWrap(True)
        self.detail_meta_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        meta_info_row.addWidget(self.detail_meta_lbl)

        self.detail_links_lbl = QLabel("")
        self.detail_links_lbl.setWordWrap(True)
        self.detail_links_lbl.setTextFormat(Qt.RichText)
        self.detail_links_lbl.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.detail_links_lbl.setOpenExternalLinks(False)
        self.detail_links_lbl.linkActivated.connect(lambda href: self._on_anchor_clicked(QUrl(href)))
        self.detail_links_lbl.setStyleSheet("color: #1D4ED8; font-size: 11px;")
        meta_info_row.addWidget(self.detail_links_lbl, 1)
        meta_layout.addLayout(meta_info_row)

        # Inline Assignment Bar (Single compact row: AIs, Sub-Topics, WIs + Auto/Pin)
        self._show_assign_bar = bool(self._settings.value("show_assign_bar", True, type=bool))
        self.assign_frame = QFrame(meta_frame)
        self.assign_frame.setObjectName("InboxAssignBar")
        self.assign_frame.setStyleSheet(
            "#InboxAssignBar { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 5px; }"
        )
        assign_row = QHBoxLayout(self.assign_frame)
        assign_row.setContentsMargins(6, 2, 6, 2)
        assign_row.setSpacing(6)

        strict_ai = bool(self._ai_suggestions)
        strict_wi = bool(self._wi_suggestions)

        assign_row.addWidget(QLabel("<b>AI:</b>"))
        self.assign_ais_pills = TagCloudWidget(
            placeholder="AI #...",
            pill_colors={"bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
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
            pill_colors={"bg": "#F0FDFA", "fg": "#0F766E", "border": "#99F6E4"},
            strict_dropdown=True,
            dropdown_label="+ Sub ▾",
            empty_dropdown_label="Assign AI",
        )
        self.assign_subs_pills.tagsChanged.connect(self._on_inline_links_changed)
        assign_row.addWidget(self.assign_subs_pills, 3)

        assign_row.addWidget(QLabel("<b>WI:</b>"))
        self.assign_wis_pills = TagCloudWidget(
            placeholder="WI...",
            pill_colors={"bg": "#EEF2FF", "fg": "#4338CA", "border": "#C7D2FE"},
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

    def _apply_layout_orientation(self, save_current: bool = True):
        if not hasattr(self, "splitter"):
            return
        if save_current:
            prev_key = "splitter_state_v" if self.splitter.orientation() == Qt.Vertical else "splitter_state_h"
            try:
                self._settings.setValue(prev_key, self.splitter.saveState())
            except Exception:
                pass

        is_vertical = self._layout_orientation == "V"
        if is_vertical:
            self.splitter.setOrientation(Qt.Vertical)
            if hasattr(self, "_left_widget"):
                self._left_widget.setMinimumWidth(0)
                self._left_widget.setMinimumHeight(140)
            if hasattr(self, "_right_widget"):
                self._right_widget.setMinimumWidth(0)
                self._right_widget.setMinimumHeight(200)
            self.btn_layout_toggle.setText("◫ Side-by-Side")
            self.btn_layout_toggle.setToolTip("Switch back to Side-by-Side layout (best for wide monitors).")
            saved = self._settings.value("splitter_state_v")
            default_sizes = [260, 440]
        else:
            self.splitter.setOrientation(Qt.Horizontal)
            if hasattr(self, "_left_widget"):
                self._left_widget.setMinimumHeight(0)
                self._left_widget.setMinimumWidth(240)
            if hasattr(self, "_right_widget"):
                self._right_widget.setMinimumHeight(0)
                self._right_widget.setMinimumWidth(280)
            self.btn_layout_toggle.setText("⬒ List on Top")
            self.btn_layout_toggle.setToolTip("Switch to Top/Bottom layout so email subjects span full window width.")
            saved = self._settings.value("splitter_state_h")
            default_sizes = [580, 660]

        restored = False
        if saved:
            try:
                restored = bool(self.splitter.restoreState(saved))
            except Exception:
                restored = False
        if not restored:
            self.splitter.setSizes(default_sizes)
        QTimer.singleShot(0, self._fit_columns_to_viewport)

    def _toggle_layout_orientation(self):
        self._layout_orientation = "H" if self._layout_orientation == "V" else "V"
        self._settings.setValue("layout_orientation", self._layout_orientation)
        self._apply_layout_orientation(save_current=True)

    @staticmethod
    def compute_fitted_column_widths(
        avail_w: int,
        base_widths: List[int],
        mins: Optional[List[int]] = None,
        flex_weights: tuple = (0.03, 0.20, 0.03, 0.48, 0.22, 0.04),
    ) -> List[int]:
        """Proportionally fits interactive column widths to fill `avail_w` without starving Subject or Sender."""
        if avail_w <= 0 or len(base_widths) != 6:
            return list(base_widths)
        min_widths = mins or [
            scale_px(64),
            scale_px(110),
            scale_px(68),
            scale_px(180),
            scale_px(110),
            scale_px(108),
        ]
        clamped = [max(m, int(w)) for w, m in zip(base_widths, min_widths)]
        total_base = sum(clamped)
        if avail_w >= total_base:
            extra = avail_w - total_base
            out = [int(w + round(extra * wt)) for w, wt in zip(clamped, flex_weights)]
        else:
            min_sum = sum(min_widths)
            if avail_w <= min_sum:
                return list(min_widths)
            shrinkable = [max(0, w - m) for w, m in zip(clamped, min_widths)]
            total_shrink = sum(shrinkable)
            if total_shrink <= 0:
                return list(min_widths)
            need = total_base - avail_w
            out = [
                max(m, int(round(w - (need * sk / total_shrink))))
                for w, m, sk in zip(clamped, min_widths, shrinkable)
            ]
        diff = avail_w - sum(out)
        if diff != 0:
            out[3] = max(min_widths[3], out[3] + diff)
        return out

    def _fit_columns_to_viewport(self, avail_width: Optional[int] = None):
        if not hasattr(self, "table"):
            return
        if avail_width is None:
            if not self.isVisible():
                return
            avail_w = int(self.table.viewport().width())
        else:
            avail_w = int(avail_width)
        if avail_w < 300:
            return
        hdr = self.table.horizontalHeader()
        base = list(getattr(self, "_col_base_widths", []) or [hdr.sectionSize(c) for c in range(6)])
        fitted = self.compute_fitted_column_widths(avail_w, base)
        self._restoring_header = True
        try:
            for col_i, col_w in enumerate(fitted):
                hdr.resizeSection(col_i, col_w)
        finally:
            self._restoring_header = False

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._fit_columns_to_viewport)

    def _on_header_section_resized(self, *_args):
        if getattr(self, "_restoring_header", False) or getattr(self, "_restoring_filters", False):
            return
        try:
            hdr = self.table.horizontalHeader()
            self._col_base_widths = [max(40, hdr.sectionSize(c)) for c in range(6)]
            prefix = self._meeting_settings_prefix()
            self._settings.setValue(f"{prefix}/table_header_state", hdr.saveState())
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            self._flush_pending_note_save(reload_view=False)
            if getattr(self, "_sync_worker", None) and self._sync_worker.isRunning():
                self._sync_worker.request_cancel()
                self._sync_worker.wait(1500)
            self._settings.setValue("geometry", self.saveGeometry())
            self._settings.setValue("layout_orientation", self._layout_orientation)
            if hasattr(self, "splitter"):
                key = "splitter_state_v" if self.splitter.orientation() == Qt.Vertical else "splitter_state_h"
                self._settings.setValue(key, self.splitter.saveState())
            if hasattr(self, "table"):
                prefix = self._meeting_settings_prefix()
                self._settings.setValue(f"{prefix}/table_header_state", self.table.horizontalHeader().saveState())
            if hasattr(self, "preview_browser"):
                self._settings.setValue("zoom_steps", int(self.preview_browser.property("_reader_zoom_steps") or 0))
            self._save_filter_settings()
        except Exception:
            pass
        super().closeEvent(event)

    def _set_mode(self, mode: str):
        norm = str(mode or "UNREAD").strip().upper()
        if norm not in ("UNREAD", "FLAGGED", "RECENT", "ALL"):
            norm = "UNREAD"
        self._active_mode = norm
        for pill, key in (
            (getattr(self, "pill_unread", None), "UNREAD"),
            (getattr(self, "pill_flagged", None), "FLAGGED"),
            (getattr(self, "pill_recent", None), "RECENT"),
            (getattr(self, "pill_all", None), "ALL"),
        ):
            if pill is not None:
                pill.blockSignals(True)
                pill.setChecked(norm == key)
                pill.blockSignals(False)
        self._populate_filter_combos()
        self._apply_filters()

    def _set_view_mode(self, view_mode: str):
        norm = "THREADS" if str(view_mode or "").upper() == "THREADS" else "MESSAGES"
        self._view_mode = norm
        if hasattr(self, "btn_view_messages"):
            self.btn_view_messages.blockSignals(True)
            self.btn_view_messages.setChecked(norm == "MESSAGES")
            self.btn_view_messages.blockSignals(False)
        if hasattr(self, "btn_view_threads"):
            self.btn_view_threads.blockSignals(True)
            self.btn_view_threads.setChecked(norm == "THREADS")
            self.btn_view_threads.blockSignals(False)
        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters()

    def _active_dataset(self) -> List[dict]:
        return self._grouped_items if self._view_mode == "THREADS" else self._all_items

    def _build_grouped_thread_items(self) -> List[dict]:
        buckets: Dict[str, List[dict]] = {}
        for it in self._all_items:
            if it.get("in_tracked_db") and it.get("thread_id"):
                key = f"THR::{it['thread_id']}"
            else:
                norm = normalize_subject(it.get("subject") or "").lower()
                key = f"UNPIN::{norm or it['id']}"
            buckets.setdefault(key, []).append(it)

        grouped: List[dict] = []
        for gkey, msgs in buckets.items():
            msgs_asc = sorted(msgs, key=lambda m: str(m.get("date_str") or ""))
            latest = msgs_asc[-1]
            in_tracked = any(m.get("in_tracked_db") for m in msgs_asc)
            tid = next((m.get("thread_id") for m in msgs_asc if m.get("thread_id")), "")
            th_title = next((m.get("thread_title") for m in msgs_asc if m.get("thread_title")), "")
            linked_ais = list(dict.fromkeys(ai for m in msgs_asc for ai in (m.get("linked_ais") or []) if ai))
            linked_subs = list(dict.fromkeys(s for m in msgs_asc for s in (m.get("linked_sub_items") or []) if s))
            linked_wis = list(dict.fromkeys(w for m in msgs_asc for w in (m.get("linked_wis") or []) if w))
            linked_tdocs = list(dict.fromkeys(td for m in msgs_asc for td in (m.get("linked_tdocs") or []) if td))
            companies = list(dict.fromkeys(m.get("company") for m in msgs_asc if m.get("company")))
            unread_cnt = sum(1 for m in msgs_asc if not m["is_read"])
            flagged_cnt = sum(1 for m in msgs_asc if m["is_flagged"])
            user_status = next(
                (str(m.get("user_status") or "").strip() for m in msgs_asc if str(m.get("user_status") or "").strip() and m.get("user_status") != "⚪ Neutral"),
                str(msgs_asc[0].get("user_status") or "⚪ Neutral").strip() if msgs_asc else "⚪ Neutral",
            )
            user_notes = next((str(m.get("user_notes") or "").strip() for m in msgs_asc if str(m.get("user_notes") or "").strip()), "")

            grouped.append({
                "id": gkey,
                "_is_group": True,
                "_group_messages": msgs_asc,
                "_unread_count": unread_cnt,
                "_total_count": len(msgs_asc),
                "_flagged_count": flagged_cnt,
                "_companies": companies,
                "source_kind": "thread" if in_tracked else "tdoc",
                "thread_id": tid,
                "thread_title": th_title or normalize_subject(latest.get("subject") or ""),
                "subject": th_title or latest.get("subject") or "(No Subject)",
                "sender_name": latest.get("sender_name") or "",
                "sender_email": latest.get("sender_email") or "",
                "company": ", ".join(companies) if companies else (latest.get("company") or ""),
                "date_str": latest.get("date_str") or "",
                "body_text": latest.get("body_text") or "",
                "is_read": unread_cnt == 0,
                "is_flagged": flagged_cnt > 0,
                "attachments": latest.get("attachments") or [],
                "linked_ais": linked_ais,
                "linked_sub_items": linked_subs,
                "linked_wis": linked_wis,
                "linked_tdocs": linked_tdocs,
                "in_tracked_db": in_tracked,
                "in_general_db": any(m.get("in_general_db") for m in msgs_asc),
                "conversation_id": next((m.get("conversation_id") for m in msgs_asc if m.get("conversation_id")), ""),
                "user_status": user_status or "⚪ Neutral",
                "user_notes": user_notes,
            })
        return sorted(grouped, key=lambda x: x.get("date_str", ""), reverse=True)

    def reload_from_dbs(
        self,
        preserve_id: Optional[str] = None,
        preserve_ids: Optional[List[str]] = None,
    ):
        """Loads and deduplicates emails from TrackedThreadsDatabase and GeneralEmailDatabase."""
        if not getattr(self, "_restoring_filters", False) and hasattr(self, "note_edit"):
            self._flush_pending_note_save(reload_view=False)
        by_id: Dict[str, dict] = {}

        if self.tracked_db is not None:
            threads = self.tracked_db.get_threads_for_meeting(self.meeting_id)
            for th in threads:
                tid = str(th.get("thread_id") or "")
                th_title = str(th.get("title") or "").strip()
                th_ais = [str(a).strip() for a in (th.get("linked_ais") or []) if str(a).strip()]
                th_subs = [str(s).strip() for s in (th.get("linked_sub_items") or []) if str(s).strip()]
                th_wis = [str(w).strip() for w in (th.get("linked_wis") or []) if str(w).strip()]
                th_tdocs = list(dict.fromkeys(
                    self._clean_tdoc_id(t) for t in (th.get("linked_tdocs") or []) if self._clean_tdoc_id(t)
                ))
                th_status = str(th.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                th_notes = str(th.get("user_notes") or "").strip()
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
                        "clean_body": str(m.get("clean_body") or ""),
                        "raw_body_text": str(m.get("body_text") or ""),
                        "html_body": "",
                        "is_read": int(m.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(m.get("flag_status", 0) or 0) == 2,
                        "attachments": list(m.get("attachments") or []),
                        "linked_ais": list(th_ais),
                        "linked_sub_items": list(th_subs),
                        "linked_wis": list(th_wis),
                        "linked_tdocs": list(th_tdocs),
                        "conversation_id": str(m.get("conversation_id") or ""),
                        "in_tracked_db": True,
                        "in_general_db": False,
                        "user_status": th_status,
                        "user_notes": th_notes,
                    }

        if self.general_email_db is not None and hasattr(self.general_email_db, "get_all_emails_with_tdocs"):
            for row in self.general_email_db.get_all_emails_with_tdocs(show_ignored=False):
                eid = str(row.get("id") or "").strip()
                if not eid:
                    continue
                matched_tdocs = list(dict.fromkeys(
                    self._clean_tdoc_id(t)
                    for t in str(row.get("matched_tdocs") or "").split(",")
                    if self._clean_tdoc_id(t)
                ))
                inferred_ais = []
                inferred_subs = []
                inferred_wis = []
                for td in matched_tdocs:
                    ai = self.tdoc_ai_lookup.get(td, "")
                    meta = self._tdoc_meta_map.get(td) or {}
                    if not ai:
                        ai = str(meta.get("ai") or "").strip()
                    if ai and ai not in inferred_ais:
                        inferred_ais.append(ai)
                    sub = str(meta.get("sub_item") or "").strip()
                    if sub and sub not in inferred_subs:
                        inferred_subs.append(sub)
                    for w in (meta.get("wis") or []):
                        if w and w not in inferred_wis:
                            inferred_wis.append(w)

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
                        "body_text": str(row.get("clean_body") or row.get("body_text") or ""),
                        "clean_body": str(row.get("clean_body") or ""),
                        "raw_body_text": str(row.get("body_text") or ""),
                        "html_body": str(row.get("html_body") or ""),
                        "is_read": int(row.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(row.get("flag_status", 0) or 0) == 2,
                        "attachments": list(row.get("attachments") or []),
                        "linked_ais": inferred_ais,
                        "linked_sub_items": inferred_subs,
                        "linked_wis": inferred_wis,
                        "linked_tdocs": matched_tdocs,
                        "conversation_id": "",
                        "in_tracked_db": False,
                        "in_general_db": True,
                        "user_status": "⚪ Neutral",
                        "user_notes": "",
                    }

        self._all_items = sorted(by_id.values(), key=lambda x: x.get("date_str", ""), reverse=True)
        self._grouped_items = self._build_grouped_thread_items()
        self._items_by_id = {it["id"]: it for it in self._all_items}

        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters(preserve_id=preserve_id, preserve_ids=preserve_ids)

    @staticmethod
    def _clean_tdoc_id(raw: Any) -> str:
        return clean_tdoc_id(raw)

    def _item_matches_mode(self, it: dict, mode: str, cutoff: str) -> bool:
        if mode == "UNREAD" and it["is_read"]:
            return False
        if mode == "FLAGGED" and not it["is_flagged"]:
            return False
        if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
            return False
        return True

    def _populate_filter_combos(self):
        cutoff = self._recent_cutoff_str()
        dataset = [it for it in self._active_dataset() if self._item_matches_mode(it, self._active_mode, cutoff)]

        companies = sorted(
            {
                c
                for it in dataset
                for c in (it.get("_companies") if it.get("_is_group") else [it.get("company")])
                if c
            },
            key=str.lower,
        )
        ais = sorted({ai for it in dataset for ai in (it.get("linked_ais") or []) if ai})
        unassigned_cnt = sum(1 for it in dataset if not (it.get("linked_ais") or []))

        self.company_combo.blockSignals(True)
        self.company_combo.updateItems(companies)
        if self._pending_company_filter is not None and companies:
            wanted_lower = {c.lower() for c in self._pending_company_filter}
            matched_comps = [c for c in companies if c.lower() in wanted_lower]
            if matched_comps:
                self.company_combo.setCheckedItems(matched_comps)
            self._pending_company_filter = None
        self.company_combo.blockSignals(False)

        ai_items = ["__UNASSIGNED__"] + ais
        ai_labels: Dict[str, str] = {
            "__UNASSIGNED__": f"❓ Unassigned (No AI) ({unassigned_cnt})",
        }
        ai_tooltips: Dict[str, str] = {
            "__UNASSIGNED__": f"Emails or threads without any linked AI ({unassigned_cnt})",
        }
        for ai in ais:
            full_label = self._ai_label_map.get(ai, ai)
            ai_tooltips[ai] = full_label
            ai_labels[ai] = (full_label[:45] + "...") if len(full_label) > 48 else full_label

        self.ai_combo.blockSignals(True)
        self.ai_combo.updateItems(ai_items, display_labels=ai_labels, item_tooltips=ai_tooltips)
        if self._pending_ai_filter is not None and dataset:
            wanted_ais = set(self._pending_ai_filter)
            matched_ais = [a for a in ai_items if a in wanted_ais]
            if matched_ais:
                self.ai_combo.setCheckedItems(matched_ais)
            self._pending_ai_filter = None
        self.ai_combo.blockSignals(False)

        if hasattr(self, "status_combo"):
            status_counts: Dict[str, int] = {s: 0 for s in MY_STATUS_OPTIONS}
            for it in dataset:
                st = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                status_counts[st] = status_counts.get(st, 0) + 1
            st_items = list(MY_STATUS_OPTIONS)
            st_labels = {s: f"{s} ({status_counts.get(s, 0)})" for s in st_items}
            self.status_combo.blockSignals(True)
            self.status_combo.updateItems(st_items, display_labels=st_labels)
            if self._pending_status_filter is not None and dataset:
                wanted_st = set(self._pending_status_filter)
                matched_st = [s for s in st_items if s in wanted_st]
                if matched_st:
                    self.status_combo.setCheckedItems(matched_st)
                self._pending_status_filter = None
            self.status_combo.blockSignals(False)

    @staticmethod
    def _recent_cutoff_str() -> str:
        return (datetime.now() - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M")

    def _update_pill_counts(self):
        cutoff = self._recent_cutoff_str()
        dataset = self._active_dataset()
        unread_cnt = sum(1 for it in dataset if not it["is_read"])
        flagged_cnt = sum(1 for it in dataset if it["is_flagged"])
        recent_cnt = sum(1 for it in dataset if str(it.get("date_str") or "") >= cutoff)
        total_cnt = len(dataset)

        self.pill_unread.setText(f"🔵 Unread Only ({unread_cnt})")
        self.pill_flagged.setText(f"🚩 Flagged ({flagged_cnt})")
        self.pill_recent.setText(f"🕒 Last 24h ({recent_cnt})")
        self.pill_all.setText(f"📋 All ({total_cnt})")

    def _apply_filters(
        self,
        *_args,
        preserve_id: Optional[str] = None,
        preserve_ids: Optional[List[str]] = None,
        select_row: Optional[int] = None,
    ):
        self._save_filter_settings()
        target_id = preserve_id if preserve_id is not None else (self._current_item["id"] if self._current_item else None)
        mode = self._active_mode
        cutoff = self._recent_cutoff_str()
        q = self.search_edit.text().strip().lower()
        filter_comp = self.company_combo.isFiltered()
        selected_comps = (
            {str(c).strip().lower() for c in self.company_combo.getCheckedItems() if str(c).strip()}
            if filter_comp
            else set()
        )
        filter_ai = self.ai_combo.isFiltered()
        selected_ais = (
            {str(a).strip() for a in self.ai_combo.getCheckedItems() if str(a).strip()}
            if filter_ai
            else set()
        )
        filter_status = hasattr(self, "status_combo") and self.status_combo.isFiltered()
        selected_statuses = (
            {str(s).strip() for s in self.status_combo.getCheckedItems() if str(s).strip()}
            if filter_status
            else set()
        )
        source_filter = (self.source_combo.currentData() or "").strip()

        dataset = self._active_dataset()
        visible: List[dict] = []
        for it in dataset:
            if mode == "UNREAD" and it["is_read"]:
                continue
            if mode == "FLAGGED" and not it["is_flagged"]:
                continue
            if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
                continue
            if source_filter == "THREAD" and not it["in_tracked_db"]:
                continue
            if source_filter == "UNPINNED" and it["in_tracked_db"]:
                continue
            if source_filter == "NEEDS_TRIAGE" and it["in_tracked_db"] and bool(it.get("linked_ais")):
                continue
            if filter_comp:
                item_comps = (
                    {str(c).strip().lower() for c in (it.get("_companies") or []) if str(c).strip()}
                    if it.get("_is_group")
                    else {str(it.get("company") or "").strip().lower()}
                )
                if not (item_comps & selected_comps):
                    continue
            if filter_ai:
                item_ais = {str(a).strip() for a in (it.get("linked_ais") or []) if str(a).strip()}
                matches_unassigned = ("__UNASSIGNED__" in selected_ais) and (not item_ais)
                matches_ai = bool(item_ais & (selected_ais - {"__UNASSIGNED__"}))
                if not (matches_unassigned or matches_ai):
                    continue
            if filter_status:
                item_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                if item_status not in selected_statuses:
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
            u_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
            u_notes = str(it.get("user_notes") or "").strip()
            note_badge = "📝" if u_notes else ""
            if it.get("_is_group"):
                u_cnt = int(it.get("_unread_count", 0))
                t_cnt = int(it.get("_total_count", 1))
                f_part = "🚩" if it["is_flagged"] else ""
                core_s = f"🔵 {u_cnt}/{t_cnt} {f_part}".strip() if u_cnt > 0 else f"✓ {t_cnt} {f_part}".strip()
                state_txt = f"{core_s} {note_badge}".strip()
            else:
                state_parts = []
                if not it["is_read"]:
                    state_parts.append("🔵")
                if it["is_flagged"]:
                    state_parts.append("🚩")
                if not state_parts:
                    state_parts.append("✓")
                if note_badge:
                    state_parts.append(note_badge)
                state_txt = " ".join(state_parts)
            state_item = QTableWidgetItem(state_txt)
            state_item.setData(Qt.UserRole, it["id"])
            state_item.setData(Qt.UserRole + 1, _status_color_hex(u_status))
            state_item.setTextAlignment(Qt.AlignCenter)
            st_tip = [f"Status: {u_status}"] if u_status != "⚪ Neutral" else []
            if u_notes:
                preview_n = (u_notes[:180] + "…") if len(u_notes) > 180 else u_notes
                st_tip.append(f"Comment: {preview_n}")
            if st_tip:
                state_item.setToolTip("\n".join(st_tip))

            if it["in_tracked_db"] and it.get("thread_title"):
                ctx_full = f"🧵 {it['thread_title']}"
                ctx = f"🧵 {it['thread_title'][:34]}"
                if len(it["thread_title"]) > 34:
                    ctx += "…"
            elif it.get("linked_tdocs"):
                ctx_full = "📄 " + ", ".join(it["linked_tdocs"])
                ctx = "📄 " + ", ".join(it["linked_tdocs"][:2])
            else:
                ctx_full = "📧 Email"
                ctx = "📧 Email"
            ctx_item = QTableWidgetItem(ctx)
            ctx_item.setToolTip(ctx_full)

            ais = it.get("linked_ais") or []
            ai_item = QTableWidgetItem(", ".join(ais) if ais else "❓ —")
            if not ais:
                ai_item.setForeground(QColor("#B45309"))
                ai_item.setToolTip("Unassigned (No AI)")
            else:
                ai_item.setToolTip("\n".join(self._ai_label_map.get(a, a) for a in ais))

            subj_full = str(it.get("subject") or "")
            subj_item = QTableWidgetItem(subj_full)
            subj_item.setToolTip(subj_full)
            if not it["is_read"]:
                f = subj_item.font()
                f.setBold(True)
                subj_item.setFont(f)

            sender_disp = format_sender_display(
                it["sender_name"] or it["sender_email"] or "Unknown",
                it.get("company") or "",
            )
            sender_item = QTableWidgetItem(sender_disp)
            sender_item.setToolTip(
                f"{sender_disp} <{it['sender_email']}>" if it.get("sender_email") else sender_disp
            )
            date_full = str(it.get("date_str") or "")
            date_item = QTableWidgetItem(date_full[:16])
            date_item.setToolTip(date_full)

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
            rows_to_select: List[int] = []
            if preserve_ids:
                pid_set = set(preserve_ids)
                rows_to_select = [r for r, it in enumerate(visible) if it["id"] in pid_set]
            if not rows_to_select:
                if select_row is not None and 0 <= select_row < len(visible):
                    rows_to_select = [select_row]
                elif select_row is not None and select_row >= len(visible):
                    rows_to_select = [len(visible) - 1]
                elif row_to_select >= 0:
                    rows_to_select = [row_to_select]
                else:
                    rows_to_select = [0]

            sel_model = self.table.selectionModel()
            if sel_model is None:
                self.table.selectRow(rows_to_select[0])
            else:
                sel_model.clearSelection()
                for r in rows_to_select:
                    sel_model.select(
                        self.table.model().index(r, 0),
                        QItemSelectionModel.Select | QItemSelectionModel.Rows,
                    )
                sel_model.setCurrentIndex(
                    self.table.model().index(rows_to_select[0], 0),
                    QItemSelectionModel.NoUpdate,
                )
        self.table.blockSignals(False)

        noun = "thread(s)" if self._view_mode == "THREADS" else "indexed email(s)"
        self.status_lbl.setText(f"Showing {len(visible)} of {len(dataset)} {noun}")
        self._update_sync_button_state(
            total_count=len(dataset),
            visible_count=len(visible),
            has_narrow_filter=bool(filter_comp or filter_ai or filter_status or q or source_filter or mode == "FLAGGED"),
        )
        self._on_selection_changed()

    def _on_selection_changed(self):
        self._flush_pending_note_save(reload_view=False)
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}) if self.table.selectionModel() else []
        selected = [self._visible_items[r] for r in rows if 0 <= r < len(self._visible_items)]
        self._selected_items = selected
        if not selected:
            self._current_item = None
            self._render_empty_detail()
            return
        self._current_item = selected[0]
        if len(selected) == 1:
            self._render_detail(selected[0])
        else:
            self._render_multi_detail(selected)

    def _render_empty_detail(self):
        self.detail_subject_lbl.setText("No matching emails in this filter")
        self.detail_meta_lbl.setText("")
        self.detail_links_lbl.setText("")
        self.assign_frame.setVisible(False)
        self.note_frame.setVisible(False)
        self.preview_browser.setMarkdown("*Select another filter tab or adjust search criteria.*")
        for btn in (self.btn_mark_next, self.btn_toggle_read, self.btn_toggle_flag, self.btn_reply, self.btn_open_outlook, self.btn_open_context, self.btn_sync_selected, self.btn_toggle_note, self.btn_merge_selected):
            btn.setEnabled(False)
        self.btn_merge_selected.setVisible(False)
        if hasattr(self, "detail_status_combo"):
            self.detail_status_combo.setEnabled(False)
        self._update_selected_sync_controls(0)

    def _on_toggle_assign_bar(self, checked: bool):
        self._show_assign_bar = bool(checked)
        self.assign_frame.setVisible(self._show_assign_bar)
        self.btn_toggle_assign_bar.setText("🏷️ AI / WI ▴" if self._show_assign_bar else "🏷️ AI / WI ▾")
        try:
            self._settings.setValue("show_assign_bar", self._show_assign_bar)
            self._settings.sync()
        except Exception:
            pass

    def _refresh_assign_sub_suggestions(self, selected_ais: List[str]):
        allowed: List[str] = []
        seen = set()
        for ai in selected_ais:
            for s in self._ai_to_sub_items.get(ai, []):
                if s not in seen:
                    seen.add(s)
                    allowed.append(s)
        self.assign_subs_pills.set_suggestions(allowed)
        if not selected_ais:
            self.assign_subs_pills.set_empty_dropdown_label("Assign AI")
        elif not allowed:
            self.assign_subs_pills.set_empty_dropdown_label("No sub-topics")

    def _populate_assign_bar(self, it: dict):
        self._updating_assign_bar = True
        try:
            self.assign_frame.setVisible(self._show_assign_bar)
            ais = list(it.get("linked_ais") or [])
            subs = list(it.get("linked_sub_items") or [])
            wis = list(it.get("linked_wis") or [])
            self._multi_base_common = {"ais": [], "subs": [], "wis": []}
            self._refresh_assign_sub_suggestions(ais)
            self.assign_ais_pills.setText(", ".join(ais))
            self.assign_subs_pills.setText(", ".join(subs))
            self.assign_wis_pills.setText(", ".join(wis))
            self.btn_pin_inline.setVisible(not it.get("in_tracked_db") and self.tracked_db is not None)
            self.btn_pin_inline.setEnabled(not it.get("in_tracked_db") and self.tracked_db is not None)
            self.btn_auto_assign.setEnabled(self.tracked_db is not None)
        finally:
            self._updating_assign_bar = False

    def _populate_assign_bar_multi(self, selected: List[dict]):
        self._updating_assign_bar = True
        try:
            self.assign_frame.setVisible(self._show_assign_bar)
            common_ais = [
                a for a in (selected[0].get("linked_ais") or [])
                if all(a in (it.get("linked_ais") or []) for it in selected[1:])
            ]
            common_subs = [
                s for s in (selected[0].get("linked_sub_items") or [])
                if all(s in (it.get("linked_sub_items") or []) for it in selected[1:])
            ]
            common_wis = [
                w for w in (selected[0].get("linked_wis") or [])
                if all(w in (it.get("linked_wis") or []) for it in selected[1:])
            ]
            union_ais = list(dict.fromkeys(
                a for it in selected for a in (it.get("linked_ais") or []) if a
            ))
            self._multi_base_common = {
                "ais": list(common_ais),
                "subs": list(common_subs),
                "wis": list(common_wis),
            }
            self._refresh_assign_sub_suggestions(union_ais)
            self.assign_ais_pills.setText(", ".join(common_ais))
            self.assign_subs_pills.setText(", ".join(common_subs))
            self.assign_wis_pills.setText(", ".join(common_wis))
            has_unpinned = any(not it.get("in_tracked_db") for it in selected)
            self.btn_pin_inline.setVisible(has_unpinned and self.tracked_db is not None)
            self.btn_pin_inline.setEnabled(has_unpinned and self.tracked_db is not None)
            self.btn_auto_assign.setEnabled(self.tracked_db is not None)
        finally:
            self._updating_assign_bar = False

    def _render_multi_detail(self, selected: List[dict]):
        noun = "Threads" if self._view_mode == "THREADS" else "Emails"
        unread_cnt = sum(1 for it in selected if not it["is_read"])
        flagged_cnt = sum(1 for it in selected if it["is_flagged"])
        unpinned_cnt = sum(1 for it in selected if not it.get("in_tracked_db"))
        pinned_cnt = len(selected) - unpinned_cnt

        self.detail_subject_lbl.setText(f"📑 {len(selected)} Selected {noun}")
        self.detail_meta_lbl.setText(
            f"{unread_cnt} unread  |  {flagged_cnt} flagged  |  "
            f"{pinned_cnt} tracked, {unpinned_cnt} unpinned"
        )

        cleaned_tdocs = list(dict.fromkeys(
            self._clean_tdoc_id(td)
            for it in selected
            for td in (it.get("linked_tdocs") or [])
            if self._clean_tdoc_id(td)
        ))
        links_html = [f"Batch selection ({len(selected)} items) — adding AIs/Sub-Topics/WIs applies to all selected items without removing each item's other existing tags"]
        if cleaned_tdocs:
            if self._open_tdoc_callback:
                td_anchors = ", ".join(
                    f'<a href="tdoc:{html.escape(td)}" '
                    f'style="color: #1D4ED8; font-weight: 600; text-decoration: underline;">'
                    f'{html.escape(td)}</a>'
                    for td in cleaned_tdocs[:10]
                )
            else:
                td_anchors = ", ".join(html.escape(td) for td in cleaned_tdocs[:10])
            if len(cleaned_tdocs) > 10:
                td_anchors += f" (+{len(cleaned_tdocs) - 10} more)"
            links_html.append(f"TDocs: {td_anchors}")
        self.detail_links_lbl.setText(" &nbsp;|&nbsp; ".join(links_html))
        self.detail_links_lbl.setVisible(True)

        self._populate_assign_bar_multi(selected)
        self._populate_status_and_note_bars(selected)

        self.btn_mark_next.setEnabled(False)
        self.btn_reply.setEnabled(False)
        self.btn_reply.setVisible(False)
        self.btn_open_outlook.setEnabled(False)
        self.btn_open_outlook.setVisible(False)
        self.btn_open_context.setEnabled(False)
        self.btn_open_context.setVisible(False)
        self.btn_toggle_read.setEnabled(True)
        self.btn_toggle_read.setText(
            f"✓ Read ({unread_cnt})" if unread_cnt > 0 else f"🔵 Unread ({len(selected)})"
        )
        self.btn_toggle_flag.setEnabled(True)
        unflagged_cnt = len(selected) - flagged_cnt
        self.btn_toggle_flag.setText(
            f"🚩 ({unflagged_cnt})" if unflagged_cnt > 0 else f"🏳️ ({len(selected)})"
        )
        self._update_selected_sync_controls(len(selected))
        merge_targets = self._distinct_mergeable_targets(selected)
        can_merge = bool(self.tracked_db is not None and len(merge_targets) >= 2)
        self.btn_merge_selected.setVisible(can_merge)
        self.btn_merge_selected.setEnabled(can_merge)
        if can_merge:
            self.btn_merge_selected.setText(f"🔗 Merge ({len(merge_targets)})")

        md_lines = [f"### Batch Triage ({len(selected)} selected {noun.lower()})", ""]
        md_lines.append("Use the **AIs**, **Sub-Topics**, **WIs**, **Auto-Assign**, **Pin**, **Merge**, **Mark Read**, or **Flag** controls above to update all selected items together.\n")
        for idx, it in enumerate(selected, 1):
            subj = str(it.get("subject") or "(No Subject)")
            comp = str(it.get("company") or "Unknown")
            ais = ", ".join(it.get("linked_ais") or []) or "Unassigned"
            wis = ", ".join(it.get("linked_wis") or []) or "None"
            status_tag = "🧵 Tracked" if it.get("in_tracked_db") else "📄 Unpinned"
            md_lines.append(f"{idx}. **{subj}** — *{comp}* · {status_tag} · AI: `{ais}` · WI: `{wis}`")
        self.preview_browser.setMarkdown("\n".join(md_lines))

    def _render_detail(self, it: dict):
        subject_text = str(it.get("subject") or "(No Subject)")
        self.detail_subject_lbl.setText(subject_text)
        if it.get("_is_group"):
            raw_comps = [str(c).strip() for c in (it.get("_companies") or []) if str(c).strip()]
            comp_disp = ", ".join(raw_comps[:5]) + (f" (+{len(raw_comps) - 5})" if len(raw_comps) > 5 else "")
            if not comp_disp:
                comp_disp = it.get("company") or "Unknown"
            self.detail_meta_lbl.setText(
                f"{it.get('_total_count', 1)} msgs ({it.get('_unread_count', 0)} unread) · "
                f"{format_compact_timestamp(it.get('date_str'))} · {comp_disp}"
            )
            self.detail_meta_lbl.setToolTip(", ".join(raw_comps) if raw_comps else comp_disp)
        else:
            sender_disp = format_sender_display(
                it.get("sender_name") or it.get("sender_email") or "Unknown",
                it.get("company") or "",
            )
            badges = []
            if not it["is_read"]:
                badges.append("🔵 NEW")
            if it["is_flagged"]:
                badges.append("🚩")
            badge_str = f" · {' '.join(badges)}" if badges else " · ✓"
            self.detail_meta_lbl.setText(
                f"{sender_disp} · {format_compact_timestamp(it.get('date_str'))}{badge_str}"
            )
            self.detail_meta_lbl.setToolTip("")

        links_html = []
        if it.get("in_tracked_db") and it.get("thread_title"):
            tt = str(it["thread_title"]).strip()
            if tt.lower() == subject_text.strip().lower():
                links_html.append("🧵 Tracked")
            else:
                links_html.append(f"🧵 {html.escape(tt)}")
        else:
            links_html.append("📄 Unpinned")
        cleaned_tdocs = list(dict.fromkeys(
            self._clean_tdoc_id(td) for td in (it.get("linked_tdocs") or []) if self._clean_tdoc_id(td)
        ))
        if cleaned_tdocs:
            if self._open_tdoc_callback:
                td_anchors = ", ".join(
                    f'<a href="tdoc:{html.escape(td)}" '
                    f'style="color: #1D4ED8; font-weight: 600; text-decoration: underline;">'
                    f'{html.escape(td)}</a>'
                    for td in cleaned_tdocs[:6]
                )
            else:
                td_anchors = ", ".join(html.escape(td) for td in cleaned_tdocs[:6])
            if len(cleaned_tdocs) > 6:
                td_anchors += f" (+{len(cleaned_tdocs) - 6})"
            links_html.append(f"TDocs: {td_anchors}")
        self.detail_links_lbl.setText(" · ".join(links_html))
        self.detail_links_lbl.setVisible(bool(links_html))

        self._populate_assign_bar(it)
        self._populate_status_and_note_bars([it])

        can_outlook = not str(it.get("id") or "").startswith("MSG-") and not it.get("_is_group")
        has_context = bool(
            (it.get("thread_id") and self._open_thread_callback)
            or (it.get("linked_tdocs") and self._open_tdoc_callback)
        )
        self.btn_mark_next.setEnabled(True)
        self.btn_toggle_read.setEnabled(True)
        self.btn_toggle_read.setText("✓ Read" if not it["is_read"] else "🔵 Unread")
        self.btn_toggle_flag.setEnabled(True)
        self.btn_toggle_flag.setText("🏳️" if it["is_flagged"] else "🚩")
        self.btn_reply.setEnabled(can_outlook)
        self.btn_reply.setVisible(can_outlook)
        self.btn_open_outlook.setEnabled(can_outlook)
        self.btn_open_outlook.setVisible(can_outlook)
        self.btn_open_context.setEnabled(has_context)
        self.btn_open_context.setVisible(has_context)
        if it.get("thread_id") and self._open_thread_callback:
            self.btn_open_context.setText("📂 Thread")
        elif it.get("linked_tdocs") and self._open_tdoc_callback:
            self.btn_open_context.setText(f"📄 {it['linked_tdocs'][0]}")
        else:
            self.btn_open_context.setText("📂 Context")
        self.btn_merge_selected.setVisible(False)
        self.btn_merge_selected.setEnabled(False)
        self._update_selected_sync_controls(1)

        if it.get("_is_group"):
            synth = {
                "thread_id": it.get("thread_id") or it["id"],
                "title": it.get("thread_title") or it.get("subject") or "Thread",
                "normalized_subject": normalize_subject(it.get("subject") or "").lower(),
                "linked_tdocs": it.get("linked_tdocs") or [],
            }
            group_msgs = [
                {
                    "id": m["id"],
                    "sender_name": m.get("sender_name", ""),
                    "sender_email": m.get("sender_email", ""),
                    "company": m.get("company", ""),
                    "date_sent": m.get("date_str", ""),
                    "subject": m.get("subject", ""),
                    "body_text": m.get("raw_body_text") or m.get("body_text", ""),
                    "clean_body": m.get("clean_body") or m.get("body_text", ""),
                    "is_read": 1 if m["is_read"] else 0,
                    "flag_status": 2 if m["is_flagged"] else 0,
                    "attachments": m.get("attachments") or [],
                }
                for m in (it.get("_group_messages") or [])
            ]
            md = export_thread_to_markdown(
                synth,
                group_msgs,
                deduplicate_quotes=True,
                interactive_links=True,
                newest_first=True,
                include_metadata=False,
                enable_tdoc_links=bool(self._open_tdoc_callback),
                include_discussion_heading=False,
            )
            self.preview_browser.setMarkdown(md)
            blocks = find_unread_heading_blocks(self.preview_browser)
            if blocks:
                QTimer.singleShot(0, lambda b=blocks[0]: scroll_browser_to_block(self.preview_browser, b))
            return

        lines: List[str] = []
        raw_atts = it.get("attachments") or []
        real_atts, inline_imgs = partition_email_attachments(raw_atts)
        if real_atts and can_outlook:
            enc_id = urllib.parse.quote(it["id"], safe="")
            att_links = [
                f"[📎 {a}](outlook://att/{enc_id}/{urllib.parse.quote(str(a), safe='')})"
                for a in real_atts
            ]
            att_line = f"**Attachments:** {' · '.join(att_links)}"
            if inline_imgs:
                att_line += f" *(+{len(inline_imgs)} inline)*"
            lines.append(att_line + "\n")

        body = str(it.get("body_text") or "")
        if it.get("html_body"):
            body = inject_html_tables_into_body(body, str(it["html_body"]))
        body = _format_trailing_quoted_history(body.strip())
        body = linkify_markdown_urls(body, shorten_long_urls=True)
        if self._open_tdoc_callback:
            body = linkify_tdoc_tokens(body)
        body = escape_non_html_angles(body)
        lines.append(body if body else "*(Empty message body)*")
        self.preview_browser.setMarkdown("\n\n".join(lines))

    def _seed_messages_for_item(self, it: dict) -> List[dict]:
        raw_msgs = it.get("_group_messages") if it.get("_is_group") else [it]
        out: List[dict] = []
        for m in (raw_msgs or []):
            out.append({
                "id": m["id"],
                "subject": m.get("subject", ""),
                "sender_name": m.get("sender_name", ""),
                "sender_email": m.get("sender_email", ""),
                "company": m.get("company", ""),
                "date_sent": m.get("date_str", ""),
                "body_text": m.get("raw_body_text") or m.get("body_text", ""),
                "clean_body": m.get("clean_body") or m.get("body_text", ""),
                "conversation_id": m.get("conversation_id", ""),
                "is_read": 1 if m.get("is_read") else 0,
                "flag_status": 2 if m.get("is_flagged") else 0,
                "attachments": list(m.get("attachments") or []),
            })
        return out

    def _pin_item_in_place(
        self,
        it: dict,
        *,
        linked_ais: Optional[List[str]] = None,
        linked_wis: Optional[List[str]] = None,
        linked_sub_items: Optional[List[str]] = None,
        reload: bool = True,
    ) -> Optional[str]:
        if self.tracked_db is None or not it:
            return None
        subj = str(it.get("thread_title") or it.get("subject") or "Email Discussion").strip()
        norm = normalize_subject(subj) or subj
        tid = pin_thread_by_subject_or_cid(
            self.tracked_db,
            title=norm,
            normalized_subject=norm,
            conversation_id=str(it.get("conversation_id") or ""),
            meeting_id=self.meeting_id,
            wg=self.wg,
            agenda_map=self._agenda_map,
            valid_wis=self._valid_wis if self._wi_suggestions else None,
            tdoc_meta_map=self._tdoc_meta_map,
            linked_ais=linked_ais,
            linked_wis=linked_wis,
            linked_tdocs=list(it.get("linked_tdocs") or []),
            linked_sub_items=linked_sub_items,
            seed_messages=self._seed_messages_for_item(it),
            # Already seeded from local DBs; never run synchronous Outlook COM scans on the UI thread
            refresh_from_outlook=False,
        )
        new_preserve = f"THR::{tid}" if self._view_mode == "THREADS" else it["id"]
        if reload:
            self.reload_from_dbs(preserve_id=new_preserve)
            self.data_changed.emit()
        return tid

    def _on_pin_inline_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        unpinned = [it for it in targets if not it.get("in_tracked_db")]
        if not unpinned or self.tracked_db is None:
            return
        if len(targets) == 1:
            tid = self._pin_item_in_place(unpinned[0])
            if tid:
                notify_success(self, "Thread Pinned", "Email conversation pinned as a Tracked Thread.")
            return

        preserved_ids: List[str] = []
        pinned_count = 0
        for it in targets:
            if it.get("in_tracked_db"):
                preserved_ids.append(it["id"])
                continue
            tid = self._pin_item_in_place(it, reload=False)
            if tid:
                pinned_count += 1
                preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
        self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
        self.data_changed.emit()
        if pinned_count:
            notify_success(self, "Threads Pinned", f"Pinned {pinned_count} conversation(s) as Tracked Threads.")

    def _on_auto_assign_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets or self.tracked_db is None:
            return

        preserved_ids: List[str] = []
        seen_tids = set()
        for it in targets:
            sugg = auto_suggest_links(
                it.get("subject", ""),
                it.get("body_text", ""),
                agenda_map=self._agenda_map,
                valid_wis=self._valid_wis if self._wi_suggestions else None,
                tdoc_meta_map=self._tdoc_meta_map,
                extra_tdocs=it.get("linked_tdocs") or [],
            )
            ais = list(dict.fromkeys((it.get("linked_ais") or []) + (sugg.get("ais") or [])))
            subs = list(dict.fromkeys((it.get("linked_sub_items") or []) + (sugg.get("sub_items") or [])))
            wis = list(dict.fromkeys((it.get("linked_wis") or []) + (sugg.get("wis") or [])))
            if not it.get("in_tracked_db"):
                tid = self._pin_item_in_place(it, linked_ais=ais, linked_wis=wis, linked_sub_items=subs, reload=False)
                if tid:
                    preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
            else:
                tid = it["thread_id"]
                if tid not in seen_tids:
                    seen_tids.add(tid)
                    self.tracked_db.update_thread_links(tid, linked_ais=ais, linked_sub_items=subs, linked_wis=wis)
                preserved_ids.append(it["id"])

        dedup_ids = list(dict.fromkeys(preserved_ids))
        if len(targets) == 1 and len(dedup_ids) == 1:
            self.reload_from_dbs(preserve_id=dedup_ids[0])
        else:
            self.reload_from_dbs(preserve_ids=dedup_ids)
        self.data_changed.emit()

    def _on_inline_links_changed(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if self._updating_assign_bar or not targets or self.tracked_db is None:
            return

        if len(targets) >= 2:
            cur_ais = self.assign_ais_pills.tags()
            cur_subs = self.assign_subs_pills.tags()
            cur_wis = self.assign_wis_pills.tags()

            base_ais = set(self._multi_base_common.get("ais", []))
            base_subs = set(self._multi_base_common.get("subs", []))
            base_wis = set(self._multi_base_common.get("wis", []))

            added_ais = [a for a in cur_ais if a not in base_ais]
            removed_ais = base_ais - set(cur_ais)
            added_subs = [s for s in cur_subs if s not in base_subs]
            removed_subs = base_subs - set(cur_subs)
            added_wis = [w for w in cur_wis if w not in base_wis]
            removed_wis = base_wis - set(cur_wis)

            preserved_ids: List[str] = []
            seen_tids = set()
            all_resulting_ais = list(cur_ais)
            for it in targets:
                it_ais = [a for a in (it.get("linked_ais") or []) if a not in removed_ais]
                for a in added_ais:
                    if a not in it_ais:
                        it_ais.append(a)
                for a in it_ais:
                    if a not in all_resulting_ais:
                        all_resulting_ais.append(a)

                it_wis = [w for w in (it.get("linked_wis") or []) if w not in removed_wis]
                for w in added_wis:
                    if w not in it_wis:
                        it_wis.append(w)

                it_subs = [s for s in (it.get("linked_sub_items") or []) if s not in removed_subs]
                for s in added_subs:
                    if s not in it_subs:
                        it_subs.append(s)
                if self._ai_to_sub_items:
                    it_allowed = {sub for ai in it_ais for sub in self._ai_to_sub_items.get(ai, [])}
                    it_subs = [s for s in it_subs if s in it_allowed]

                if not it.get("in_tracked_db"):
                    tid = self._pin_item_in_place(
                        it,
                        linked_ais=it_ais,
                        linked_wis=it_wis,
                        linked_sub_items=it_subs,
                        reload=False,
                    )
                    if tid:
                        preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
                else:
                    tid = it["thread_id"]
                    if tid not in seen_tids:
                        seen_tids.add(tid)
                        self.tracked_db.update_thread_links(
                            tid,
                            linked_ais=it_ais,
                            linked_sub_items=it_subs,
                            linked_wis=it_wis,
                        )
                    preserved_ids.append(it["id"])

            self._multi_base_common = {
                "ais": list(cur_ais),
                "subs": list(cur_subs),
                "wis": list(cur_wis),
            }
            self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
            self.data_changed.emit()
            return

        it = targets[0]
        selected_ais = self.assign_ais_pills.tags()
        self._refresh_assign_sub_suggestions(selected_ais)
        allowed_subs = {
            sub
            for ai in selected_ais
            for sub in self._ai_to_sub_items.get(ai, [])
        }
        selected_subs = self.assign_subs_pills.tags()
        if self._ai_to_sub_items:
            pruned_subs = [s for s in selected_subs if s in allowed_subs]
            if pruned_subs != selected_subs:
                selected_subs = pruned_subs
                self._updating_assign_bar = True
                try:
                    self.assign_subs_pills.setText(", ".join(pruned_subs))
                finally:
                    self._updating_assign_bar = False
        selected_wis = self.assign_wis_pills.tags()

        if not it.get("in_tracked_db"):
            self._pin_item_in_place(
                it,
                linked_ais=selected_ais,
                linked_wis=selected_wis,
                linked_sub_items=selected_subs,
            )
            return

        tid = it["thread_id"]
        self.tracked_db.update_thread_links(
            tid,
            linked_ais=selected_ais,
            linked_sub_items=selected_subs,
            linked_wis=selected_wis,
        )
        for m in self._all_items:
            if m.get("thread_id") == tid:
                m["linked_ais"] = list(selected_ais)
                m["linked_sub_items"] = list(selected_subs)
                m["linked_wis"] = list(selected_wis)
        self._grouped_items = self._build_grouped_thread_items()
        cur_row = self.table.currentRow()
        if 0 <= cur_row < len(self._visible_items):
            self._visible_items[cur_row]["linked_ais"] = list(selected_ais)
            self._visible_items[cur_row]["linked_sub_items"] = list(selected_subs)
            self._visible_items[cur_row]["linked_wis"] = list(selected_wis)
            ai_cell = self.table.item(cur_row, 2)
            if ai_cell:
                ai_cell.setText(", ".join(selected_ais) if selected_ais else "❓ —")
                ai_cell.setForeground(QColor("#0F172A" if selected_ais else "#B45309"))
        self._populate_filter_combos()
        self.data_changed.emit()

    def _mutate_read(self, it: dict, is_read: bool):
        targets = (
            [m["id"] for m in (it.get("_group_messages") or []) if m["is_read"] != is_read]
            if it.get("_is_group")
            else [it["id"]]
        )
        if not targets:
            return
        if self.tracked_db is not None:
            for eid in targets:
                self.tracked_db.set_message_read_status(eid, is_read)
        if self.general_email_db is not None and hasattr(self.general_email_db, "set_emails_read_status"):
            self.general_email_db.set_emails_read_status(targets, is_read, sync_outlook=False)
        outlook_ids = [eid for eid in targets if not str(eid).startswith("MSG-")]
        if outlook_ids:
            OutlookClient.set_messages_read_status(outlook_ids, is_read=is_read)

        tset = set(targets)
        for m in self._all_items:
            if m["id"] in tset:
                m["is_read"] = bool(is_read)
        self._grouped_items = self._build_grouped_thread_items()
        self._update_pill_counts()
        self.data_changed.emit()

    def _style_detail_status_combo(self, status_val: str):
        if not hasattr(self, "detail_status_combo"):
            return
        self.detail_status_combo.setStyleSheet(get_status_combo_qss(status_val))

    def _populate_status_and_note_bars(self, targets: List[dict]):
        if not hasattr(self, "detail_status_combo"):
            return
        self._updating_status_bar = True
        self._updating_note_bar = True
        try:
            self.detail_status_combo.setEnabled(bool(targets) and self.tracked_db is not None)
            self.btn_toggle_note.setEnabled(bool(targets) and self.tracked_db is not None)
            statuses = [str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral" for it in targets]
            common_status = statuses[0] if (statuses and len(set(statuses)) == 1) else "⚪ Neutral"
            idx = self.detail_status_combo.findText(common_status)
            self.detail_status_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self._style_detail_status_combo(common_status)

            if len(targets) == 1:
                it = targets[0]
                note_txt = str(it.get("user_notes") or "").strip()
                self._note_target_item = it
                self._note_dirty = False
                self.note_edit.setPlainText(note_txt)
                self.note_saved_lbl.setText("")
                has_note = bool(note_txt)
                self.note_frame.setVisible(has_note)
                self.btn_toggle_note.setChecked(has_note)
                self.btn_toggle_note.setText("📝 Note ✓" if has_note else "📝 Note")
                self.note_edit.setEnabled(True)
            else:
                self._note_target_item = None
                self._note_dirty = False
                self.note_frame.setVisible(False)
                self.btn_toggle_note.setChecked(False)
                self.btn_toggle_note.setText("📝 Note")
        finally:
            self._updating_status_bar = False
            self._updating_note_bar = False

    def _refresh_table_state_cells_for_thread(self, tid: str):
        for r, vis_it in enumerate(self._visible_items):
            if str(vis_it.get("thread_id") or "") != tid:
                continue
            item_cell = self.table.item(r, 0)
            if not item_cell:
                continue
            u_status = str(vis_it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
            u_notes = str(vis_it.get("user_notes") or "").strip()
            note_badge = "📝" if u_notes else ""
            if vis_it.get("_is_group"):
                u_cnt = int(vis_it.get("_unread_count", 0))
                t_cnt = int(vis_it.get("_total_count", 1))
                f_part = "🚩" if vis_it["is_flagged"] else ""
                core_s = f"🔵 {u_cnt}/{t_cnt} {f_part}".strip() if u_cnt > 0 else f"✓ {t_cnt} {f_part}".strip()
                state_txt = f"{core_s} {note_badge}".strip()
            else:
                parts = []
                if not vis_it["is_read"]:
                    parts.append("🔵")
                if vis_it["is_flagged"]:
                    parts.append("🚩")
                if not parts:
                    parts.append("✓")
                if note_badge:
                    parts.append(note_badge)
                state_txt = " ".join(parts)
            item_cell.setText(state_txt)
            item_cell.setData(Qt.UserRole + 1, _status_color_hex(u_status))
            st_tip = [f"Status: {u_status}"] if u_status != "⚪ Neutral" else []
            if u_notes:
                st_tip.append(f"Comment: {(u_notes[:180] + '…') if len(u_notes) > 180 else u_notes}")
            item_cell.setToolTip("\n".join(st_tip))

    def _on_detail_status_changed(self):
        if getattr(self, "_updating_status_bar", False) or self.tracked_db is None:
            return
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            return
        new_status = str(self.detail_status_combo.currentText() or "⚪ Neutral").strip() or "⚪ Neutral"
        self._style_detail_status_combo(new_status)

        preserved_ids: List[str] = []
        seen_tids = set()
        needs_full_reload = False
        for it in targets:
            if not it.get("in_tracked_db"):
                tid = self._pin_item_in_place(it, reload=False)
                needs_full_reload = True
            else:
                tid = str(it.get("thread_id") or "")
            if tid and tid not in seen_tids:
                seen_tids.add(tid)
                self.tracked_db.update_thread_annotations(tid, user_status=new_status)
                for m in self._all_items:
                    if m.get("thread_id") == tid:
                        m["user_status"] = new_status
                for vis in self._visible_items:
                    if vis.get("thread_id") == tid:
                        vis["user_status"] = new_status
            preserved_ids.append(f"THR::{tid}" if (self._view_mode == "THREADS" and tid) else it["id"])

        self._grouped_items = self._build_grouped_thread_items()
        self._populate_filter_combos()
        status_filter_active = hasattr(self, "status_combo") and self.status_combo.isFiltered()
        if needs_full_reload or status_filter_active:
            self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
        else:
            for tid in seen_tids:
                self._refresh_table_state_cells_for_thread(tid)
            self._update_sync_button_state(
                total_count=len(self._active_dataset()),
                visible_count=len(self._visible_items),
                has_narrow_filter=bool(
                    self.company_combo.isFiltered()
                    or self.ai_combo.isFiltered()
                    or (hasattr(self, "status_combo") and self.status_combo.isFiltered())
                    or self.search_edit.text().strip()
                    or (self.source_combo.currentData() or "").strip()
                    or self._active_mode == "FLAGGED"
                ),
            )
        self.data_changed.emit()

    def _on_toggle_note_clicked(self, checked: bool):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if len(targets) != 1:
            if len(targets) > 1:
                notify_info(self, "Select Single Thread", "Select a single thread or email to add/edit its comment.")
            self.btn_toggle_note.setChecked(False)
            return
        if checked or str(self.note_edit.toPlainText() or "").strip():
            self.note_frame.setVisible(True)
            self.note_edit.setFocus()
        else:
            self.note_frame.setVisible(False)

    def _on_note_text_changed(self):
        if getattr(self, "_updating_note_bar", False):
            return
        self._note_dirty = True
        if hasattr(self, "note_saved_lbl"):
            self.note_saved_lbl.setText("Saving...")
        self._note_save_timer.start()

    def _flush_pending_note_save(self, reload_view: bool = True):
        if self._note_save_timer.isActive():
            self._note_save_timer.stop()
        if not getattr(self, "_note_dirty", False) or self.tracked_db is None:
            return
        it = self._note_target_item
        if not it:
            self._note_dirty = False
            return
        new_note = self.note_edit.toPlainText().strip()
        self._note_dirty = False
        if not it.get("in_tracked_db"):
            if not new_note:
                if hasattr(self, "note_saved_lbl"):
                    self.note_saved_lbl.setText("")
                return
            tid = self._pin_item_in_place(it, reload=False)
            if tid:
                self.tracked_db.update_thread_annotations(tid, user_notes=new_note)
                if reload_view:
                    pres = f"THR::{tid}" if self._view_mode == "THREADS" else it["id"]
                    self.reload_from_dbs(preserve_id=pres)
        else:
            tid = str(it.get("thread_id") or "")
            if tid:
                self.tracked_db.update_thread_annotations(tid, user_notes=new_note)
                for m in self._all_items:
                    if m.get("thread_id") == tid:
                        m["user_notes"] = new_note
                for vis in self._visible_items:
                    if vis.get("thread_id") == tid:
                        vis["user_notes"] = new_note
                self._grouped_items = self._build_grouped_thread_items()
                self._refresh_table_state_cells_for_thread(tid)
        if hasattr(self, "note_saved_lbl"):
            self.note_saved_lbl.setText("✓ Saved" if new_note else "")
        if hasattr(self, "btn_toggle_note"):
            has_note = bool(new_note)
            self.btn_toggle_note.setText("📝 Note ✓" if has_note else "📝 Note")
        self.data_changed.emit()

    def _mutate_flag(self, it: dict, is_flagged: bool):
        targets = (
            [it["_group_messages"][-1]["id"]]
            if it.get("_is_group") and it.get("_group_messages")
            else [it["id"]]
        )
        if self.tracked_db is not None:
            for eid in targets:
                self.tracked_db.set_message_flag_status(eid, is_flagged)
        if self.general_email_db is not None and hasattr(self.general_email_db, "set_emails_flag_status"):
            self.general_email_db.set_emails_flag_status(targets, is_flagged, sync_outlook=False)
        outlook_ids = [eid for eid in targets if not str(eid).startswith("MSG-")]
        if outlook_ids:
            OutlookClient.set_messages_flag_status(outlook_ids, is_flagged=is_flagged)

        tset = set(targets)
        for m in self._all_items:
            if m["id"] in tset:
                m["is_flagged"] = bool(is_flagged)
        self._grouped_items = self._build_grouped_thread_items()
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
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            return
        if len(targets) == 1:
            it = targets[0]
            cur_row = self.table.currentRow()
            self._mutate_read(it, not it["is_read"])
            self._apply_filters(
                preserve_id=it["id"],
                select_row=cur_row if self._active_mode == "UNREAD" else None,
            )
            return
        mark_read = any(not it["is_read"] for it in targets)
        for it in targets:
            self._mutate_read(it, mark_read)
        self._apply_filters(preserve_ids=[it["id"] for it in targets])

    def _on_toggle_flag(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            return
        if len(targets) == 1:
            it = targets[0]
            cur_row = self.table.currentRow()
            self._mutate_flag(it, not it["is_flagged"])
            self._apply_filters(
                preserve_id=it["id"],
                select_row=cur_row if self._active_mode == "FLAGGED" else None,
            )
            return
        flag_on = any(not it["is_flagged"] for it in targets)
        for it in targets:
            self._mutate_flag(it, flag_on)
        self._apply_filters(preserve_ids=[it["id"] for it in targets])

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
        if scheme == "tdoc" and self._open_tdoc_callback:
            target_tdoc = url.path().strip() or url.toString().split(":", 1)[1].strip()
            if target_tdoc:
                self._open_tdoc_callback(target_tdoc.upper())
            return

        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if not path_parts:
                return
            msg_id = path_parts[0]
            if host == "msg":
                if not open_outlook_message(msg_id):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply":
                if not reply_to_outlook_message(msg_id, reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
            elif host == "reload" and self._current_item and self._current_item.get("thread_id") and self.tracked_db:
                sb = self.preview_browser.verticalScrollBar()
                prev_scroll = sb.value()
                if reload_outlook_message(self.tracked_db, self._current_item["thread_id"], msg_id):
                    self.reload_from_dbs(preserve_id=self._current_item["id"])
                    sb.setValue(min(prev_scroll, sb.maximum()))
                    self.data_changed.emit()
            elif host in ("read", "unread"):
                target_item = self._items_by_id.get(msg_id)
                if target_item:
                    sb = self.preview_browser.verticalScrollBar()
                    prev_scroll = sb.value()
                    cur_sel = self._current_item["id"] if self._current_item else None
                    self._mutate_read(target_item, is_read=(host == "read"))
                    self._apply_filters(preserve_id=cur_sel)
                    sb.setValue(min(prev_scroll, sb.maximum()))
            elif host in ("flag", "unflag"):
                target_item = self._items_by_id.get(msg_id)
                if target_item:
                    sb = self.preview_browser.verticalScrollBar()
                    prev_scroll = sb.value()
                    cur_sel = self._current_item["id"] if self._current_item else None
                    self._mutate_flag(target_item, is_flagged=(host == "flag"))
                    self._apply_filters(preserve_id=cur_sel)
                    sb.setValue(min(prev_scroll, sb.maximum()))
            elif host == "copy":
                target_item = self._items_by_id.get(msg_id)
                if target_item:
                    QApplication.clipboard().setText(str(target_item.get("body_text") or "").strip())
                    notify_info(self, "Copied", "Message copied to clipboard.")
            elif host == "att" and len(path_parts) >= 2:
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

    def _update_sync_button_state(self, total_count: int = 0, visible_count: int = 0, has_narrow_filter: bool = False):
        if not hasattr(self, "sync_outlook_btn"):
            return
        can_scope = (0 < visible_count < total_count)
        self._default_sync_scoped = bool(can_scope and has_narrow_filter)
        self._act_sync_filtered.setEnabled(visible_count > 0)
        self._act_sync_filtered.setText(f"⚡ Sync Filtered Only ({visible_count})")
        if self._default_sync_scoped:
            self.sync_outlook_btn.setStyleSheet(
                f"QToolButton {{ background-color: {C.INFO_SOFT}; color: {C.INFO}; "
                f"border: 1px solid {C.INFO_BORDER}; border-radius: 5px; padding: 4px 20px 4px 9px; font-weight: bold; }} "
                f"QToolButton:hover {{ background-color: {C.ACCENT_SOFT_HOVER}; }} "
                f"QToolButton:disabled {{ background-color: {C.BG_MUTED}; color: {C.TEXT_FAINT}; border: 1px solid {C.BORDER_STRONG}; }} "
                f"QToolButton::menu-button {{ border-left: 1px solid {C.INFO_BORDER}; width: 18px; }}"
            )
        else:
            self.sync_outlook_btn.setStyleSheet(
                f"QToolButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 5px; padding: 4px 20px 4px 9px; }} "
                f"QToolButton:hover {{ background-color: {C.BG_HOVER}; }} "
                f"QToolButton:disabled {{ background-color: {C.BG_MUTED}; color: {C.TEXT_FAINT}; border: 1px solid {C.BORDER_STRONG}; }} "
                f"QToolButton::menu-button {{ border-left: 1px solid {C.BORDER_SUBTLE}; width: 18px; }}"
            )
        if not (getattr(self, "_sync_worker", None) and self._sync_worker.isRunning()):
            if self._default_sync_scoped:
                self.sync_outlook_btn.setText(f"⚡ Sync Filtered ({visible_count})")
                self.sync_outlook_btn.setToolTip(
                    f"Fast sync ONLY the {visible_count} filtered conversation(s) with Outlook (skips full folder scan). "
                    "Use dropdown ▾ for Full Sync."
                )
            else:
                self.sync_outlook_btn.setText("📡 Sync Outlook")
                self.sync_outlook_btn.setToolTip(
                    "Pull new Tracked Thread replies, TDoc emails, and Outlook read/flag updates asynchronously. "
                    "Use dropdown ▾ to sync only currently visible items."
                )

    def _update_selected_sync_controls(self, selected_count: int):
        running = bool(getattr(self, "_sync_worker", None) and self._sync_worker.isRunning())
        if hasattr(self, "btn_sync_selected"):
            self.btn_sync_selected.setEnabled(selected_count > 0 and not running)
            self.btn_sync_selected.setText(
                f"🔄 Sync ({selected_count})" if selected_count > 1 else "🔄 Sync"
            )
        if hasattr(self, "_act_sync_selected"):
            self._act_sync_selected.setEnabled(selected_count > 0 and not running)
            self._act_sync_selected.setText(
                f"🔄 Sync Selected Threads ({selected_count})" if selected_count > 1 else "🔄 Sync Selected Thread"
            )

    def _on_table_context_menu(self, pos):
        item = self.table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        sel_rows = {idx.row() for idx in self.table.selectionModel().selectedRows()} if self.table.selectionModel() else set()
        if row not in sel_rows:
            self.table.selectRow(row)
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            return
        menu = QMenu(self.table)
        sync_label = (
            f"🔄 Sync Selected Threads ({len(targets)})"
            if len(targets) > 1
            else ("🔄 Sync This Thread" if targets[0].get("_is_group") or targets[0].get("in_tracked_db") else "🔄 Sync This Conversation")
        )
        act_sync = menu.addAction(sync_label)
        act_sync.setEnabled(not bool(getattr(self, "_sync_worker", None) and self._sync_worker.isRunning()))
        act_sync.triggered.connect(self._on_sync_selected_clicked)
        menu.addSeparator()
        act_read = menu.addAction(self.btn_toggle_read.text())
        act_read.triggered.connect(self._on_toggle_read)
        act_flag = menu.addAction(self.btn_toggle_flag.text())
        act_flag.triggered.connect(self._on_toggle_flag)
        if self.btn_open_context.isEnabled():
            menu.addSeparator()
            act_ctx = menu.addAction(self.btn_open_context.text())
            act_ctx.triggered.connect(self._on_open_context)
        if self.tracked_db is not None:
            merge_targets = self._distinct_mergeable_targets(targets)
            if len(merge_targets) >= 2:
                menu.addSeparator()
                act_merge = menu.addAction(f"🔗 Merge Selected Threads ({len(merge_targets)})...")
                act_merge.triggered.connect(lambda _chk=False, mt=merge_targets: self._on_merge_selected_clicked(mt))
            elif len(merge_targets) == 1:
                menu.addSeparator()
                act_merge = menu.addAction("🔗 Merge With Another Thread...")
                act_merge.triggered.connect(lambda _chk=False, mt=merge_targets: self._on_merge_selected_clicked(mt))
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def _enrich_merge_item(self, it: dict) -> dict:
        enriched = dict(it)
        tid = str(enriched.get("thread_id") or "").strip()
        if tid and self.tracked_db is not None:
            rec = self.tracked_db.get_thread(tid)
            if rec:
                merged = dict(rec)
                merged.update({
                    k: enriched[k]
                    for k in ("_total_count", "_unread_count", "_companies", "_group_messages", "in_tracked_db", "date_str", "company")
                    if k in enriched and enriched[k] is not None
                })
                if not merged.get("in_tracked_db"):
                    merged["in_tracked_db"] = True
                return merged
        return enriched

    def _distinct_mergeable_targets(self, items: List[dict]) -> List[dict]:
        seen = set()
        out: List[dict] = []
        for it in (items or []):
            tid = str(it.get("thread_id") or "").strip()
            key = f"THR::{tid}" if tid else f"ITEM::{it.get('id') or id(it)}"
            if key in seen:
                continue
            seen.add(key)
            out.append(self._enrich_merge_item(it))
        return out

    def _collect_all_merge_candidates(self) -> List[dict]:
        seen = set()
        out: List[dict] = []
        if self.tracked_db is not None and self.meeting_id:
            for t in self.tracked_db.get_threads_for_meeting(self.meeting_id):
                tid = str(t.get("thread_id") or "").strip()
                if not tid:
                    continue
                k = f"THR::{tid}"
                if k not in seen:
                    seen.add(k)
                    rec = dict(t)
                    rec["in_tracked_db"] = True
                    out.append(rec)
        for g in (getattr(self, "_grouped_items", None) or []):
            tid = str(g.get("thread_id") or "").strip()
            k = f"THR::{tid}" if tid else f"ITEM::{g.get('id') or ''}"
            if k and k not in seen:
                seen.add(k)
                out.append(self._enrich_merge_item(g))
        return out

    def _on_merge_selected_clicked(self, explicit_targets: Optional[List[dict]] = None):
        if self.tracked_db is None:
            return
        raw_targets = explicit_targets if explicit_targets is not None else (
            self._selected_items or ([self._current_item] if self._current_item else [])
        )
        merge_targets = self._distinct_mergeable_targets(raw_targets)
        if not merge_targets:
            return

        from modules.emails.ui.merge_threads_dialog import MergeThreadsDialog

        all_candidates = self._collect_all_merge_candidates()
        if len(merge_targets) >= 2:
            dlg = MergeThreadsDialog(
                selected_items=merge_targets,
                candidates=all_candidates,
                parent=self,
            )
        else:
            single = merge_targets[0]
            single_tid = str(single.get("thread_id") or "").strip()
            single_id = str(single.get("id") or "").strip()
            others = [
                c for c in all_candidates
                if (not single_tid or str(c.get("thread_id") or "").strip() != single_tid)
                and (not single_id or str(c.get("id") or "").strip() != single_id)
            ]
            if not others:
                notify_info(self, "No Other Threads", "No other tracked or grouped threads available to merge.")
                return
            dlg = MergeThreadsDialog(
                target_thread=single,
                candidates=others,
                parent=self,
            )

        if dlg.exec_() != QDialog.Accepted:
            return

        target_item = dlg.target_item()
        source_items = dlg.selected_source_items()
        if not source_items:
            return

        target_tid = str(target_item.get("thread_id") or "").strip()
        if not target_tid:
            target_tid = self._pin_item_in_place(target_item, reload=False) or ""
        if not target_tid:
            return

        source_tids: List[str] = []
        for src in source_items:
            stid = str(src.get("thread_id") or "").strip()
            if not stid:
                stid = self._pin_item_in_place(src, reload=False) or ""
            if stid and stid != target_tid and stid not in source_tids:
                source_tids.append(stid)

        if source_tids:
            self.tracked_db.merge_threads(target_tid, source_tids)

        preserve_id = f"THR::{target_tid}" if self._view_mode == "THREADS" else None
        self.reload_from_dbs(preserve_id=preserve_id)
        self.data_changed.emit()
        notify_success(
            self,
            "Threads Merged",
            f"Merged {len(source_items)} thread(s) into {target_tid}.",
        )

    def _collect_scoped_sync_targets(self, items: Optional[List[dict]] = None) -> tuple:
        tracked_ids: List[str] = []
        seen_t = set()
        unpinned_ids: List[str] = []
        seen_u = set()
        source = self._visible_items if items is None else items
        for it in (source or []):
            tid = str(it.get("thread_id") or "").strip()
            if it.get("in_tracked_db") and tid:
                if tid not in seen_t:
                    seen_t.add(tid)
                    tracked_ids.append(tid)
            elif it.get("_is_group") and it.get("_group_messages"):
                for gm in it["_group_messages"]:
                    gid = str(gm.get("id") or "").strip()
                    if gid and gid not in seen_u:
                        seen_u.add(gid)
                        unpinned_ids.append(gid)
            else:
                eid = str(it.get("id") or "").strip()
                if eid and eid not in seen_u:
                    seen_u.add(eid)
                    unpinned_ids.append(eid)
        return tracked_ids, unpinned_ids

    def _on_sync_selected_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            notify_info(self, "No Thread Selected", "Select an email or thread first.")
            return
        self._start_outlook_sync(scoped=True, target_items=targets)

    def _on_sync_outlook_clicked(self):
        self._start_outlook_sync(scoped=bool(getattr(self, "_default_sync_scoped", False)))

    def _start_outlook_sync(self, scoped: bool = False, target_items: Optional[List[dict]] = None):
        if getattr(self, "_sync_worker", None) and self._sync_worker.isRunning():
            return
        tdb_path = getattr(self.tracked_db, "db_path", None) if self.tracked_db is not None else None
        gdb_path = getattr(self.general_email_db, "db_path", None) if self.general_email_db is not None else None
        if not tdb_path and not gdb_path:
            self.reload_from_dbs()
            return

        from modules.emails.core.tracked_threads_workers import InboxOutlookSyncWorker

        scoped_tids = None
        scoped_eids = None
        if scoped:
            scoped_tids, scoped_eids = self._collect_scoped_sync_targets(target_items)
            if not scoped_tids and not scoped_eids:
                notify_info(self, "Nothing in Filter", "No conversations in the current filter to sync.")
                return

        self.sync_outlook_btn.setEnabled(False)
        self.sync_outlook_btn.setText("⏳ Syncing...")
        if hasattr(self, "btn_sync_selected"):
            self.btn_sync_selected.setEnabled(False)
        if hasattr(self, "_act_sync_selected"):
            self._act_sync_selected.setEnabled(False)
        worker = InboxOutlookSyncWorker(
            tdb_path,
            gdb_path,
            meeting_id=self.meeting_id,
            wg=self.wg,
            start_date=self._meeting_start,
            end_date=self._meeting_end,
            parent=self,
            scoped_thread_ids=scoped_tids,
            scoped_email_ids=scoped_eids,
        )
        worker.progress_text.connect(lambda msg: self.status_lbl.setText(f"⏳ {msg}"))
        worker.finished_sync.connect(self._on_inbox_sync_finished)
        self._sync_worker = worker
        worker.start()

    def _on_inbox_sync_finished(self, ok: bool, new_thread_msgs: int, general_synced: int, err_msg: str):
        self.sync_outlook_btn.setEnabled(True)
        self._sync_worker = None
        self.reload_from_dbs()
        self.data_changed.emit()
        if not ok:
            notify_toast(
                self,
                "Outlook Sync Warning",
                err_msg or "Outlook sync did not finish cleanly.",
                variant="warning",
            )
            return
        if new_thread_msgs > 0 or general_synced > 0:
            notify_success(
                self,
                "Inbox Synced with Outlook",
                f"Synced {new_thread_msgs} new tracked message(s) and {general_synced} TDoc email match(es).",
            )
        else:
            notify_info(self, "Inbox Up to Date", "No new matching messages found in Outlook.")
