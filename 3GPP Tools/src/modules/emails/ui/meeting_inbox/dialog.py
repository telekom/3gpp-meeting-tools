"""Unified Email Triage Cockpit dialog composed from modular mixins."""

from __future__ import annotations

from typing import Any, Callable, Dict, List, Optional

from PyQt5.QtCore import QSettings, Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import QDialog, QWidget

from modules.emails.core.sync_watermark import (
    DEFAULT_LIVE_SYNC_INTERVAL_MS,
    load_meeting_live_sync_enabled,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .mixins import (
    InboxDataAndFilteringMixin,
    InboxDetailRenderingMixin,
    InboxInlineAssignmentMixin,
    InboxLayoutAndSettingsMixin,
    InboxSyncAndMergeMixin,
    InboxTriageActionsMixin,
    InboxUiSetupMixin,
)


class MeetingEmailInboxDialog(
    InboxUiSetupMixin,
    InboxLayoutAndSettingsMixin,
    InboxDataAndFilteringMixin,
    InboxDetailRenderingMixin,
    InboxInlineAssignmentMixin,
    InboxTriageActionsMixin,
    InboxSyncAndMergeMixin,
    QDialog,
):
    """Unified cockpit for triaging unread, flagged, unassigned, and unpinned emails/threads."""

    data_changed = pyqtSignal()
    live_sync_toggled = pyqtSignal(bool)

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
        tdocs_db: Optional[Any] = None,
        user_data_callback: Optional[Callable[[str, str, str], None]] = None,
        jump_tdoc_callback: Optional[Callable[[str], None]] = None,
    ):
        super().__init__(parent)
        self.tracked_db = tracked_db
        self.general_email_db = general_email_db
        self.tdocs_db = tdocs_db
        self._user_data_callback = user_data_callback
        self._jump_tdoc_callback = jump_tdoc_callback
        self.meeting_id = str(meeting_id or "").strip()
        self.wg = str(wg or "SA2").strip().upper()
        self._meeting_start = str(meeting_start or "").strip()
        self._meeting_end = str(meeting_end or "").strip()
        self._sync_worker = None
        self._stats_thread = None
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
        self._tdoc_grouped_items: List[dict] = []
        self._visible_items: List[dict] = []
        self._items_by_id: Dict[str, dict] = {}
        self._selected_items: List[dict] = []
        self._current_item: Optional[dict] = None
        self._multi_base_common: Dict[str, List[str]] = {"ais": [], "subs": [], "wis": []}
        self._settings = settings if settings is not None else QSettings("3GPP_Tools", "MeetingEmailInboxDialog")
        self._has_custom_settings = settings is not None
        self._live_sync_store = self._settings if self._has_custom_settings else None
        self._last_sync_silent = False
        self._live_catchup_done = False
        self._live_sync_enabled = load_meeting_live_sync_enabled(self.meeting_id, settings=self._live_sync_store)
        self._live_sync_timer = QTimer(self)
        self._live_sync_timer.setInterval(DEFAULT_LIVE_SYNC_INTERVAL_MS)
        self._live_sync_timer.timeout.connect(self._on_live_sync_timer_tick)
        self._restoring_filters = True
        self._active_mode = "UNREAD"
        self._view_mode = "MESSAGES"
        self._hide_objected = False
        self._followed_ais: set = set()
        self._inherited_followed_ais: set = set()
        self._followed_ais_customized = False
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

        if self._live_sync_enabled:
            self._live_sync_timer.start()
