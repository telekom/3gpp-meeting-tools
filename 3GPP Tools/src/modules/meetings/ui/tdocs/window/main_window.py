# --- File: src/modules/meetings/ui/tdocs/window/main_window.py ---
"""
TDocsWindow shell: constructor, Qt lifecycle events, cache setup and user-data persistence.
Feature logic lives in the mixins/ package (see README.md for a 'where to edit what' map).
"""
import datetime
import json
import os
import re
from pathlib import Path
from typing import Optional
from PyQt5.QtCore import (
    QTimer,
    pyqtSignal,
    QPoint,
    QSettings,
    Qt,
)
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTableView
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.ui.tdocs_menu import handle_table_context_menu
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.sync_watermark import DEFAULT_LIVE_SYNC_INTERVAL_MS
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.meetings.ui.tdocs_drop_overlay import DropOverlayWidget
from modules.meetings.ui.tdocs.window.mixins import (
    HeaderMixin,
    FilterSetupMixin,
    TableLayoutMixin,
    TrackedWIFilterMixin,
    FilterFacetsMixin,
    TDocActionsMixin,
    EmailMixin,
    EmailSyncMixin,
    TrackedThreadsMixin,
    ExportsMixin,
    MeetingDataMixin,
    DragDropMixin,
)


class TDocsWindow(
    HeaderMixin,
    FilterSetupMixin,
    TableLayoutMixin,
    TrackedWIFilterMixin,
    FilterFacetsMixin,
    TDocActionsMixin,
    EmailMixin,
    EmailSyncMixin,
    TrackedThreadsMixin,
    ExportsMixin,
    MeetingDataMixin,
    DragDropMixin,
    QWidget,
):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str, filter_settings: Optional[QSettings] = None):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self._filter_settings = filter_settings if filter_settings is not None else QSettings("3GPP_Tools", "TDocsWindow")
        self._restoring_filters = True
        self._saved_filters = {}
        self._pending_combo_filters = {}
        self.active_threads = {}
        self.chair_agreements = []
        self._email_dialogs = {}
        self._triage_dialogs = {}
        self._highlight_clear_timer = QTimer(self)
        self._highlight_clear_timer.setSingleShot(True)
        self._highlight_clear_timer.setInterval(4500)
        self._highlight_clear_timer.timeout.connect(
            lambda: self.model.set_highlighted_tdoc(None) if hasattr(self, "model") else None
        )

        self.agenda_dir = Path(filepath).parent
        self.agenda_map = AgendaManager.load_local_agenda(self.agenda_dir)
        self.agenda_dl_thread = None

        self.general_email_db = GeneralEmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.general_email_sync_thread = None
        self.tracked_threads_db = TrackedThreadsDatabase(self.meeting_dir / "Agenda" / "tracked_threads.db")
        self._tracked_thread_dialogs = {}
        self._unified_email_sync_worker = None
        self._live_sync_catchup_scheduled = False
        self.tracked_threads_timer = QTimer(self)
        self.tracked_threads_timer.setInterval(DEFAULT_LIVE_SYNC_INTERVAL_MS)  # 2 minutes live auto-sync
        self.tracked_threads_timer.timeout.connect(
            lambda: self._start_unified_email_sync(sync_mode="incremental", silent=True)
        )

        self._import_queue = []
        self._is_importing_agenda = False
        self._total_import_batch_count = 0
        self._import_accumulated_merged = 0
        self._import_success_files = []
        self._import_failed_files = []

        self.setAcceptDrops(True)
        self.db = TDocsDatabase(self.meeting_dir / "Agenda" / "user_tdocs.db")
        user_data = self.db.get_all()

        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        self.is_sa2 = ('SA2' in wg_name)
        self.is_ran = bool(re.search(r'\bRAN\d?\b', wg_name) or wg_name.startswith("RAN"))
        is_electronic = bool(self.mtg_info.get('is_electronic', 0))
        self.is_sa2_electronic = self.is_sa2 and is_electronic

        main_ftp = self.mtg_info.get("url_key", "")
        if main_ftp and not main_ftp.startswith("http"):
            main_ftp = "https://www.3gpp.org/ftp/" + main_ftp.lstrip('/')
        self.main_ftp_url = main_ftp

        docs_ftp = self.mtg_info.get("docs_folder_url", "")
        if docs_ftp and not docs_ftp.startswith("http"):
            docs_ftp = "https://www.3gpp.org/ftp/" + docs_ftp.lstrip('/')
        self.docs_ftp_url = docs_ftp

        self.revisions_url = self.main_ftp_url.rstrip('/') + '/INBOX/Revisions/' if (
                self.is_sa2_electronic and self.main_ftp_url) else ""

        mtg_icon = "💻" if is_electronic else "🤝"
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')} {mtg_icon}"
        self.setWindowTitle(title)
        self.resize(1400, 750)

        self._row_height_timer = QTimer(self)
        self._row_height_timer.setSingleShot(True)
        self._row_height_timer.setInterval(35)
        self._row_height_timer.timeout.connect(self._adjust_table_row_heights)
        self._user_resized_cols = False
        self._programmatic_column_resize = False

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        try:
            self._load_filter_settings()
            self._setup_header(main_layout, title, is_electronic, len(tdocs_data))
            self._setup_filters(main_layout)
            self._setup_table(main_layout, tdocs_data, user_data)
            self._setup_cache()
        finally:
            self._restoring_filters = False

        self.drop_overlay = DropOverlayWidget(self)
        self._start_background_wi_enrichment()
        if hasattr(self, "act_auto_refresh_threads") and self.act_auto_refresh_threads.isChecked():
            self.tracked_threads_timer.start()

    def showEvent(self, event):
        super().showEvent(event)
        self._load_tracked_wi_tokens()
        self._update_filter_facets()
        self._schedule_column_fit()
        self._schedule_row_height_adjustment()
        if (
            hasattr(self, "act_auto_refresh_threads")
            and self.act_auto_refresh_threads.isChecked()
            and not getattr(self, "_live_sync_catchup_scheduled", False)
        ):
            self._live_sync_catchup_scheduled = True
            QTimer.singleShot(
                2500,
                lambda: self._start_unified_email_sync(sync_mode="incremental", silent=True)
                if self.isVisible()
                else None,
            )

    def _setup_cache(self):
        agenda_dir = self.meeting_dir / "Agenda"
        if self.is_sa2:
            local_agenda = agenda_dir / "TdocsByAgenda.htm"
            if local_agenda.exists():
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(local_agenda))
                if agenda_data:
                    self.model.merge_agenda_data(agenda_data)
                    self._sync_tracked_threads_to_table()
                    self._refresh_comboboxes()

        if getattr(self, "is_ran", False):
            from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
            cached_chair_agenda, cached_agreements = ChairNotesAgreementsParser.load_cached_chair_notes_data(self.meeting_dir)
            if cached_chair_agenda:
                self.model.merge_agenda_data(cached_chair_agenda)
                self._refresh_comboboxes()
            if cached_agreements:
                self.chair_agreements = cached_agreements
                self.model.set_chair_agreements(cached_agreements)
            self._update_agreements_counter()
        else:
            if hasattr(self, "btn_meeting_agreements"):
                self.btn_meeting_agreements.setVisible(False)
            if hasattr(self, "btn_agreements_filter"):
                self.btn_agreements_filter.setVisible(False)

        if self.is_sa2_electronic:
            local_revs = agenda_dir / "revisions.json"
            if local_revs.exists():
                try:
                    with open(local_revs, "r", encoding="utf-8") as f:
                        self.model.revisions = json.load(f)
                        self.model.dataChanged.emit(self.model.index(0, 0),
                                                    self.model.index(self.model.rowCount() - 1, 0))
                except Exception:
                    if self.revisions_url:
                        self._refresh_revisions(silent=True)
            else:
                if self.revisions_url:
                    self._refresh_revisions(silent=True)

    def _get_mod_date_str(self):
        try:
            return f"List last updated: {datetime.datetime.fromtimestamp(os.path.getmtime(self.filepath)).strftime('%Y-%m-%d %H:%M')}"
        except Exception:
            return "List last updated: Unknown"

    def locate_tdoc(self, tdoc_id: str):
        """Public API to scroll to and highlight a specific TDoc in the table."""
        self._scroll_to_tdoc(tdoc_id)

    def _update_count_label(self):
        self.count_lbl.setText(f"Showing {self.proxy.rowCount()} of {self.model.rowCount()} TDocs")
        self._schedule_row_height_adjustment()

    def _save_user_data(self, tdoc_id: str, status: str, notes: str, tags: str = None):
        vbar = self.table.verticalScrollBar() if hasattr(self, "table") else None
        prev_scroll = vbar.value() if vbar else 0
        self.db.upsert(tdoc_id, status, notes, tags=tags)
        if str(tdoc_id or "").strip().upper().startswith("THR-") and hasattr(self, "tracked_threads_db"):
            tag_list = None
            if tags is not None:
                tag_list = [t.strip() for t in str(tags).replace(";", ",").split(",") if t.strip()]
            self.tracked_threads_db.update_thread_annotations(
                str(tdoc_id).strip(),
                user_status=status,
                user_notes=notes,
                user_tags=tag_list,
            )
            open_dlg = getattr(self, "_tracked_thread_dialogs", {}).get(str(tdoc_id).strip())
            if open_dlg is not None:
                try:
                    open_dlg._load_data()
                except Exception:
                    pass
        self.model.user_data = self.db.get_all()
        self.model.apply_user_data_refresh()
        self._refresh_comboboxes()
        self._adjust_table_row_heights()
        if vbar is not None:
            vbar.setValue(prev_scroll)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.setGeometry(self.rect())
        if not getattr(self, "_user_resized_cols", False):
            self._apply_scaled_table_columns()
            self._schedule_column_fit()
        self._schedule_row_height_adjustment()

    def _on_table_context_menu(self, pos: QPoint):
        handle_table_context_menu(self, pos)

    def _open_help_dialog(self):
        """Opens the comprehensive TDocs Cockpit Quick Guide & Shortcuts dialog."""
        from modules.meetings.ui.tdocs_help_dialog import TDocsHelpDialog
        dlg = TDocsHelpDialog(self, is_ran=getattr(self, "is_ran", False))
        dlg.exec_()

    def get_session_state(self) -> dict:
        """Persist filter state and return serializable TDocsWindow context."""
        if hasattr(self, "_save_filter_settings"):
            try:
                self._save_filter_settings()
            except Exception:
                pass

        selected_tdoc = ""
        try:
            if hasattr(self, "table") and hasattr(self, "proxy"):
                idx = self.table.currentIndex()
                if idx.isValid():
                    selected_tdoc = str(self.proxy.data(self.proxy.index(idx.row(), 1), Qt.UserRole) or "")
        except Exception:
            pass

        return {
            "filepath": str(self.filepath or ""),
            "mtg_info": dict(self.mtg_info) if isinstance(self.mtg_info, dict) else {},
            "selected_tdoc": selected_tdoc,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore selected TDoc row if present in the filtered view."""
        if not isinstance(state, dict):
            return
        target = str(state.get("selected_tdoc") or "").strip()
        if not target or not hasattr(self, "_find_proxy_row"):
            return
        row = self._find_proxy_row(target)
        if row is not None and hasattr(self, "table") and hasattr(self, "proxy"):
            idx = self.proxy.index(row, 1)
            if idx.isValid():
                self.table.selectRow(row)
                self.table.setCurrentIndex(idx)
                self.table.scrollTo(idx, QTableView.PositionAtCenter)

    def closeEvent(self, event):
        if hasattr(self, 'search_timer'):
            if self.search_timer.isActive() and hasattr(self, 'proxy') and hasattr(self, 'search_input'):
                self.search_timer.stop()
                self.proxy.setGlobalFilter(self.search_input.text())
            else:
                self.search_timer.stop()
        if hasattr(self, '_save_filter_settings'):
            self._save_filter_settings()
        if hasattr(self, 'routing_timer'):
            self.routing_timer.stop()
        if hasattr(self, 'tracked_threads_timer'):
            self.tracked_threads_timer.stop()
        for w_attr in ("_tracked_threads_sync_worker", "_unified_email_sync_worker"):
            w = getattr(self, w_attr, None)
            if w and w.isRunning():
                try:
                    w.request_cancel()
                    w.wait(1500)
                except Exception:
                    pass
        if hasattr(self, 'search_timer'):
            self.search_timer.stop()

        if getattr(self, 'email_window', None):
            try:
                self.email_window.close()
            except Exception:
                pass

        if getattr(self, '_session_plan_window', None):
            try:
                self._session_plan_window.close()
            except Exception:
                pass
            self._session_plan_window = None

        if getattr(self, '_inbox_dialog', None):
            try:
                self._inbox_dialog.close()
            except Exception:
                pass
            self._inbox_dialog = None

        if getattr(self, '_agreement_tracker_dlg', None):
            try:
                self._agreement_tracker_dlg.close()
            except Exception:
                pass
            self._agreement_tracker_dlg = None

        for dlg in list(getattr(self, '_email_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_email_dialogs'):
            self._email_dialogs.clear()

        for dlg in list(getattr(self, '_triage_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_triage_dialogs'):
            self._triage_dialogs.clear()

        for dlg in list(getattr(self, '_tracked_thread_dialogs', {}).values()):
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        if hasattr(self, '_tracked_thread_dialogs'):
            self._tracked_thread_dialogs.clear()

        all_threads = [
            getattr(self, 'rev_thread', None),
            getattr(self, 'agenda_dl_thread', None),
            getattr(self, 'general_email_sync_thread', None),
            getattr(self, 'dl_thread', None),
            getattr(self, 'package_export_thread', None),
            getattr(self, 'package_excel_exporter_thread', None),
            getattr(self, 'export_thread', None),
            getattr(self, 'excel_exporter_thread', None),
            getattr(self, 'stats_thread', None),
            getattr(self, 'agenda_thread', None),
            getattr(self, 'llm_thread', None),
            getattr(self, 'word_import_thread', None),
            getattr(self, 'chair_notes_thread', None),
            getattr(self, 'cacher_thread', None),
            getattr(self, 'cli_md_export_thread', None),
        ]
        if hasattr(self, 'active_threads'):
            all_threads.extend(self.active_threads.values())

        for th in filter(None, all_threads):
            if hasattr(th, 'is_cancelled'):
                th.is_cancelled = True
            if th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(100)

        super().closeEvent(event)
