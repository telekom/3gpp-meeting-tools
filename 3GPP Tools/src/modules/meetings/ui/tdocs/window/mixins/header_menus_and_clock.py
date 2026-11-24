# --- File: src/modules/meetings/ui/tdocs/window/mixins/header_menus_and_clock.py ---
"""Resources menu population, instant fetch/compare, routing indicator, and venue clock."""

import datetime
import logging
import re
import webbrowser

from PyQt5.QtWidgets import QApplication, QMessageBox

from core.network.network_state import NetworkState
from core.ui.theme import C
from core.ui.world_clock_dialog import WorldClockDialog
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.terminal_opener import open_terminal as _open_terminal
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.meetings_menu import add_meeting_portal_actions


class HeaderMenusAndClockMixin:
    """Dynamic resources menu, instant TDoc fetch/compare, routing badge, and venue clock."""

    def _populate_resources_menu(self):
        self.folder_menu.clear()
        self.folder_menu.addAction("📁 Local: Meeting Folder", self._open_meeting_folder)
        if self.meeting_dir.is_dir():
            act = self.folder_menu.addAction(
                "💻 Open Terminal in Meeting Cache",
                lambda: _open_terminal(self.meeting_dir, self),
            )
            act.setToolTip(
                "Open Windows Terminal with this meeting's local cache folder as its working directory."
            )
        export_dir = self.meeting_dir / "Export"
        if export_dir.is_dir():
            act_exp = self.folder_menu.addAction("📂 Local: Export Folder", self._open_export_folder)
            act_exp.setToolTip("Open this meeting's local Export directory containing packages and Markdown dossiers.")
        local_agenda_csv = self.agenda_dir / "agenda.csv"
        if local_agenda_csv.exists():
            self.folder_menu.addAction("📋 Local: agenda.csv", lambda: _open_folder(local_agenda_csv))
        session_plans_dir = self.meeting_dir / "Agenda" / "SessionPlans"
        if session_plans_dir.is_dir():
            self.folder_menu.addAction(
                "🗓️ Local: Session Plans Folder", lambda: _open_folder(session_plans_dir)
            )
        chair_notes_dir = self.meeting_dir / "Agenda" / "ChairNotes"
        if chair_notes_dir.is_dir():
            local_docs = []
            try:
                local_docs = [
                    f for f in chair_notes_dir.iterdir()
                    if f.is_file()
                    and f.suffix.lower() in (".docx", ".doc")
                    and not f.name.startswith(("~$", "."))
                ]
            except Exception:
                local_docs = []

            has_any_file = False
            try:
                has_any_file = any(chair_notes_dir.iterdir())
            except Exception:
                pass

            if has_any_file:
                self.folder_menu.addAction("📂 Local: ChairNotes Folder", self._open_chair_notes_folder)

            if local_docs:
                if len(local_docs) == 1:
                    doc = local_docs[0]
                    act = self.folder_menu.addAction(
                        f"📝 Open Notes: {doc.name}", lambda d=doc: self._open_doc_in_word(d)
                    )
                    act.setToolTip(f"Open {doc.name} in Microsoft Word.")
                else:
                    notes_sub = self.folder_menu.addMenu("📝 Open Chairman's Notes (.docx)")
                    notes_sub.setToolTipsVisible(True)
                    sorted_docs = sorted(local_docs, key=lambda f: f.stat().st_mtime, reverse=True)
                    for doc in sorted_docs:
                        act = notes_sub.addAction(doc.name, lambda d=doc: self._open_doc_in_word(d))
                        act.setToolTip(f"Open {doc.name} in Microsoft Word.")
        if self.is_sa2:
            self.folder_menu.addAction("📄 Local: TdocsByAgenda.htm", self._open_agenda_file)

        self.folder_menu.addSeparator()
        self.folder_menu.addAction("⚠️ View Unmatched Companies", self._show_unmatched_companies)
        act_resource_md_export = self.folder_menu.addAction("📝 Export Markdown Reports", self._export_reports)
        act_resource_md_export.setToolTip(
            "Generate Markdown summary reports for the meeting from the currently loaded TDoc data."
        )
        act_resource_md_export.setStatusTip(act_resource_md_export.toolTip())

        if NetworkState.get_instance().is_local_active():
            self.folder_menu.addSeparator()
            wg_name = self.mtg_info.get("wg_name", "").upper()
            local_base = URLRouter._get_local_server_base(wg_name)
            self.folder_menu.addAction(
                "🟢 Local Server: Main WG Folder", lambda u=local_base: webbrowser.open(u)
            )
            self.folder_menu.addAction(
                "🟢 Local Server: Docs Folder", lambda u=f"{local_base}/Docs": webbrowser.open(u)
            )
            self.folder_menu.addAction(
                "🟢 Local Server: Inbox Folder", lambda u=f"{local_base}/Inbox": webbrowser.open(u)
            )
            if self.is_sa2:
                self.folder_menu.addAction(
                    "🟢 Local Server: Revisions Folder",
                    lambda u=f"{local_base}/Inbox/Revisions": webbrowser.open(u),
                )
                self.folder_menu.addAction(
                    "🟢 Local Server: TdocsByAgenda.htm",
                    lambda u=f"{local_base}/TdocsByAgenda.htm": webbrowser.open(u),
                )
            self.folder_menu.addAction(
                "🟢 Local Server: Home (10.10.10.10)", lambda: webbrowser.open("http://10.10.10.10/")
            )

        has_ftp = bool(self.main_ftp_url or self.docs_ftp_url or self.revisions_url)
        if has_ftp:
            self.folder_menu.addSeparator()
            if self.main_ftp_url:
                self.folder_menu.addAction("🌐 Web FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
            if self.docs_ftp_url:
                self.folder_menu.addAction("🌐 Web FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
                act_cache = self.folder_menu.addAction("📥 Cache TDocs (Docs/)", self._cache_all_tdocs)
                act_cache.setToolTip(
                    "Cache all official TDoc archives from the 3GPP Docs directory into the local meeting cache."
                )
            if self.revisions_url:
                self.folder_menu.addAction(
                    "🌐 Web FTP: Revisions Folder", lambda: webbrowser.open(self.revisions_url)
                )

        if self.mtg_info.get("mtg_id"):
            self.folder_menu.addSeparator()
            add_meeting_portal_actions(self.folder_menu, self.mtg_info)

    def _trigger_download_thread(
        self,
        base_tdoc: str,
        target_filename: str,
        legacy_url: str = None,
        is_silent_compare: bool = False,
        auto_compare_active_word: bool = False,
        sync_table_on_success: bool = False,
        send_email_on_finish: bool = False,
    ):
        self.model.set_loading(base_tdoc, True)
        is_active = self.mtg_info.get("is_active_sync", False)

        url_list = URLRouter.build_priority_url_list(
            self.mtg_info.get("wg_name", ""),
            self.mtg_info.get("folder_name") or self.mtg_info.get("meeting_number", ""),
            self.main_ftp_url,
            is_active,
            target_filename=target_filename,
        )

        thread = TDocActionThread(
            base_tdoc,
            target_filename,
            url_list,
            self.meeting_dir,
            open_file=not (is_silent_compare or send_email_on_finish),
        )
        thread.is_silent_compare = is_silent_compare
        thread.auto_compare_active_word = auto_compare_active_word
        thread.sync_table_on_success = sync_table_on_success
        thread.send_email_on_finish = send_email_on_finish
        thread.target_filename = target_filename
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))

        self.active_threads[base_tdoc] = thread
        thread.start()

    def _compare_tdoc_with_active_word(
        self,
        base_tdoc: str,
        target_filename: str,
        sync_table_on_success: bool = False,
    ):
        """Compare the currently active Microsoft Word document (Slot A) against a TDoc (Slot B)."""
        mgr = ComparisonManager.get_instance()
        active_doc = mgr.stage_active_word_document(slot="A")
        if not active_doc:
            QMessageBox.warning(
                self,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and the document you wish to compare is open on your screen.",
            )
            return

        self._trigger_download_thread(
            base_tdoc,
            target_filename,
            is_silent_compare=True,
            auto_compare_active_word=True,
            sync_table_on_success=sync_table_on_success,
        )

    def _get_tdoc_prefix(self):
        first_tdoc = self.mtg_info.get("first_tdoc", "")
        if first_tdoc:
            match = re.match(r"^([A-Z0-9]+-\d{2})", first_tdoc.upper())
            if match:
                return match.group(1)

        wg = self.mtg_info.get("wg_name", "").upper()
        start_date = self.mtg_info.get("start_date", "")
        year_str = start_date[2:4] if len(start_date) >= 4 else datetime.datetime.now().strftime("%y")

        wg_map = {
            "SA1": "S1", "SA2": "S2", "SA3": "S3", "SA4": "S4", "SA5": "S5", "SA6": "S6", "SA": "SP",
            "RAN1": "R1", "RAN2": "R2", "RAN3": "R3", "RAN4": "R4", "RAN5": "R5", "RAN6": "R6", "RAN": "RP",
            "CT1": "C1", "CT3": "C3", "CT4": "C4", "CT6": "C6", "CT": "CP",
        }
        prefix = wg_map.get(wg, wg)
        return f"{prefix}-{year_str}"

    def _update_routing_indicator(self):
        net_state = NetworkState.get_instance()
        is_active = self.mtg_info.get("is_active_sync", False)

        if net_state.is_local_active():
            self.routing_indicator.setText("🟢 Local Server")
            self.routing_indicator.setStyleSheet(
                f"color: {C.WG_RAN_TEXT}; background-color: {C.EMERALD_SOFT}; "
                "font-weight: bold; padding: 2px 8px; border-radius: 4px;"
            )
            self.routing_indicator.setToolTip("Downloads are routed through the high-speed local 10.10.10.10 network.")
        elif is_active:
            self.routing_indicator.setText("🔵 SYNC Folder")
            self.routing_indicator.setStyleSheet(
                f"color: {C.INFO_DARK}; background-color: {C.INFO_SOFT_STRONG}; "
                "font-weight: bold; padding: 2px 8px; border-radius: 4px;"
            )
            self.routing_indicator.setToolTip("Downloads are routed through the active meeting SYNC folder.")
        else:
            self.routing_indicator.setText("🌐 Standard Web")
            self.routing_indicator.setStyleSheet(
                f"color: {C.TEXT_BODY}; background-color: {C.BG_CANVAS}; "
                "font-weight: bold; padding: 2px 8px; border-radius: 4px;"
            )
            self.routing_indicator.setToolTip("Downloads are routed through the standard 3GPP web archive.")
        self._update_venue_clock()

    def _update_venue_clock(self):
        if not hasattr(self, "venue_clock_btn") or not hasattr(self, "_clock_service"):
            return
        try:
            state = self._clock_service.resolve_single_meeting(
                self.mtg_info, meeting_dir=self.meeting_dir
            )
            if state.venue_dt is not None and state.tz_offset_minutes is not None:
                is_elec = bool(self.mtg_info.get("is_electronic", 0)) or state.location == "Online (UTC)"
                if is_elec:
                    self.venue_clock_btn.setText(f"🌐 {state.venue_dt:%H:%M} UTC (eMeeting)")
                else:
                    off_str = format_utc_offset(state.tz_offset_minutes)
                    if off_str.endswith(":00"):
                        off_str = off_str[:-3]
                    day_tag = f" ({state.day_diff_vs_laptop:+d}d)" if state.day_diff_vs_laptop != 0 else ""
                    self.venue_clock_btn.setText(
                        f"📍 {state.location} {state.venue_dt:%H:%M}{day_tag} ({off_str})"
                    )
                self.venue_clock_btn.setToolTip(
                    f"Meeting Venue: {state.location}\n"
                    f"Venue Time: {state.venue_dt:%a %d %b %Y, %H:%M:%S}\n"
                    f"Timezone: {state.tz_label}\n"
                    f"Source: {state.source}\n\n"
                    f"Click to manage venue timezone, sync from web, or set manual override"
                )
            else:
                self.venue_clock_btn.setText(f"📍 {state.location or 'Venue'} (--:--)")
                self.venue_clock_btn.setToolTip(
                    "Venue timezone unresolved. Click to configure or sync from web."
                )
        except Exception as e:
            logging.debug(f"[TDocsWindow] Venue clock update failed: {e}")

    def _open_venue_clock_dialog(self):
        dlg = WorldClockDialog(
            self._clock_service,
            focus_meeting=self.mtg_info,
            focus_meeting_dir=self.meeting_dir,
            parent=self,
        )

        def _sync_all_clocks():
            self._update_venue_clock()
            for top in QApplication.topLevelWidgets():
                sb = getattr(top, "status_bar", None)
                if sb is not None and hasattr(sb, "refresh_world_clock"):
                    sb.refresh_world_clock()

        dlg.settings_changed.connect(_sync_all_clocks)
        dlg.exec_()
        _sync_all_clocks()

    def _on_instant_fetch(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r"^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$", tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._trigger_download_thread(
            base_tdoc,
            target_filename,
            legacy_url=None,
            is_silent_compare=False,
            sync_table_on_success=True,
        )

    def _on_instant_compare_active_word(self):
        tdoc_str = self.instant_fetch_input.text().strip()
        match = re.match(r"^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$", tdoc_str, re.IGNORECASE)
        if not match:
            QMessageBox.warning(self, "Invalid TDoc", "Please enter a valid TDoc number (e.g., S2-261234).")
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        self._compare_tdoc_with_active_word(base_tdoc, target_filename, sync_table_on_success=True)
