# --- File: src/modules/meetings/ui/tdocs/window/mixins/email_sync.py ---
"""
Unified email synchronization, Live Auto-Sync (2m), and Emails toolbar menu builder for TDocsWindow.
"""

from __future__ import annotations

import logging
from PyQt5.QtCore import QSettings, QTimer
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QMenu, QMessageBox, QToolTip

from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.sync_watermark import (
    load_meeting_live_sync_enabled,
    save_meeting_live_sync_enabled,
)
from modules.emails.core.tracked_threads_workers import InboxOutlookSyncWorker
from modules.emails.ui.general_email_dialog import load_wg_email_config


class EmailSyncMixin:
    """Builds the Emails dropdown menu and runs unified incremental/full/live Outlook syncs."""

    def _build_email_menu(self) -> QMenu:
        email_menu = QMenu(self)
        email_menu.setToolTipsVisible(True)
        act = email_menu.addAction("📬 Unread && Flagged Email Inbox...", self._open_meeting_email_inbox)
        act.setToolTip(
            "Open the unified Email Triage Cockpit combining all unread, flagged, and recent emails across Tracked Threads and TDoc Emails."
        )
        act_brief = email_menu.addAction("🤖 AI Email Briefings...", lambda: self._open_email_briefings_dialog())
        act_brief.setToolTip(
            "Open Per-AI cross-thinking email briefings and saved selection summaries for this meeting."
        )
        email_menu.addSeparator()
        act = email_menu.addAction(
            "⚡ Sync New Emails Now",
            lambda: self._start_unified_email_sync(sync_mode="incremental", silent=False),
        )
        act.setToolTip(
            "Fast incremental Outlook sync across Tracked Threads and TDoc Emails using MAPI Table.Restrict() and the persisted watermark (covers app-close / reboot gaps automatically)."
        )
        self.act_auto_refresh_threads = email_menu.addAction("🟢 Live Auto-Sync (Every 2m)")
        self.act_auto_refresh_threads.setCheckable(True)
        self.act_auto_refresh_threads.setToolTip(
            "Continuous eMeeting Update Mode: silently checks Outlook every 2 minutes for new Tracked Thread replies and TDoc emails without rescanning old history."
        )
        live_store = self._live_sync_store() if hasattr(self, "_live_sync_store") else None
        live_enabled = (
            load_meeting_live_sync_enabled(self._get_meeting_id_label(), settings=live_store)
            if hasattr(self, "_get_meeting_id_label")
            else False
        )
        self.act_auto_refresh_threads.setChecked(bool(live_enabled))
        self.act_auto_refresh_threads.toggled.connect(self._on_toggle_auto_refresh_threads)
        act = email_menu.addAction("📅 Full Meeting Sync (Date Range)...", self._on_sync_general_emails)
        act.setToolTip(
            "Choose an explicit date window and buffer to scan all configured Outlook folders and Tracked Threads."
        )
        email_menu.addSeparator()
        act = email_menu.addAction("📌 Pin Email Thread from Outlook Folder...", self._on_pin_thread_from_outlook)
        act.setToolTip("Browse and pin email discussion threads from your configured WG Outlook folders.")
        act = email_menu.addAction("📥 Pin Email Thread from .msg File...", self._on_pin_thread_from_msg_dialog)
        act.setToolTip(
            "Select a local .msg file, locate its full thread in configured Outlook folders, and pin it to the meeting table."
        )
        act = email_menu.addAction("⚙️ Configure Outlook Folders...", self._on_configure_email_folders)
        act.setToolTip("Choose which Outlook folders are searched when syncing TDoc-related emails.")
        act = email_menu.addAction("✔️ Mark All Emails as Read", self._on_mark_all_emails_read)
        act.setToolTip("Mark all locally indexed meeting emails as read.")
        email_menu.addSeparator()
        act = email_menu.addAction("📊 Generate Email Statistics Report...", self._on_generate_meeting_email_statistics)
        act.setToolTip(
            "Generate an interactive HTML Email Analytics Dashboard across all indexed meeting emails and threads."
        )
        email_menu.addSeparator()
        act = email_menu.addAction("🗑️ Wipe Generic Emails Database...", self._on_wipe_generic_emails)
        act.setToolTip("Delete locally indexed generic meeting emails. They can be rebuilt by syncing Outlook again.")
        return email_menu

    def _live_sync_store(self):
        fs = getattr(self, "_filter_settings", None)
        if fs is not None and fs.format() == QSettings.IniFormat:
            return fs
        return None

    def _on_inbox_live_sync_toggled(self, enabled: bool):
        state = bool(enabled)
        save_meeting_live_sync_enabled(self._get_meeting_id_label(), state, settings=self._live_sync_store())
        self._syncing_from_inbox_signal = True
        try:
            act = getattr(self, "act_auto_refresh_threads", None)
            if act is not None and act.isChecked() != state:
                act.setChecked(state)
            timer = getattr(self, "tracked_threads_timer", None)
            if timer is not None:
                if state and not timer.isActive():
                    timer.start()
                elif not state:
                    timer.stop()
        finally:
            self._syncing_from_inbox_signal = False

    def _start_unified_email_sync(
        self,
        *,
        sync_mode: str = "incremental",
        start_date: str = "",
        end_date: str = "",
        days_buffer: int = 3,
        silent: bool = False,
    ):
        """Runs unified async Outlook sync across Tracked Threads and TDoc Emails via InboxOutlookSyncWorker."""
        inbox = getattr(self, "_inbox_dialog", None)
        if silent and inbox is not None:
            try:
                if inbox.isVisible() and (
                    inbox.is_live_sync_enabled()
                    or (getattr(inbox, "_sync_worker", None) and inbox._sync_worker.isRunning())
                ):
                    return
            except RuntimeError:
                self._inbox_dialog = None

        existing = getattr(self, "_unified_email_sync_worker", None)
        if existing is not None and existing.isRunning():
            if not silent:
                notify_info(self, "Email Sync Running", "Outlook email sync is already in progress.")
            return

        tdb = getattr(self, "tracked_threads_db", None)
        gdb = getattr(self, "general_email_db", None)
        if tdb is None and gdb is None:
            return

        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        folders = load_wg_email_config(wg) if wg else []
        meeting_id = self._get_meeting_id_label()
        has_threads = bool(tdb.get_threads_for_meeting(meeting_id)) if tdb is not None else False
        if not folders and not has_threads:
            if not silent:
                if (
                    QMessageBox.question(
                        self,
                        "No Folders Configured",
                        f"No Outlook folders configured for {wg}.\nConfigure folders now?",
                        QMessageBox.Yes | QMessageBox.No,
                    )
                    == QMessageBox.Yes
                ):
                    self._on_configure_email_folders()
            return

        self._unified_sync_silent = bool(silent)
        if not silent:
            if hasattr(self, "thread_sync_bar"):
                self.thread_sync_bar.start_busy("Syncing emails & tracked threads...", total=0, remaining=[])
            if hasattr(self, "email_btn"):
                if not hasattr(self, "_default_email_btn_tooltip"):
                    self._default_email_btn_tooltip = self.email_btn.toolTip()
                self.email_btn.setText("⏳ Syncing... ▾")

        worker = InboxOutlookSyncWorker(
            getattr(tdb, "db_path", None) if tdb is not None else None,
            getattr(gdb, "db_path", None) if gdb is not None else None,
            meeting_id=meeting_id,
            wg=wg,
            start_date=str(start_date or self.mtg_info.get("start_date") or ""),
            end_date=str(end_date or self.mtg_info.get("end_date") or ""),
            parent=self,
            sync_mode=sync_mode,
            days_buffer=days_buffer,
            silent=silent,
        )
        worker.progress_text.connect(self._on_unified_email_sync_progress)
        worker.finished_sync.connect(self._on_unified_email_sync_finished)
        self._unified_email_sync_worker = worker
        worker.start()

    def _on_unified_email_sync_progress(self, msg: str):
        if getattr(self, "_unified_sync_silent", False):
            return
        if hasattr(self, "email_btn"):
            self.email_btn.setToolTip(f"Syncing Outlook emails: {msg}")
        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.update_progress(0, 0, msg, [])

    def _on_unified_email_sync_finished(self, ok: bool, new_thread_msgs: int, general_synced: int, err: str):
        silent = bool(getattr(self, "_unified_sync_silent", False))
        self._unified_email_sync_worker = None
        if hasattr(self, "general_email_db") and self.general_email_db is not None and hasattr(self, "model"):
            try:
                counts = self.general_email_db.get_email_counts_per_tdoc()
                self.model.set_email_counts(counts)
            except Exception as exc:
                logging.debug("Could not refresh TDoc email counts after sync: %s", exc)
        self._on_tracked_thread_changed()

        if not ok:
            if hasattr(self, "thread_sync_bar") and not silent:
                self.thread_sync_bar.show_error(err or "Outlook email sync interrupted.")
            if not silent:
                notify_toast(
                    self,
                    "Outlook Email Sync Warning",
                    err or "Outlook email sync did not complete cleanly.",
                    variant="warning",
                )
            return

        if hasattr(self, "thread_sync_bar") and not silent:
            self.thread_sync_bar.finish_ok()
        if new_thread_msgs > 0 or general_synced > 0:
            notify_success(
                self,
                "Emails Synced with Outlook",
                f"Synced {new_thread_msgs} tracked thread update(s) and {general_synced} TDoc email update(s).",
            )
            if silent:
                QToolTip.showText(
                    QCursor.pos(),
                    f"📧 {new_thread_msgs} tracked thread / {general_synced} TDoc email update(s) synced!",
                    self,
                )
        elif not silent:
            notify_info(
                self,
                "Emails Up to Date",
                "No new matching emails or status updates found in Outlook.",
            )

    def _on_refresh_tracked_threads(self, silent: bool = False):
        self._start_unified_email_sync(sync_mode="incremental", silent=silent)

    def _on_cancel_tracked_threads_sync(self):
        for w_attr in ("_tracked_threads_sync_worker", "_unified_email_sync_worker"):
            w = getattr(self, w_attr, None)
            if w and w.isRunning():
                w.request_cancel()

    def _on_tracked_threads_sync_progress(self, completed: int, total: int, current_title: str, remaining: list):
        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.update_progress(completed, total, current_title, remaining)
        if hasattr(self, "email_btn"):
            self.email_btn.setText(f"⏳ Emails ({completed}/{total}) ▾")
            rem_txt = "\n• ".join(remaining[:8]) if remaining else "(none)"
            self.email_btn.setToolTip(
                f"Syncing thread ({completed}/{total}): {current_title}\n\nRemaining:\n• {rem_txt}"
            )

    def _on_tracked_threads_sync_finished(self, ok: bool, updates: dict, err: str):
        silent = bool(getattr(self, "_sync_silent_mode", False))
        if hasattr(self, "email_btn") and hasattr(self, "_default_email_btn_tooltip"):
            self.email_btn.setToolTip(self._default_email_btn_tooltip)
        self._on_tracked_thread_changed()
        if not ok:
            if hasattr(self, "thread_sync_bar"):
                self.thread_sync_bar.show_error(err or "Tracked thread sync interrupted.")
            notify_toast(
                self,
                "Tracked Thread Sync Warning",
                err or "Tracked thread sync did not complete cleanly.",
                variant="warning",
            )
            return

        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.finish_ok()
        new_total = sum((updates or {}).values())
        if new_total > 0:
            notify_success(
                self,
                "Tracked Threads Refreshed",
                f"Found {new_total} new message(s) across {len(updates)} tracked thread(s).",
            )
            if silent:
                QToolTip.showText(
                    QCursor.pos(),
                    f"📧 {new_total} new reply/replies across {len(updates)} tracked email thread(s)!",
                    self,
                )
        elif not silent:
            notify_info(
                self,
                "Tracked Threads Up-to-Date",
                "No new replies found in configured Outlook folders.",
            )

    def _on_toggle_auto_refresh_threads(self, enabled: bool):
        state = bool(enabled)
        save_meeting_live_sync_enabled(self._get_meeting_id_label(), state, settings=self._live_sync_store())
        if hasattr(self, "tracked_threads_timer"):
            if state:
                if not self.tracked_threads_timer.isActive():
                    self.tracked_threads_timer.start()
            else:
                self.tracked_threads_timer.stop()
        inbox = getattr(self, "_inbox_dialog", None)
        if inbox is not None:
            try:
                inbox.set_live_sync_enabled(state, emit_signal=False)
            except RuntimeError:
                self._inbox_dialog = None
        if state and self.isVisible() and not getattr(self, "_syncing_from_inbox_signal", False):
            inbox_active = False
            try:
                inbox_active = bool(inbox is not None and inbox.isVisible())
            except RuntimeError:
                inbox_active = False
            if not inbox_active:
                self._start_unified_email_sync(sync_mode="incremental", silent=True)
            elif inbox is not None:
                QTimer.singleShot(100, inbox._on_live_sync_timer_tick)
