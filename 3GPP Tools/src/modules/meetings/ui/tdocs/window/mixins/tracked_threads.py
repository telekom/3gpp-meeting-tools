# --- File: src/modules/meetings/ui/tdocs/window/mixins/tracked_threads.py ---
"""
Tracked e-mail threads: pinning (Outlook/.msg), sync worker, inbox dialog, table sync.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
import re
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QMessageBox,
    QToolTip,
    QDialog,
)
from modules.meetings.ui.tdocs_models import natural_sort_key
from modules.emails.core.tracked_threads_service import pin_thread_from_msg, pin_thread_from_outlook_selection
from modules.emails.core.tracked_threads_workers import TrackedThreadsRefreshWorker
from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog
from modules.emails.ui.tracked_thread_dialog import PinFromOutlookDialog, TrackedThreadDialog
from core.ui.toast import (
    notify_info,
    notify_success,
    notify_toast,
)
from modules.emails.ui.general_email_dialog import load_wg_email_config


class TrackedThreadsMixin:
    """Tracked e-mail threads: pinning (Outlook/.msg), sync worker, inbox dialog, table sync."""


    def _get_meeting_id_label(self) -> str:
        wg = str(self.mtg_info.get("wg_name", "")).strip().upper()
        num = str(self.mtg_info.get("meeting_number", "")).strip()
        return f"{wg}#{num}" if (wg and num) else (wg or num or "MEETING")

    def _expand_tdoc_family_with_merges(self, tdocs: list) -> list:
        """Expands TDoc IDs across revision chains and merged-to/replaced-by contributions."""
        if not tdocs or not hasattr(self, "model"):
            return list(tdocs or [])
        tdoc_map = getattr(self.model, "_tdoc_map", {}) or {}
        seen = []
        seen_set = set()
        queue = [str(t).strip().upper() for t in tdocs if str(t).strip()]
        tdoc_token_re = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{5,8}(?:r\d{1,2}[a-z]?)?\b", re.IGNORECASE)

        while queue:
            cur = queue.pop(0)
            if cur in seen_set:
                continue
            seen_set.add(cur)
            seen.append(cur)

            # Revision family from model
            for fam in self.model.get_family_tdocs(cur):
                f_up = str(fam).strip().upper()
                if f_up and f_up not in seen_set:
                    queue.append(f_up)

            # Check merge / replacement hints in row metadata
            row = tdoc_map.get(cur)
            if row:
                for col in ("Merged to", "Merged into", "Replaced by", "Revised to", "Secretary Remarks", "TDoc Status"):
                    val = str(row.get(col, "") or "")
                    if not val:
                        continue
                    if col in ("Secretary Remarks", "TDoc Status") and not any(k in val.lower() for k in ("merg", "revis", "replac", "supersed")):
                        continue
                    for m in tdoc_token_re.findall(val):
                        m_up = m.upper()
                        if m_up not in seen_set:
                            queue.append(m_up)
        return seen

    def _build_thread_autocomplete_suggestions(self):
        """Builds AI, WI, TDoc, Sub-Topic lists and AI -> Sub-Topics mapping for TrackedThreadDialog."""
        ai_items = []
        seen_ais = set()
        for ai_num, item in sorted((self.agenda_map or {}).items(), key=lambda kv: natural_sort_key(kv[0])):
            desc = getattr(item, "description", "") or ""
            label = f"{ai_num} — {desc}" if desc else str(ai_num)
            if ai_num not in seen_ais:
                seen_ais.add(ai_num)
                ai_items.append(label)

        wi_set = set()
        tdoc_items = []
        seen_tdocs = set()
        sub_items = []
        seen_subs = set()
        ai_to_sub_items = {}
        for row in (getattr(self.model, "_data", None) or []):
            if row.get("_is_tracked_thread"):
                continue
            ai_val = str(row.get("Agenda Item", "") or "").strip()
            valid_ai = bool(ai_val and ai_val not in ("N/A", "-", "None"))
            if valid_ai and ai_val not in seen_ais:
                seen_ais.add(ai_val)
                ai_desc = str(row.get("Agenda item description", "") or "").strip()
                ai_items.append(f"{ai_val} — {ai_desc}" if ai_desc else ai_val)
            sub_val = str(row.get("Sub-Item", "") or "").strip()
            if sub_val:
                if valid_ai:
                    bucket = ai_to_sub_items.setdefault(ai_val, [])
                    if sub_val not in bucket:
                        bucket.append(sub_val)
                if sub_val not in seen_subs:
                    seen_subs.add(sub_val)
                    sub_items.append(f"{sub_val} — AI {ai_val}" if valid_ai else sub_val)
            for wi in self.model._split_related_wis(row.get("Related WIs", "")):
                if wi and wi.lower() not in ("(no wi)", "none", "nan", "-"):
                    wi_set.add(wi)
            tid = str(row.get("TDoc", "") or "").strip().upper()
            if tid and tid not in seen_tdocs:
                seen_tdocs.add(tid)
                title = str(row.get("Title", "") or "").strip()
                tdoc_items.append(f"{tid} — {title[:70]}" if title else tid)
        return ai_items, sorted(wi_set), tdoc_items, sub_items, ai_to_sub_items

    def _sync_tracked_threads_to_table(self):
        if not hasattr(self, "tracked_threads_db") or not hasattr(self, "model"):
            return
        try:
            threads = self.tracked_threads_db.get_threads_for_meeting(self._get_meeting_id_label())
            tracked_unread_ids = set()
            for t in threads:
                cur_tdocs = t.get("linked_tdocs") or []
                if cur_tdocs:
                    expanded = self._expand_tdoc_family_with_merges(cur_tdocs)
                    if [x.upper() for x in expanded] != [x.upper() for x in cur_tdocs]:
                        t["linked_tdocs"] = expanded
                        self.tracked_threads_db.update_thread_links(t["thread_id"], linked_tdocs=expanded)
                msgs = self.tracked_threads_db.get_thread_messages(t["thread_id"])
                for m in msgs:
                    if int(m.get("is_read", 0) or 0) == 0 and m.get("id"):
                        tracked_unread_ids.add(str(m["id"]))
                comps = list(dict.fromkeys(
                    str(m.get("company") or "").strip()
                    for m in msgs
                    if str(m.get("company") or "").strip() and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
                ))
                if comps:
                    t["participants_companies"] = ", ".join(comps)
            self.model.sync_tracked_thread_rows(threads)
            tdoc_unread_ids = set()
            if hasattr(self, "general_email_db") and self.general_email_db is not None:
                for stats in (self.general_email_db.get_email_counts_per_tdoc() or {}).values():
                    tdoc_unread_ids.update(str(eid) for eid in (stats.get("unread_ids") or ()) if eid)
            total_unread = len(tracked_unread_ids | tdoc_unread_ids)
            if hasattr(self, "email_btn"):
                if total_unread > 0:
                    self.email_btn.setText(f"📧 Emails ({total_unread} new) ▾")
                    self.email_btn.setToolTip(
                        f"Email Discussions & Tracked Threads\n"
                        f"{total_unread} unread email(s) total "
                        f"({len(tracked_unread_ids)} in Tracked Threads, {len(tdoc_unread_ids)} in TDoc Emails)."
                    )
                else:
                    self.email_btn.setText("📧 Emails ▾")
                    self.email_btn.setToolTip("Email Discussions & Tracked Threads (no unread emails).")
        except Exception as e:
            logging.error(f"Error syncing tracked email threads to table: {e}")

    def _open_tracked_thread_dialog(self, thread_id: str):
        tid = str(thread_id or "").strip()
        if not tid:
            return
        if tid in self._tracked_thread_dialogs:
            dlg = self._tracked_thread_dialogs[tid]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return
        ai_sugg, wi_sugg, tdoc_sugg, sub_sugg, ai_to_sub_items = self._build_thread_autocomplete_suggestions()
        all_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        dlg = TrackedThreadDialog(
            self.tracked_threads_db,
            tid,
            parent=None,
            ai_suggestions=ai_sugg,
            wi_suggestions=wi_sugg,
            tdoc_suggestions=tdoc_sugg,
            tdoc_family_expander=self._expand_tdoc_family_with_merges,
            sub_item_suggestions=sub_sugg,
            all_existing_tags=all_tags,
            annotations_save_callback=self._save_user_data,
            ai_to_sub_items=ai_to_sub_items,
            open_tdoc_callback=self._open_tdoc_document_directly,
        )
        dlg.thread_changed.connect(self._on_tracked_thread_changed)
        self._tracked_thread_dialogs[tid] = dlg
        dlg.finished.connect(lambda _, k=tid: self._tracked_thread_dialogs.pop(k, None))
        dlg.show()

    def _refresh_inbox_dialog_if_open(self):
        inbox = getattr(self, "_inbox_dialog", None)
        if inbox is None:
            return
        try:
            if inbox.isVisible():
                inbox.reload_from_dbs()
        except RuntimeError:
            self._inbox_dialog = None

    def _on_tracked_thread_changed(self):
        self._sync_tracked_threads_to_table()
        self._refresh_comboboxes()
        self._schedule_row_height_adjustment()
        self._refresh_inbox_dialog_if_open()

    def _build_tdoc_meta_map(self) -> dict:
        """Maps TDoc IDs to their Agenda Item, Sub-Item, and Related WIs for email link inheritance."""
        meta_map = {}
        if not hasattr(self, "model") or not hasattr(self.model, "_data"):
            return meta_map
        for r in self.model._data:
            if r.get("_is_tracked_thread"):
                continue
            td = str(r.get("TDoc") or "").strip().upper()
            if not td:
                continue
            ai = str(r.get("Agenda Item") or "").strip()
            if ai in ("N/A", "-", "None"):
                ai = ""
            sub = str(r.get("Sub-Item") or "").strip()
            if sub in ("N/A", "-", "None"):
                sub = ""
            wis = [
                w for w in self.model._split_related_wis(r.get("Related WIs", ""))
                if w and w.lower() not in ("(no wi)", "none", "nan", "-")
            ]
            meta_map[td] = {"ai": ai, "sub_item": sub, "wis": wis}
        return meta_map

    def _get_pin_meeting_context(self):
        """Returns (agenda_map, valid_wis) for resolving KI/AI/WI tokens against the active meeting."""
        _, wi_sugg, _, _, _ = self._build_thread_autocomplete_suggestions()
        return getattr(self, "agenda_map", None), wi_sugg

    def _process_dropped_outlook_items(self, mime_data):
        """Handles emails dragged directly from Microsoft Outlook UI (FileGroupDescriptor / RenPrivateMessages)."""
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        meeting_id = self._get_meeting_id_label()
        agenda_map, valid_wis = self._get_pin_meeting_context()
        tdoc_meta_map = self._build_tdoc_meta_map()
        try:
            # First try reading Outlook ActiveExplorer Selection (most reliable for .msg drags in Windows)
            pinned_ids = pin_thread_from_outlook_selection(
                self.tracked_threads_db,
                meeting_id=meeting_id,
                wg=wg,
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=tdoc_meta_map,
            )
            if not pinned_ids and mime_data:
                # Fallback: try extracting FileContents if non-empty
                import struct
                formats = mime_data.formats()
                fgd_fmt = next((f for f in formats if "FileGroupDescriptor" in f), None)
                fc_fmt = next((f for f in formats if "FileContents" in f), None)
                if fgd_fmt and fc_fmt:
                    contents = bytes(mime_data.data(fc_fmt))
                    desc = bytes(mime_data.data(fgd_fmt))
                    if contents and len(desc) >= 4:
                        is_uni = "FileGroupDescriptorW" in fgd_fmt
                        rec_size = 592 if is_uni else 332
                        if len(desc) >= 4 + rec_size:
                            raw_name = desc[4 + 72 : 4 + rec_size]
                            fname = (raw_name.decode("utf-16le", errors="ignore") if is_uni else raw_name.decode("mbcs", errors="ignore")).split("\0")[0].strip()
                            if fname.lower().endswith(".msg"):
                                drop_dir = self.meeting_dir / "Agenda" / "DroppedEmails"
                                drop_dir.mkdir(parents=True, exist_ok=True)
                                tmp_msg = drop_dir / fname
                                tmp_msg.write_bytes(contents)
                                tid = pin_thread_from_msg(
                                    self.tracked_threads_db,
                                    tmp_msg,
                                    meeting_id=meeting_id,
                                    wg=wg,
                                    agenda_map=agenda_map,
                                    valid_wis=valid_wis,
                                    tdoc_meta_map=tdoc_meta_map,
                                )
                                if tid:
                                    pinned_ids.append(tid)
            if pinned_ids:
                self._on_tracked_thread_changed()
                self._open_tracked_thread_dialog(pinned_ids[-1])
            else:
                QMessageBox.warning(
                    self,
                    "Outlook Drag & Drop",
                    "Could not read the dragged email from Microsoft Outlook.\n"
                    "Ensure the email is selected in Classic Outlook or save it as a .msg file.",
                )
        except Exception as e:
            logging.error(f"Error handling Outlook drag & drop: {e}", exc_info=True)
            QMessageBox.warning(self, "Outlook Drag & Drop Failed", f"Failed to process dragged Outlook email:\n{e}")

    def _process_dropped_msg_files(self, msg_files: list):
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        meeting_id = self._get_meeting_id_label()
        agenda_map, valid_wis = self._get_pin_meeting_context()
        tdoc_meta_map = self._build_tdoc_meta_map()
        pinned_ids = []
        try:
            for mp in msg_files:
                tid = pin_thread_from_msg(
                    self.tracked_threads_db,
                    Path(mp),
                    meeting_id=meeting_id,
                    wg=wg,
                    agenda_map=agenda_map,
                    valid_wis=valid_wis,
                    tdoc_meta_map=tdoc_meta_map,
                )
                if tid:
                    pinned_ids.append(tid)
            if pinned_ids:
                self._on_tracked_thread_changed()
                self._open_tracked_thread_dialog(pinned_ids[-1])
            else:
                QMessageBox.warning(
                    self,
                    "Could Not Parse .msg",
                    "Failed to parse the dropped .msg file via Outlook COM. Ensure Microsoft Outlook is available.",
                )
        except Exception as e:
            logging.error(f"Error processing dropped .msg file(s): {e}", exc_info=True)
            QMessageBox.warning(self, "Error Processing .msg", f"Failed to process dropped .msg file:\n{e}")

    def _on_pin_thread_from_msg_dialog(self):
        from PyQt5.QtWidgets import QFileDialog
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Outlook .msg File(s) to Pin Thread", "", "Outlook Message (*.msg)")
        if paths:
            self._process_dropped_msg_files([Path(p) for p in paths])

    def _on_pin_thread_from_outlook(self):
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        folders = load_wg_email_config(wg)
        if not folders:
            if QMessageBox.question(
                self,
                "No Folders Configured",
                f"No Outlook folders configured for {wg}.\nConfigure folders now?",
                QMessageBox.Yes | QMessageBox.No,
            ) == QMessageBox.Yes:
                self._on_configure_email_folders()
            return
        agenda_map, valid_wis = self._get_pin_meeting_context()
        dlg = PinFromOutlookDialog(
            self.tracked_threads_db,
            self._get_meeting_id_label(),
            wg,
            self,
            agenda_map=agenda_map,
            valid_wis=valid_wis,
            meeting_dir=getattr(self, "meeting_dir", None),
            general_email_db=getattr(self, "general_email_db", None),
            valid_tdocs=getattr(self.model, "valid_tdocs", set()),
            mtg_start_date=str(self.mtg_info.get("start_date") or ""),
            mtg_end_date=str(self.mtg_info.get("end_date") or ""),
            tdoc_meta_map=self._build_tdoc_meta_map(),
        )
        if dlg.exec_() == QDialog.Accepted and dlg.pinned_thread_ids:
            self._on_tracked_thread_changed()
            self._open_tracked_thread_dialog(dlg.pinned_thread_ids[-1])

    def _open_meeting_email_inbox(self):
        """Opens the unified Unread & Flagged Email Inbox cockpit across Tracked Threads and TDoc Emails."""
        existing = getattr(self, "_inbox_dialog", None)
        if existing is not None:
            try:
                existing.reload_from_dbs()
                existing.setWindowState(existing.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
                existing.raise_()
                existing.activateWindow()
                return
            except RuntimeError:
                self._inbox_dialog = None

        tdoc_ai_lookup = {}
        if hasattr(self, "model") and hasattr(self.model, "_data"):
            for r in self.model._data:
                if r.get("_is_tracked_thread"):
                    continue
                td = str(r.get("TDoc") or "").strip().upper()
                ai = str(r.get("Agenda Item") or "").strip()
                if td and ai:
                    tdoc_ai_lookup[td] = ai

        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        ai_sugg, wi_sugg, _tdoc_sugg, _sub_sugg, ai_to_sub_items = self._build_thread_autocomplete_suggestions()
        dlg = MeetingEmailInboxDialog(
            tracked_db=getattr(self, "tracked_threads_db", None),
            general_email_db=getattr(self, "general_email_db", None),
            meeting_id=self._get_meeting_id_label(),
            tdoc_ai_lookup=tdoc_ai_lookup,
            open_thread_callback=self._open_tracked_thread_dialog,
            open_tdoc_callback=self._open_tdoc_document_directly,
            parent=None,
            wg=wg,
            ai_suggestions=ai_sugg,
            wi_suggestions=wi_sugg,
            ai_to_sub_items=ai_to_sub_items,
            tdoc_meta_map=self._build_tdoc_meta_map(),
            agenda_map=getattr(self, "agenda_map", None),
            meeting_start=str(self.mtg_info.get("start_date", "") or ""),
            meeting_end=str(self.mtg_info.get("end_date", "") or ""),
        )
        dlg.data_changed.connect(self._refresh_email_counts)
        self._inbox_dialog = dlg
        dlg.finished.connect(lambda _=None: setattr(self, "_inbox_dialog", None))
        dlg.show()

    def _on_refresh_tracked_threads(self, silent: bool = False):
        if getattr(self, "_tracked_threads_sync_worker", None) and self._tracked_threads_sync_worker.isRunning():
            if not silent:
                notify_info(self, "Thread Sync Running", "Tracked email thread sync is already in progress.")
            return
        if not hasattr(self, "tracked_threads_db"):
            return
        meeting_id = self._get_meeting_id_label()
        threads = self.tracked_threads_db.get_threads_for_meeting(meeting_id)
        if not threads:
            if not silent:
                notify_info(self, "Tracked Threads", "No pinned email threads found for this meeting.")
            return

        self._sync_silent_mode = bool(silent)
        titles = [str(t.get("title") or t.get("thread_id") or "") for t in threads]
        if hasattr(self, "thread_sync_bar"):
            self.thread_sync_bar.start_busy("Syncing tracked threads...", total=len(titles), remaining=titles)
        if hasattr(self, "email_btn"):
            if not hasattr(self, "_default_email_btn_tooltip"):
                self._default_email_btn_tooltip = self.email_btn.toolTip()
            self.email_btn.setText(f"⏳ Emails (0/{len(titles)}) ▾")

        worker = TrackedThreadsRefreshWorker(self.tracked_threads_db.db_path, meeting_id, parent=self)
        worker.progress.connect(self._on_tracked_threads_sync_progress)
        worker.thread_updated.connect(lambda _tid, _n: (self._sync_tracked_threads_to_table(), self._refresh_inbox_dialog_if_open()))
        worker.finished_sync.connect(self._on_tracked_threads_sync_finished)
        self._tracked_threads_sync_worker = worker
        worker.start()

    def _on_cancel_tracked_threads_sync(self):
        if getattr(self, "_tracked_threads_sync_worker", None) and self._tracked_threads_sync_worker.isRunning():
            self._tracked_threads_sync_worker.request_cancel()

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
        if enabled:
            self.tracked_threads_timer.start()
        else:
            self.tracked_threads_timer.stop()
