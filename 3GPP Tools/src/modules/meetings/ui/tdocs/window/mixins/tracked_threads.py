# --- File: src/modules/meetings/ui/tdocs/window/mixins/tracked_threads.py ---
"""
Tracked e-mail threads: pinning (Outlook/.msg), sync worker, inbox dialog, table sync.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
import re
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QMessageBox,
    QDialog,
)
from modules.meetings.ui.tdocs_models import natural_sort_key
from modules.emails.core.tracked_threads_service import (
    pin_thread_from_msg,
    pin_thread_from_outlook_selection,
    pin_threads_from_msg_files,
)
from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog
from modules.emails.ui.tracked_thread_dialog import PinFromOutlookDialog, TrackedThreadDialog
from core.ui.toast import notify_success
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
            tdoc_counts_map = {}
            if hasattr(self, "general_email_db") and self.general_email_db is not None:
                tdoc_counts_map = self.general_email_db.get_email_counts_per_tdoc() or {}
                for stats in tdoc_counts_map.values():
                    tdoc_unread_ids.update(str(eid) for eid in (stats.get("unread_ids") or ()) if eid)
            total_unread = len(tracked_unread_ids | tdoc_unread_ids)
            priority_unread_ids = set()
            if total_unread > 0:
                try:
                    from modules.emails.ui.meeting_inbox.followed_ais_helper import compute_inherited_followed_ais
                    meta_map = self._build_tdoc_meta_map()
                    gdb_path = getattr(self.general_email_db, "db_path", None) if getattr(self, "general_email_db", None) else None
                    followed_ais = compute_inherited_followed_ais(tdoc_meta_map=meta_map, general_email_db_path=gdb_path)
                    for t in threads:
                        if bool(t.get("is_objected")):
                            continue
                        st = str(t.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                        t_ais = set(t.get("linked_ais") or [])
                        is_prio_th = (st != "⚪ Neutral") or bool(t.get("flagged_count")) or bool(followed_ais and (t_ais & followed_ais))
                        if is_prio_th:
                            for m in self.tracked_threads_db.get_thread_messages(t["thread_id"]):
                                if int(m.get("is_read", 0) or 0) == 0 and m.get("id"):
                                    priority_unread_ids.add(str(m["id"]))
                    for tdoc_key, stats in tdoc_counts_map.items():
                        u_ids = stats.get("unread_ids") or set()
                        if not u_ids:
                            continue
                        tmeta = meta_map.get(tdoc_key.upper()) or {}
                        if bool(tmeta.get("is_objected")):
                            continue
                        tst = str(tmeta.get("my_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                        tai = str(tmeta.get("ai") or "").strip()
                        if (tst != "⚪ Neutral") or bool(stats.get("flagged")) or (tai and tai in followed_ais):
                            priority_unread_ids.update(str(eid) for eid in u_ids if eid)
                except Exception:
                    pass
            prio_cnt = len(priority_unread_ids)
            if hasattr(self, "email_btn"):
                if total_unread > 0:
                    if prio_cnt > 0:
                        self.email_btn.setText(f"📧 Emails (🔥 {prio_cnt} Priority | {total_unread} new) ▾")
                    else:
                        self.email_btn.setText(f"📧 Emails ({total_unread} new) ▾")
                    self.email_btn.setToolTip(
                        f"Email Discussions & Tracked Threads\n"
                        f"{total_unread} unread email(s) total ({prio_cnt} on your Priority Radar; "
                        f"{len(tracked_unread_ids)} in Tracked Threads, {len(tdoc_unread_ids)} in TDoc Emails)."
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
        """Maps TDoc IDs to their Agenda Item, Sub-Item, WIs, Title, Source, Type, Status, and user annotations."""
        meta_map = {}
        if not hasattr(self, "model") or not hasattr(self.model, "_data"):
            return meta_map
        user_data = {}
        if hasattr(self, "db") and self.db is not None and hasattr(self.db, "get_all_user_data"):
            try:
                user_data = self.db.get_all_user_data() or {}
            except Exception:
                user_data = {}
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
            ud = user_data.get(td) or {}
            meta_map[td] = {
                "ai": ai,
                "sub_item": sub,
                "wis": wis,
                "title": str(r.get("Title") or "").strip(),
                "source": str(r.get("Source") or "").strip(),
                "type": str(r.get("Type") or "").strip(),
                "tdoc_status": str(r.get("TDoc Status") or "").strip(),
                "my_status": str(r.get("My Status") or ud.get("my_status") or "⚪ Neutral").strip() or "⚪ Neutral",
                "my_notes": str(r.get("My Notes") or ud.get("my_notes") or "").strip(),
                "is_objected": bool(ud.get("is_objected")),
                "objected_at_rev": str(ud.get("objected_at_rev") or "").strip(),
            }
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
        general_db = getattr(self, "general_email_db", None)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            pinned_ids = pin_thread_from_outlook_selection(
                self.tracked_threads_db,
                meeting_id=meeting_id,
                wg=wg,
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=tdoc_meta_map,
                general_email_db=general_db,
            )
            if not pinned_ids and mime_data:
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
                                    general_email_db=general_db,
                                )
                                if tid:
                                    pinned_ids.append(tid)
            if pinned_ids:
                self._on_tracked_thread_changed()
                if hasattr(self, "_refresh_email_counts"):
                    self._refresh_email_counts()
                if len(pinned_ids) == 1:
                    self._open_tracked_thread_dialog(pinned_ids[0])
                else:
                    notify_success(
                        self,
                        "Threads Imported",
                        f"Imported {len(pinned_ids)} email thread(s) with automatic AI/WI/TDoc linking.",
                    )
                    self._open_meeting_email_inbox()
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
        finally:
            QApplication.restoreOverrideCursor()

    def _process_dropped_msg_files(self, msg_files: list):
        wg = str(self.mtg_info.get("wg_name", "SA2")).strip().upper()
        meeting_id = self._get_meeting_id_label()
        agenda_map, valid_wis = self._get_pin_meeting_context()
        tdoc_meta_map = self._build_tdoc_meta_map()
        general_db = getattr(self, "general_email_db", None)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            res = pin_threads_from_msg_files(
                self.tracked_threads_db,
                [Path(mp) for mp in msg_files],
                meeting_id=meeting_id,
                wg=wg,
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=tdoc_meta_map,
                general_email_db=general_db,
            )
            pinned_ids = list(res.get("thread_ids") or [])
            if pinned_ids:
                self._on_tracked_thread_changed()
                if hasattr(self, "_refresh_email_counts"):
                    self._refresh_email_counts()
                if len(pinned_ids) == 1:
                    self._open_tracked_thread_dialog(pinned_ids[0])
                else:
                    notify_success(
                        self,
                        "Threads Imported",
                        f"Imported {res.get('message_count', len(msg_files))} email(s) across {len(pinned_ids)} thread(s).",
                    )
                    self._open_meeting_email_inbox()
            else:
                QMessageBox.warning(
                    self,
                    "Could Not Parse .msg",
                    "Failed to parse the dropped .msg file via Outlook COM. Ensure Microsoft Outlook is available.",
                )
        except Exception as e:
            logging.error(f"Error processing dropped .msg file(s): {e}", exc_info=True)
            QMessageBox.warning(self, "Error Processing .msg", f"Failed to process dropped .msg file:\n{e}")
        finally:
            QApplication.restoreOverrideCursor()

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
            tdocs_db=getattr(self, "db", None),
            user_data_callback=getattr(self, "_save_user_data", None),
            jump_tdoc_callback=getattr(self, "_jump_to_tdoc_from_signal", None),
        )
        dlg.data_changed.connect(self._refresh_email_counts)
        dlg.live_sync_toggled.connect(self._on_inbox_live_sync_toggled)
        self._inbox_dialog = dlg
        dlg.finished.connect(lambda _=None: setattr(self, "_inbox_dialog", None))
        dlg.show()

    def _open_email_briefings_dialog(self, initial_ai_code: str = ""):
        """Opens the modeless Per-AI & Saved Selections AI Email Briefings window."""
        from modules.emails.ui.briefings import EmailBriefingsDialog
        from modules.emails.ui.meeting_inbox.followed_ais_helper import load_effective_followed_ais

        existing = getattr(self, "_briefings_dialog", None)
        if existing is not None:
            try:
                if initial_ai_code:
                    existing.focus_ai(initial_ai_code, auto_run=False)
                else:
                    existing.refresh_navigation()
                existing.setWindowState(existing.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
                existing.raise_()
                existing.activateWindow()
                return
            except RuntimeError:
                self._briefings_dialog = None

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
        ai_sugg, _wi_sugg, _tdoc_sugg, _sub_sugg, _ai_to_sub = self._build_thread_autocomplete_suggestions()
        ai_label_map = {}
        for item in ai_sugg:
            code = str(item).split(" — ", 1)[0].strip()
            if code and code not in ai_label_map:
                ai_label_map[code] = str(item)
        meta_map = self._build_tdoc_meta_map()
        gdb = getattr(self, "general_email_db", None)
        followed_ais, _inherited, _cust = load_effective_followed_ais(
            meeting_id=self._get_meeting_id_label(),
            tdoc_meta_map=meta_map,
            general_email_db_path=getattr(gdb, "db_path", None) if gdb else None,
        )
        tdb = getattr(self, "tracked_threads_db", None)
        base_p = getattr(tdb, "db_path", None) or getattr(gdb, "db_path", None)
        sum_db_p = (Path(base_p).parent / "email_summaries.db") if base_p else (self.meeting_dir / "Agenda" / "email_summaries.db")

        dlg = EmailBriefingsDialog(
            tracked_db=tdb,
            general_email_db=gdb,
            summary_db_path=sum_db_p,
            meeting_id=self._get_meeting_id_label(),
            wg=wg,
            ai_label_map=ai_label_map,
            followed_ais=followed_ais,
            tdoc_ai_lookup=tdoc_ai_lookup,
            tdoc_meta_map=meta_map,
            open_tdoc_callback=getattr(self, "_open_tdoc_document_directly", None),
            open_thread_callback=getattr(self, "_open_tracked_thread_dialog", None),
            jump_tdoc_callback=getattr(self, "_jump_to_tdoc_from_signal", None),
            initial_ai_code=initial_ai_code,
            parent=None,
        )
        self._briefings_dialog = dlg
        dlg.finished.connect(lambda _=None: setattr(self, "_briefings_dialog", None))
        dlg.show()
