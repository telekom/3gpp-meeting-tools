"""Context menu, split-thread merging, and adaptive async Outlook sync mixin."""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtWidgets import QDialog, QMenu

from core.ui.theme import C
from core.ui.toast import notify_info, notify_success, notify_toast


class InboxSyncAndMergeMixin:
    """Manages row context menus, thread merge workflows, and scoped/full Outlook syncs."""

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
        if hasattr(self, "btn_toggle_objected") and self.btn_toggle_objected.isEnabled():
            any_obj = any(bool(t.get("is_objected")) for t in targets)
            obj_lbl = "🛑 Unmark Objected (Unpark)" if any_obj else "🛑 Mark as Objected (Park)"
            act_obj = menu.addAction(obj_lbl)
            act_obj.triggered.connect(self._on_toggle_objected_clicked)
        if self.btn_open_context.isEnabled():
            menu.addSeparator()
            act_ctx = menu.addAction(self.btn_open_context.text())
            act_ctx.triggered.connect(self._on_open_context)
        if hasattr(self, "btn_jump_tdoc_table") and self.btn_jump_tdoc_table.isEnabled():
            act_jmp = menu.addAction("🎯 Show TDoc in Meeting Table")
            act_jmp.triggered.connect(self._on_jump_tdoc_table_clicked)
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
