"""Bidirectional read, flag, status, note, and objection triage mutation mixin."""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtCore import Qt

from core.ui.toast import notify_info
from modules.emails.core.emeeting_signals import extract_base_tdoc_id
from modules.emails.core.outlook_client import OutlookClient
from ..widgets import _status_color_hex, get_status_combo_qss


class InboxTriageMutationsMixin:
    """Handles read/flag/status/note/objection mutations and user_tdocs.db sync."""

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
        self._rebuild_all_grouped_datasets()
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
            can_annotate = bool(targets) and (
                self.tracked_db is not None
                or getattr(self, "tdocs_db", None) is not None
                or getattr(self, "_user_data_callback", None) is not None
            )
            self.detail_status_combo.setEnabled(can_annotate)
            self.btn_toggle_note.setEnabled(can_annotate)
            if hasattr(self, "btn_toggle_objected"):
                self.btn_toggle_objected.setEnabled(can_annotate)
                any_obj = any(bool(it.get("is_objected")) for it in targets)
                self.btn_toggle_objected.blockSignals(True)
                self.btn_toggle_objected.setChecked(any_obj)
                self.btn_toggle_objected.setText("🛑 Objected ✓" if any_obj else "🛑 Objected")
                self.btn_toggle_objected.blockSignals(False)

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
            state_txt, state_tip = self._format_row_state_cell_text_and_tip(vis_it)
            item_cell.setText(state_txt)
            item_cell.setData(Qt.UserRole + 1, _status_color_hex(u_status))
            item_cell.setToolTip(state_tip)

    def _sync_tdoc_user_data(
        self,
        base_tdoc: str,
        *,
        new_status: str = "",
        new_notes: Optional[str] = None,
    ):
        """Synchronizes status and notes directly with TDocsWindow (`user_tdocs.db` + live model callback)."""
        btd = extract_base_tdoc_id(base_tdoc)
        if not btd:
            return
        meta = self._tdoc_meta_map.setdefault(btd, {})
        cur_status = new_status or str(meta.get("my_status") or "⚪ Neutral").strip() or "⚪ Neutral"
        cur_notes = new_notes if new_notes is not None else str(meta.get("my_notes") or "").strip()
        meta["my_status"] = cur_status
        meta["my_notes"] = cur_notes
        if getattr(self, "_user_data_callback", None) is not None:
            try:
                self._user_data_callback(btd, cur_status, cur_notes)
                return
            except Exception:
                pass
        if getattr(self, "tdocs_db", None) is not None and hasattr(self.tdocs_db, "update_tdoc_user_data"):
            try:
                self.tdocs_db.update_tdoc_user_data(btd, status=cur_status, notes=cur_notes)
            except Exception:
                pass

    def _on_detail_status_changed(self):
        if getattr(self, "_updating_status_bar", False):
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
            primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
            if not primary_tdoc and it.get("linked_tdocs"):
                primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
            if primary_tdoc:
                self._sync_tdoc_user_data(primary_tdoc, new_status=new_status)

            tid = str(it.get("thread_id") or "")
            if not it.get("in_tracked_db") and self.tracked_db is not None and self._view_mode != "TDOCS":
                tid = self._pin_item_in_place(it, reload=False)
                needs_full_reload = True
            if tid and tid not in seen_tids and self.tracked_db is not None:
                seen_tids.add(tid)
                self.tracked_db.update_thread_annotations(tid, user_status=new_status)
                for m in self._all_items:
                    if m.get("thread_id") == tid:
                        m["user_status"] = new_status
            it["user_status"] = new_status
            if self._view_mode == "TDOCS" and primary_tdoc:
                preserved_ids.append(f"TDOC::{primary_tdoc}")
            elif self._view_mode == "THREADS" and tid:
                preserved_ids.append(f"THR::{tid}")
            else:
                preserved_ids.append(it["id"])

        self._rebuild_all_grouped_datasets()
        self._populate_filter_combos()
        self._update_pill_counts()
        status_filter_active = hasattr(self, "status_combo") and self.status_combo.isFiltered()
        if (
            needs_full_reload
            or status_filter_active
            or self._view_mode == "TDOCS"
            or self._active_mode == "PRIORITY"
        ):
            self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
        else:
            for tid in seen_tids:
                self._refresh_table_state_cells_for_thread(tid)
        self.data_changed.emit()

    def _on_toggle_objected_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets:
            return
        mark_obj = any(not bool(it.get("is_objected")) for it in targets)
        preserved_ids: List[str] = []

        for it in targets:
            sig = it.get("_signals") or {}
            obj_rev = str(sig.get("max_rev_tag") or "").strip() if mark_obj else ""
            primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
            if not primary_tdoc and it.get("linked_tdocs"):
                primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
            if (
                primary_tdoc
                and getattr(self, "tdocs_db", None) is not None
                and hasattr(self.tdocs_db, "set_tdoc_objected")
            ):
                self.tdocs_db.set_tdoc_objected(primary_tdoc, mark_obj, objected_at_rev=obj_rev)
                meta = self._tdoc_meta_map.setdefault(primary_tdoc, {})
                meta["is_objected"] = mark_obj
                meta["objected_at_rev"] = obj_rev

            tid = str(it.get("thread_id") or "")
            if not tid and self.tracked_db is not None and not primary_tdoc:
                tid = self._pin_item_in_place(it, reload=False)
            if tid and self.tracked_db is not None:
                self.tracked_db.update_thread_annotations(
                    tid,
                    is_objected=mark_obj,
                    objected_at_rev=obj_rev,
                )
            if self._view_mode == "TDOCS" and primary_tdoc:
                preserved_ids.append(f"TDOC::{primary_tdoc}")
            elif self._view_mode == "THREADS" and tid:
                preserved_ids.append(f"THR::{tid}")
            else:
                preserved_ids.append(it["id"])

        self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
        self.data_changed.emit()

    def _on_toggle_note_clicked(self, checked: bool):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if len(targets) != 1:
            if len(targets) > 1:
                notify_info(
                    self,
                    "Select Single Item",
                    "Select a single TDoc, thread, or email to add/edit its comment.",
                )
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
        if not getattr(self, "_note_dirty", False):
            return
        it = self._note_target_item
        if not it:
            self._note_dirty = False
            return
        new_note = self.note_edit.toPlainText().strip()
        self._note_dirty = False

        primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
        if not primary_tdoc and it.get("linked_tdocs"):
            primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
        if primary_tdoc:
            self._sync_tdoc_user_data(primary_tdoc, new_notes=new_note)

        if not it.get("in_tracked_db") and self.tracked_db is not None and self._view_mode != "TDOCS":
            if not new_note and not primary_tdoc:
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
            if tid and self.tracked_db is not None:
                self.tracked_db.update_thread_annotations(tid, user_notes=new_note)
                for m in self._all_items:
                    if m.get("thread_id") == tid:
                        m["user_notes"] = new_note
            it["user_notes"] = new_note
            self._rebuild_all_grouped_datasets()
            if tid:
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
        self._rebuild_all_grouped_datasets()
        self._update_pill_counts()
        self.data_changed.emit()

    def _on_mark_read_and_next(self):
        if not self._current_item:
            return
        cur_row = self.table.currentRow()
        it = self._current_item
        if not it["is_read"]:
            self._mutate_read(it, True)
            if self._active_mode in ("UNREAD", "PRIORITY"):
                self._apply_filters(preserve_id="", select_row=max(0, cur_row))
                return
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
                select_row=cur_row if self._active_mode in ("UNREAD", "PRIORITY") else None,
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
