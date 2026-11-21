"""Read/flag/status/note triage mutations and anchor/Outlook link dispatch mixin."""

from __future__ import annotations

from pathlib import Path
from typing import List
import urllib.parse

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QMessageBox

from core.ui.toast import notify_info, notify_toast
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.thread_markdown_exporter import is_cacheable_document_url
from modules.emails.core.tracked_threads_service import (
    cache_and_open_thread_document_url,
    open_outlook_attachment,
    open_outlook_message,
    reload_outlook_message,
    reply_to_outlook_message,
)
from ..widgets import _status_color_hex, get_status_combo_qss


class InboxTriageActionsMixin:
    """Implements bidirectional read/flag/status/note updates and interactive reader links."""

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
