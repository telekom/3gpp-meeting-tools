"""Timeline filtering, Markdown rendering, unread navigation, and note saving mixin."""

from typing import List

from modules.emails.core.ai_summary import SummaryScope
from modules.emails.core.emeeting_signals import extract_base_tdoc_id
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    find_unread_heading_blocks,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_service import normalize_subject
from modules.emails.ui.meeting_email_inbox_dialog import get_status_combo_qss


class TDocEmailsRenderingMixin:
    """Loads, filters, and renders TDoc email timelines plus status/note state."""

    @staticmethod
    def _compress_blank_lines(text: str, max_consecutive_empty: int = 1) -> str:
        if not text:
            return ""
        normalized = text.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
        lines = normalized.split("\n")
        cleaned_lines = []
        empty_count = 0
        for line in lines:
            if not line.strip():
                empty_count += 1
                if empty_count <= max_consecutive_empty:
                    cleaned_lines.append("")
            else:
                empty_count = 0
                cleaned_lines.append(line.rstrip())
        return "\n".join(cleaned_lines).strip()

    def _on_order_changed(self):
        self._settings.setValue("newest_first", bool(self.order_combo.currentData()))
        self._render_messages()

    def _on_show_metadata_toggled(self, checked: bool):
        self._settings.setValue("show_metadata_header", bool(checked))
        self._render_messages()

    def _build_synthetic_thread_dict(self) -> dict:
        scope = self.family_tdocs if self.chk_family.isChecked() else [self.target_tdoc]
        base_title = normalize_subject(self.emails[0].get("subject", "")) if self.emails else ""
        return {
            "thread_id": f"TDOC-{self.target_tdoc}",
            "title": f"Emails Referencing {self.target_tdoc}" + (f": {base_title}" if base_title else ""),
            "normalized_subject": base_title.lower(),
            "wg": self.wg,
            "linked_tdocs": list(scope),
        }

    def _populate_company_filter(self, messages: List[dict]):
        companies = sorted(
            {
                str(m.get("company") or "").strip()
                for m in messages
                if str(m.get("company") or "").strip()
                and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
            },
            key=str.lower,
        )
        self.company_filter_combo.blockSignals(True)
        self.company_filter_combo.updateItems(companies)
        self.company_filter_combo.blockSignals(False)

    def _load_emails(self):
        self._loading = True
        try:
            query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
            raw_emails = self.db.get_emails_for_tdocs(query_set, show_ignored=self.chk_show_ignored.isChecked())
            if not self.chk_include_quoted.isChecked():
                self.emails = [e for e in raw_emails if e.get("match_location") != "Quoted"]
            else:
                self.emails = raw_emails

            # Normalize chronological ascending order internally; display order respects order_combo
            self.emails.sort(key=lambda m: str(m.get("date_received") or m.get("date_sent") or ""))
            self._populate_company_filter(self.emails)
            self._update_unread_count_label()
        finally:
            self._loading = False
        self._render_messages()
        if hasattr(self, "summary_card"):
            base_tdoc = extract_base_tdoc_id(self.target_tdoc)
            store = getattr(self, "_summary_store", None)
            mdir = store.db_path.parent.parent if store is not None else None
            tmeta = {
                "title": str(self._tdoc_info.get("Title") or "").strip(),
                "source": str(self._tdoc_info.get("Source") or "").strip(),
                "ai": str(self._tdoc_info.get("Agenda Item") or "").strip(),
                "tdoc_status": str(self._tdoc_info.get("TDoc Status") or "").strip(),
                "my_status": str(self._tdoc_info.get("My Status") or "⚪ Neutral").strip(),
            }
            self.summary_card.bind_scope(
                store,
                SummaryScope(
                    scope_type="tdoc",
                    scope_key=base_tdoc,
                    title=str(self._tdoc_info.get("Title") or base_tdoc),
                ),
                self.emails,
                meeting_dir=mdir,
                wg=self.wg,
                tdoc_meta=tmeta,
            )

    def _update_unread_count_label(self):
        unread = sum(1 for e in self.emails if int(e.get("is_read", 0) or 0) == 0 and int(e.get("is_ignored", 0) or 0) == 0)
        flagged = sum(1 for e in self.emails if int(e.get("flag_status", 0) or 0) == 2 and int(e.get("is_ignored", 0) or 0) == 0)
        flag_part = f", {flagged} Flagged" if flagged > 0 else ""
        self.lbl_counts.setText(f"{len(self.emails)} Total ({unread} Unread{flag_part})")

    def _render_messages(self):
        if self._loading:
            return
        query = self.msg_search_edit.text().strip().lower()
        filter_comp = self.company_filter_combo.isFiltered()
        selected_comps = (
            {str(c).strip().lower() for c in self.company_filter_combo.getCheckedItems() if str(c).strip()}
            if filter_comp
            else set()
        )
        state_filter = (self.state_filter_combo.currentData() or "all").strip().lower()

        filtered: List[dict] = []
        for m in self.emails:
            if state_filter == "unread" and int(m.get("is_read", 0) or 0) != 0:
                continue
            if state_filter == "flagged" and int(m.get("flag_status", 0) or 0) != 2:
                continue
            if filter_comp and str(m.get("company") or "").strip().lower() not in selected_comps:
                continue
            if query:
                hay = " ".join([
                    str(m.get("sender_name") or ""),
                    str(m.get("sender_email") or ""),
                    str(m.get("company") or ""),
                    str(m.get("subject") or ""),
                    str(m.get("matched_tdoc") or ""),
                    str(m.get("body_text") or ""),
                ]).lower()
                if query not in hay:
                    continue
            filtered.append(m)

        self._filtered_emails = filtered
        if len(filtered) == len(self.emails):
            self.msg_count_lbl.setText(f"{len(self.emails)} message(s)")
        else:
            self.msg_count_lbl.setText(f"Showing {len(filtered)} of {len(self.emails)}")

        newest_first = bool(self.order_combo.currentData())
        include_meta = self.show_meta_action.isChecked()
        thread_meta = self._build_synthetic_thread_dict()
        md = export_thread_to_markdown(
            thread_meta,
            filtered,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=True,
            newest_first=newest_first,
            include_metadata=include_meta,
            include_ignore_action=True,
            enable_tdoc_links=True,
            include_discussion_heading=False,
        )
        self.browser.setMarkdown(md)
        self._unread_blocks = find_unread_heading_blocks(self.browser)
        self._unread_cursor = -1
        self._update_next_unread_button()

    def _update_next_unread_button(self):
        cnt = len(self._unread_blocks)
        self.next_unread_btn.setText(f"🔵 Next Unread ({cnt})")
        self.next_unread_btn.setEnabled(cnt > 0)
        if cnt <= 0:
            self.next_unread_btn.setToolTip("No unread emails in the current view.")
        elif self._unread_cursor >= 0:
            self.next_unread_btn.setToolTip(f"Jump to next unread email ({self._unread_cursor + 1}/{cnt}).")
        else:
            self.next_unread_btn.setToolTip(f"Jump to the first of {cnt} unread email(s).")

    def _scroll_to_initial_unread(self):
        if self._initial_scroll_done:
            return
        self._initial_scroll_done = True
        blocks = find_unread_heading_blocks(self.browser)
        self._unread_blocks = blocks
        if blocks:
            self._unread_cursor = 0
            scroll_browser_to_block(self.browser, blocks[0])
            self._update_next_unread_button()

    def _on_next_unread_clicked(self):
        blocks = find_unread_heading_blocks(self.browser)
        self._unread_blocks = blocks
        if not blocks:
            self._update_next_unread_button()
            return
        self._unread_cursor = (self._unread_cursor + 1) % len(blocks)
        scroll_browser_to_block(self.browser, blocks[self._unread_cursor])
        self._update_next_unread_button()

    def _populate_tdoc_status_and_note(self):
        self._updating_status_bar = True
        self._updating_note_bar = True
        try:
            st = str(self._tdoc_info.get("My Status") or "⚪ Neutral").strip() or "⚪ Neutral"
            idx = self.detail_status_combo.findText(st)
            self.detail_status_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self.detail_status_combo.setStyleSheet(get_status_combo_qss(st))

            note_txt = str(self._tdoc_info.get("My Notes") or "").strip()
            self._note_dirty = False
            self.note_edit.setPlainText(note_txt)
            has_note = bool(note_txt)
            self.note_frame.setVisible(has_note)
            self.btn_toggle_note.setChecked(has_note)
            self.btn_toggle_note.setText("📝 Note ✓" if has_note else "📝 Note")
            self.note_saved_lbl.setText("")

            can_save = self._user_data_callback is not None
            self.detail_status_combo.setEnabled(can_save)
            self.btn_toggle_note.setEnabled(can_save)
            self.note_edit.setEnabled(can_save)
        finally:
            self._updating_status_bar = False
            self._updating_note_bar = False

    def _on_tdoc_status_changed(self):
        if getattr(self, "_updating_status_bar", False) or self._user_data_callback is None:
            return
        new_status = str(self.detail_status_combo.currentText() or "⚪ Neutral").strip() or "⚪ Neutral"
        self._tdoc_info["My Status"] = new_status
        self.detail_status_combo.setStyleSheet(get_status_combo_qss(new_status))
        curr_notes = str(self.note_edit.toPlainText() or "").strip()
        self._tdoc_info["My Notes"] = curr_notes
        self._user_data_callback(self.target_tdoc, new_status, curr_notes)
        self.data_changed.emit()

    def _on_toggle_note_clicked(self, checked: bool):
        if checked or str(self.note_edit.toPlainText() or "").strip():
            self.note_frame.setVisible(True)
            self.note_edit.setFocus()
        else:
            self.note_frame.setVisible(False)

    def _on_note_text_changed(self):
        if getattr(self, "_updating_note_bar", False):
            return
        self._note_dirty = True
        self.note_saved_lbl.setText("Saving...")
        self._note_save_timer.start()

    def _flush_pending_note_save(self):
        if self._note_save_timer.isActive():
            self._note_save_timer.stop()
        if not getattr(self, "_note_dirty", False) or self._user_data_callback is None:
            return
        self._note_dirty = False
        new_note = self.note_edit.toPlainText().strip()
        curr_status = str(self.detail_status_combo.currentText() or "⚪ Neutral").strip() or "⚪ Neutral"
        self._tdoc_info["My Notes"] = new_note
        self._tdoc_info["My Status"] = curr_status
        self._user_data_callback(self.target_tdoc, curr_status, new_note)
        has_note = bool(new_note)
        self.note_saved_lbl.setText("✓ Saved" if has_note else "")
        self.btn_toggle_note.setText("📝 Note ✓" if has_note else "📝 Note")
        self.data_changed.emit()
