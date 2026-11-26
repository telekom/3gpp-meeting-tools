"""Data loading, filtering, Markdown timeline rendering, and unread navigation mixin."""

from pathlib import Path

from modules.emails.core.ai_summary import SummaryScope
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    find_unread_heading_blocks,
    scroll_browser_to_block,
)


class TrackedThreadRenderingMixin:
    """Populates thread controls, filters messages, and navigates unread headings."""

    def _populate_company_filter(self, messages: list):
        companies = sorted({
            str(m.get("company") or "").strip()
            for m in messages
            if str(m.get("company") or "").strip() and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
        }, key=str.lower)
        self.company_filter_combo.blockSignals(True)
        self.company_filter_combo.updateItems(companies)
        self.company_filter_combo.blockSignals(False)

    def _load_data(self):
        thread = self.db.get_thread(self.thread_id)
        if not thread:
            return
        self._loading = True
        try:
            self.title_label.setText(f"📧 {thread.get('title', 'Email Thread')}")
            self.meta_label.setText(
                f"{thread.get('meeting_id') or 'N/A'} | "
                f"{thread.get('message_count', 0)} msgs | "
                f"Unread: {thread.get('unread_count', 0)} | "
                f"Flagged: {thread.get('flagged_count', 0)}"
            )
            self.meta_label.setToolTip(f"Last synced: {thread.get('last_synced_at') or 'Never'}")
            self._update_links_toggle_label(thread)
            self.ais_pills.setText(", ".join(thread.get("linked_ais") or []))
            self.sub_items_pills.setText(", ".join(thread.get("linked_sub_items") or []))
            self.wis_pills.setText(", ".join(thread.get("linked_wis") or []))
            self.tdocs_pills.setText(", ".join(thread.get("linked_tdocs") or []))
            self.tags_pills.setText(", ".join(thread.get("user_tags") or []))
            self._refresh_sub_items_dropdown(thread.get("linked_ais") or [])
            status = str(thread.get("user_status") or "⚪ Neutral").strip()
            has_notes = bool(str(thread.get("user_notes") or "").strip())
            if hasattr(self, "notes_btn"):
                if status and status != "⚪ Neutral":
                    self.notes_btn.setText(f"{status} · 📝 Notes")
                elif has_notes:
                    self.notes_btn.setText("📓 Notes && Status...")
                else:
                    self.notes_btn.setText("📝 Notes && Status...")
            self._populate_company_filter(self.db.get_thread_messages(self.thread_id))
        finally:
            self._loading = False
        self._render_messages()
        if hasattr(self, "summary_card"):
            msgs = self.db.get_thread_messages(self.thread_id)
            store = getattr(self, "_summary_store", None)
            mdir = store.db_path.parent.parent if store is not None else None
            self.summary_card.bind_scope(
                store,
                SummaryScope(
                    scope_type="thread",
                    scope_key=self.thread_id,
                    title=str(thread.get("title") or self.thread_id),
                ),
                msgs,
                meeting_dir=mdir,
                wg=str(thread.get("wg") or ""),
                meeting_id=str(thread.get("meeting_id") or ""),
            )

    def _render_messages(self):
        thread = self.db.get_thread(self.thread_id)
        all_messages = self.db.get_thread_messages(self.thread_id)
        if not thread:
            return

        query = (self.msg_search_edit.text() if hasattr(self, "msg_search_edit") else "").strip().lower()
        filter_comp = hasattr(self, "company_filter_combo") and self.company_filter_combo.isFiltered()
        selected_comps = (
            {str(c).strip().lower() for c in self.company_filter_combo.getCheckedItems() if str(c).strip()}
            if filter_comp
            else set()
        )
        state_filter = (self.state_filter_combo.currentData() if hasattr(self, "state_filter_combo") else "all") or "all"

        filtered = []
        for m in all_messages:
            if state_filter == "unread" and int(m.get("is_read", 0) or 0) != 0:
                continue
            if state_filter == "flagged" and int(m.get("flag_status", 0) or 0) != 2:
                continue
            if filter_comp and (m.get("company") or "").strip().lower() not in selected_comps:
                continue
            if query:
                haystack = " ".join([
                    str(m.get("sender_name") or ""),
                    str(m.get("sender_email") or ""),
                    str(m.get("company") or ""),
                    str(m.get("subject") or ""),
                    str(m.get("body_text") or ""),
                ]).lower()
                if query not in haystack:
                    continue
            filtered.append(m)

        if hasattr(self, "msg_count_lbl"):
            if len(filtered) == len(all_messages):
                self.msg_count_lbl.setText(f"{len(all_messages)} message(s)")
            else:
                self.msg_count_lbl.setText(f"Showing {len(filtered)} of {len(all_messages)}")

        newest_first = bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else True
        include_meta = self.show_meta_action.isChecked() if hasattr(self, "show_meta_action") else False
        md = export_thread_to_markdown(
            thread,
            filtered,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=True,
            newest_first=newest_first,
            include_metadata=include_meta,
            enable_tdoc_links=bool(self._open_tdoc_callback),
        )
        self.browser.setMarkdown(md)
        self._unread_blocks = find_unread_heading_blocks(self.browser)
        self._unread_cursor = -1
        self._update_next_unread_button()

    def _update_next_unread_button(self):
        if not hasattr(self, "next_unread_btn"):
            return
        cnt = len(self._unread_blocks)
        self.next_unread_btn.setText(f"🔵 Next Unread ({cnt})")
        self.next_unread_btn.setEnabled(cnt > 0)
        if cnt <= 0:
            self.next_unread_btn.setToolTip("No unread emails in the current view.")
        elif self._unread_cursor >= 0:
            self.next_unread_btn.setToolTip(
                f"Jump to next unread email ({self._unread_cursor + 1}/{cnt})."
            )
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

