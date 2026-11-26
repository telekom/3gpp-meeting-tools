"""Outlook reply/open, LLM export, statistics dashboard, and reader anchor dispatch mixin."""

from __future__ import annotations

import os
from pathlib import Path
import urllib.parse
import webbrowser

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QCursor, QDesktopServices
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox

from core.ai.llm_runner import ModelSpec, resolve_effective_model_spec
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.ai_summary import SummaryScope
from modules.emails.core.emeeting_signals import extract_base_tdoc_id
from modules.emails.ui.ai_summary import populate_summary_model_menu
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    find_heading_block_for_message_index,
    is_cacheable_document_url,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_service import (
    cache_and_open_thread_document_url,
    normalize_subject,
    open_outlook_attachment,
    open_outlook_message,
    reload_outlook_message,
    reply_to_outlook_message,
)


class InboxReaderDispatchMixin:
    """Implements Outlook message actions, pure Markdown/AI export, stats, and reader anchor clicks."""

    def _latest_outlook_message_id(self) -> str:
        it = self._current_item
        if not it:
            return ""
        if it.get("_is_group") and it.get("_group_messages"):
            return str(it["_group_messages"][-1].get("id") or "")
        return str(it.get("id") or "")

    def _on_reply(self):
        msg_id = self._latest_outlook_message_id()
        if not msg_id:
            return
        if not reply_to_outlook_message(msg_id, reply_all=True):
            QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")

    def _on_open_outlook(self):
        msg_id = self._latest_outlook_message_id()
        if not msg_id:
            return
        if not open_outlook_message(msg_id):
            QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")

    def _on_copy_start_template_clicked(self):
        it = self._current_item
        if not it:
            return
        primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
        if not primary_tdoc and it.get("linked_tdocs"):
            primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
        sig = it.get("_signals") or {}
        rev_tag = str(sig.get("max_rev_tag") or "").strip()
        tdoc_ref = f"{primary_tdoc}{rev_tag}" if primary_tdoc else "TDoc"
        template = f"<<START>>\n[{tdoc_ref}] \n<<END>>"
        QApplication.clipboard().setText(template)
        notify_info(self, "e-Meeting Template Copied", f"Copied <<START>>...<<END>> block for {tdoc_ref} to clipboard.")

    def _on_jump_tdoc_table_clicked(self):
        it = self._current_item
        if not it or not getattr(self, "_jump_tdoc_callback", None):
            return
        primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
        if not primary_tdoc and it.get("linked_tdocs"):
            primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
        if primary_tdoc:
            self._jump_tdoc_callback(primary_tdoc)

    def _build_current_item_pure_markdown(self) -> str:
        it = self._current_item
        if not it:
            return ""
        msgs = (
            [
                {
                    "id": m["id"],
                    "sender_name": m.get("sender_name", ""),
                    "sender_email": m.get("sender_email", ""),
                    "company": m.get("company", ""),
                    "date_sent": m.get("date_str", ""),
                    "subject": m.get("subject", ""),
                    "body_text": m.get("raw_body_text") or m.get("body_text", ""),
                    "clean_body": m.get("clean_body") or m.get("body_text", ""),
                    "is_read": 1 if m["is_read"] else 0,
                    "flag_status": 2 if m["is_flagged"] else 0,
                    "attachments": m.get("attachments") or [],
                }
                for m in (it.get("_group_messages") or [])
            ]
            if it.get("_is_group")
            else [
                {
                    "id": it["id"],
                    "sender_name": it.get("sender_name", ""),
                    "sender_email": it.get("sender_email", ""),
                    "company": it.get("company", ""),
                    "date_sent": it.get("date_str", ""),
                    "subject": it.get("subject", ""),
                    "body_text": it.get("raw_body_text") or it.get("body_text", ""),
                    "clean_body": it.get("clean_body") or it.get("body_text", ""),
                    "is_read": 1 if it["is_read"] else 0,
                    "flag_status": 2 if it["is_flagged"] else 0,
                    "attachments": it.get("attachments") or [],
                }
            ]
        )
        synth = {
            "thread_id": it.get("thread_id") or it["id"],
            "title": it.get("thread_title") or it.get("subject") or "Email Discussion",
            "normalized_subject": normalize_subject(it.get("subject") or "").lower(),
            "meeting_id": self.meeting_id,
            "wg": self.wg,
            "linked_ais": it.get("linked_ais") or [],
            "linked_sub_items": it.get("linked_sub_items") or [],
            "linked_wis": it.get("linked_wis") or [],
            "linked_tdocs": it.get("linked_tdocs") or [],
            "user_status": it.get("user_status") or "⚪ Neutral",
            "user_notes": it.get("user_notes") or "",
        }
        return export_thread_to_markdown(
            synth,
            msgs,
            deduplicate_quotes=True,
            interactive_links=False,
            newest_first=False,
            include_metadata=True,
        )

    def _resolve_item_summary_scope(self, it: dict) -> tuple:
        """Resolve `(SummaryScope, scope_messages, tdoc_meta)` for any Inbox row item."""
        is_tdoc_bucket = it.get("_group_kind") == "TDOC" and bool(it.get("_primary_tdoc"))
        primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
        if not primary_tdoc and it.get("linked_tdocs"):
            primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
        tmeta = self._tdoc_meta_map.get(primary_tdoc) if primary_tdoc else None

        if is_tdoc_bucket and primary_tdoc:
            scope = SummaryScope(
                scope_type="tdoc",
                scope_key=primary_tdoc,
                title=str(it.get("subject") or primary_tdoc),
            )
            msgs = list(it.get("_group_messages") or [])
        elif it.get("thread_id"):
            tid = str(it["thread_id"])
            msgs = list(it.get("_group_messages") or [m for m in self._all_items if m.get("thread_id") == tid] or [it])
            scope = SummaryScope(
                scope_type="thread",
                scope_key=tid,
                title=str(it.get("thread_title") or it.get("subject") or tid),
            )
        elif primary_tdoc:
            msgs = list(it.get("_group_messages") or [it])
            scope = SummaryScope(
                scope_type="tdoc",
                scope_key=primary_tdoc,
                title=str(it.get("subject") or primary_tdoc),
            )
        else:
            norm = normalize_subject(it.get("subject") or "").lower()
            skey = f"UNPIN::{norm or it.get('id', 'msg')}"
            msgs = list(it.get("_group_messages") or [it])
            scope = SummaryScope(
                scope_type="thread",
                scope_key=skey,
                title=str(it.get("subject") or "Email Discussion"),
            )
        return scope, msgs, tmeta

    def _refresh_summary_freshness_cache(self, dataset: list) -> None:
        store = getattr(self, "_summary_store", None)
        if store is None:
            self._summary_freshness_cache = {}
            return
        buckets_by_type: dict = {"thread": {}, "tdoc": {}}
        for it in dataset:
            if not it.get("_is_group"):
                continue
            scope, msgs, _ = self._resolve_item_summary_scope(it)
            if scope.scope_type in buckets_by_type:
                buckets_by_type[scope.scope_type][scope.scope_key] = msgs
        cache = {}
        for stype, smap in buckets_by_type.items():
            if smap:
                for skey, fresh in store.get_freshness_map(stype, smap).items():
                    cache[f"{stype}:{skey}"] = fresh
        self._summary_freshness_cache = cache

    def _update_ai_briefing_strip(self, selected_ais: set) -> None:
        strip = getattr(self, "ai_briefing_strip", None)
        if strip is None:
            return
        concrete = [a for a in selected_ais if a and a != "__UNASSIGNED__"]
        if len(concrete) != 1:
            strip.setVisible(False)
            self._strip_ai_code = ""
            return
        ai_code = concrete[0]
        self._strip_ai_code = ai_code
        ai_lbl = self._ai_label_map.get(ai_code, f"AI {ai_code}")
        ai_short = (ai_lbl[:48] + "…") if len(ai_lbl) > 50 else ai_lbl
        store = getattr(self, "_summary_store", None)
        rec = store.get_summary("ai", ai_code) if store is not None else None
        ai_msgs = [m for m in self._all_items if ai_code in (m.get("linked_ais") or [])]
        fresh = store.compute_freshness("ai", ai_code, ai_msgs) if store is not None else None
        if rec is None:
            status_txt = f"🤖 AI {ai_short}: No briefing yet ({len(ai_msgs)} msgs)"
            self.btn_ai_strip_run.setText("✨ Generate")
        elif fresh is not None and fresh.new_count > 0:
            status_txt = f"🤖 AI {ai_short}: 🟡 +{fresh.new_count} new message(s)"
            self.btn_ai_strip_run.setText(f"⚡ Update (+{fresh.new_count})")
        else:
            status_txt = f"🤖 AI {ai_short}: ✅ Briefing up to date"
            self.btn_ai_strip_run.setText("🔄 Refresh")
        self.ai_briefing_strip_lbl.setText(status_txt)
        strip.setVisible(True)

    def _on_ai_strip_action(self, *, auto_run: bool = False) -> None:
        ai_code = str(getattr(self, "_strip_ai_code", "") or "").strip()
        if ai_code:
            self._open_briefings_dialog(ai_code=ai_code, auto_run=auto_run)

    def _open_briefings_dialog(
        self,
        *,
        ai_code: str = "",
        selection_payload: dict = None,
        auto_run: bool = False,
    ) -> None:
        from modules.emails.ui.briefings import EmailBriefingsDialog

        existing = getattr(self, "_briefings_dialog", None)
        if existing is not None:
            try:
                if selection_payload and selection_payload.get("messages"):
                    existing.start_selection_briefing(
                        title=str(selection_payload.get("title") or "Custom Selection"),
                        messages=selection_payload.get("messages") or [],
                        focus_prompt=str(selection_payload.get("focus_prompt") or ""),
                        auto_run=auto_run,
                    )
                elif ai_code:
                    existing.focus_ai(ai_code, auto_run=auto_run)
                else:
                    existing.refresh_navigation()
                existing.setWindowState(existing.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
                existing.raise_()
                existing.activateWindow()
                return
            except RuntimeError:
                self._briefings_dialog = None

        dlg = EmailBriefingsDialog(
            tracked_db=self.tracked_db,
            general_email_db=self.general_email_db,
            summary_db_path=self._summary_store.db_path if getattr(self, "_summary_store", None) else None,
            meeting_id=self.meeting_id,
            wg=self.wg,
            ai_label_map=self._ai_label_map,
            followed_ais=getattr(self, "_followed_ais", set()),
            tdoc_ai_lookup=self.tdoc_ai_lookup,
            tdoc_meta_map=self._tdoc_meta_map,
            open_tdoc_callback=self._open_tdoc_callback,
            open_thread_callback=self._open_thread_callback,
            jump_tdoc_callback=getattr(self, "_jump_tdoc_callback", None),
            initial_ai_code=ai_code,
            initial_selection=selection_payload,
            auto_run_initial=auto_run,
            parent=None,
        )
        dlg.briefing_saved.connect(lambda _rec: self._apply_filters())
        self._briefings_dialog = dlg
        dlg.finished.connect(lambda _=None: setattr(self, "_briefings_dialog", None))
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _summarize_selected_items(self, selected: list) -> None:
        if not selected:
            return
        dedup: dict = {}
        titles: list = []
        for it in selected:
            t = str(it.get("_primary_tdoc") or it.get("thread_title") or it.get("subject") or "").strip()
            if t and t not in titles:
                titles.append(t)
            msgs = it.get("_group_messages") if it.get("_is_group") else [it]
            for m in (msgs or []):
                mid = str(m.get("id") or "").strip()
                if mid:
                    dedup[mid] = m
        title_str = ", ".join(titles[:3]) + (f" (+{len(titles) - 3})" if len(titles) > 3 else "")
        self._open_briefings_dialog(
            selection_payload={
                "title": title_str or f"{len(selected)} Selected Items",
                "messages": list(dedup.values()),
                "focus_prompt": "",
            },
            auto_run=False,
        )

    def _on_append_summary_to_notes(self, block: str) -> None:
        if not block or not hasattr(self, "note_edit"):
            return
        curr = self.note_edit.toPlainText().strip()
        merged = f"{curr}\n\n{block}" if curr else block
        if hasattr(self, "note_frame"):
            self.note_frame.setVisible(True)
        if hasattr(self, "btn_toggle_note"):
            self.btn_toggle_note.setChecked(True)
        self.note_edit.setPlainText(merged)
        self._note_dirty = True
        self._flush_pending_note_save(reload_view=False)
        notify_success(self, "Saved to Notes", "Appended AI summary to delegate notes.")

    def _on_inline_summary_saved(self, _record) -> None:
        if hasattr(self, "_refresh_summary_freshness_cache"):
            self._refresh_summary_freshness_cache(self._active_dataset())
        # Refresh current table badges without losing active reader state
        for r, it in enumerate(getattr(self, "_visible_items", [])):
            if it.get("_is_group"):
                st_txt, st_tip = self._format_row_state_cell_text_and_tip(it)
                cell = self.table.item(r, 0)
                if cell is not None:
                    cell.setText(st_txt)
                    if st_tip:
                        cell.setToolTip(st_tip)

    def _on_llm_export_clicked(self):
        selected = list(getattr(self, "_selected_items", []) or [])
        if len(selected) > 1:
            self._summarize_selected_items(selected)
            return
        it = self._current_item
        if not it:
            return
        menu = QMenu(self)
        menu.setToolTipsVisible(True)
        spec = self.summary_card._active_spec() if hasattr(self, "summary_card") else resolve_effective_model_spec("leaf")
        fresh = self.summary_card.current_freshness() if hasattr(self, "summary_card") else None
        if fresh and fresh.has_summary and fresh.new_count > 0:
            sum_lbl = f"⚡ Update Summary (+{fresh.new_count}) [{spec.short_badge()}]"
        elif fresh and fresh.has_summary:
            sum_lbl = f"🔄 Refresh Summary [{spec.short_badge()}]"
        else:
            sum_lbl = f"✨ Summarize Discussion [{spec.short_badge()}]"
        act_ai_sum = menu.addAction(sum_lbl)
        act_rebuild = None
        if fresh and fresh.has_summary:
            act_rebuild = menu.addAction("🔄 Full Rebuild Summary")

        model_sub = menu.addMenu("⚙️ Summary Model")
        populate_summary_model_menu(
            model_sub,
            role="leaf",
            active_spec=spec,
            on_selected=lambda s: setattr(self.summary_card, "_model_override", s) if hasattr(self, "summary_card") else None,
            parent=self,
        )
        menu.addSeparator()
        act_copy_md = menu.addAction("📋 Copy Pure Markdown (for ChatGPT / Claude / Codex)")
        chosen = menu.exec_(QCursor.pos())
        if chosen == act_copy_md:
            md = self._build_current_item_pure_markdown()
            if md:
                QApplication.clipboard().setText(md)
                notify_success(self, "Markdown Copied", "Copied clean thread Markdown to clipboard.")
        elif chosen == act_ai_sum and hasattr(self, "summary_card"):
            self.summary_card.trigger_summarize(mode="auto")
        elif act_rebuild is not None and chosen == act_rebuild and hasattr(self, "summary_card"):
            self.summary_card.trigger_summarize(mode="full")

    def _on_generate_email_statistics(self):
        if not self._all_items:
            QMessageBox.information(
                self, "No Emails", "There are no indexed emails in this meeting to analyze."
            )
            return
        gdb_path = getattr(self.general_email_db, "db_path", None) if self.general_email_db is not None else None
        tdb_path = getattr(self.tracked_db, "db_path", None) if self.tracked_db is not None else None
        base_db_path = gdb_path or tdb_path
        if not base_db_path:
            return
        meeting_dir = Path(base_db_path).parent.parent
        from modules.emails.core.stats.dataset_builder import build_unified_email_stats_dataset
        from modules.emails.core.stats.statistics_threads import EmailStatsExporterThread
        from modules.meetings.ui.tdocs_dialogs import StatisticsSettingsDialog

        dataset = build_unified_email_stats_dataset(
            all_inbox_items=self._all_items,
            tdoc_ai_lookup=self.tdoc_ai_lookup,
            tdoc_meta_map=self._tdoc_meta_map,
        )
        cfg = StatisticsSettingsDialog().load_config()
        if hasattr(self, "btn_stats"):
            self.btn_stats.setEnabled(False)
            self.btn_stats.setText("⏳ Stats...")
        self._stats_thread = EmailStatsExporterThread(
            meeting_dir,
            dataset,
            self.meeting_id or meeting_dir.name,
            cfg,
        )
        self._stats_thread.finished.connect(self._on_email_stats_finished)
        self._stats_thread.start()

    def _on_email_stats_finished(self, success: bool, msg: str):
        if hasattr(self, "btn_stats"):
            self.btn_stats.setEnabled(True)
            self.btn_stats.setText("📊 Stats")
        if success:
            notify_success(self, "Email Analytics Ready", "Opened HTML Email Statistics Dashboard.")
            if hasattr(os, "startfile"):
                os.startfile(msg)
            else:
                webbrowser.open(f"file:///{msg}")
        else:
            QMessageBox.warning(
                self, "Email Statistics Failed", f"Could not generate email statistics:\n{msg}"
            )

    def _on_open_context(self):
        if not self._current_item:
            return
        it = self._current_item
        is_tdoc_bucket = it.get("_group_kind") == "TDOC" and bool(it.get("_primary_tdoc"))
        if is_tdoc_bucket and self._open_tdoc_callback:
            primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
            sig = it.get("_signals") or {}
            rev_tokens = sig.get("rev_tokens") or []
            latest_rev_tok = next(
                (rt for rt in reversed(rev_tokens) if rt.upper().startswith(primary_tdoc)),
                primary_tdoc,
            )
            self._open_tdoc_callback(latest_rev_tok)
            return
        if it.get("thread_id") and self._open_thread_callback:
            self._open_thread_callback(it["thread_id"])
            self.reload_from_dbs(preserve_id=it["id"])
        elif it.get("linked_tdocs") and self._open_tdoc_callback:
            self._open_tdoc_callback(it["linked_tdocs"][0])
            self.reload_from_dbs(preserve_id=it["id"])

    def _on_anchor_clicked(self, url: QUrl):
        scheme = url.scheme().lower()
        if scheme == "jump":
            path_parts = [p for p in url.path().strip("/").split("/") if p]
            raw_idx = path_parts[0] if path_parts else url.host()
            try:
                msg_idx = int(raw_idx)
                blk = find_heading_block_for_message_index(self.preview_browser, msg_idx)
                if blk is not None:
                    scroll_browser_to_block(self.preview_browser, blk)
            except Exception:
                pass
            return

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
                    QMessageBox.warning(
                        self, "Outlook Message", "Could not open the selected message in Microsoft Outlook."
                    )
            elif host == "reply":
                if not reply_to_outlook_message(msg_id, reply_all=True):
                    QMessageBox.warning(
                        self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook."
                    )
            elif (
                host == "reload"
                and self._current_item
                and self._current_item.get("thread_id")
                and self.tracked_db
            ):
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
                cache_dir = (
                    Path(self.tracked_db.db_path).parent / "InboxAttachments"
                    if self.tracked_db
                    else Path.home() / ".3gpp_tools_inbox"
                )
                cache_dir.mkdir(parents=True, exist_ok=True)
                ok = open_outlook_attachment(msg_id, att_name, cache_dir)
                if not ok:
                    QMessageBox.warning(
                        self,
                        "Attachment",
                        f"Could not extract attachment '{att_name}' from Outlook.",
                    )
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
