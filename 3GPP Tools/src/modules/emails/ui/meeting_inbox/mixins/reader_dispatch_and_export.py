"""Outlook reply/open, LLM export, statistics dashboard, and reader anchor dispatch mixin."""

from __future__ import annotations

import os
from pathlib import Path
import urllib.parse
import webbrowser

from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QCursor, QDesktopServices
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox

from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.emeeting_signals import extract_base_tdoc_id
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

    def _on_llm_export_clicked(self):
        it = self._current_item
        if not it:
            return
        menu = QMenu(self)
        act_copy_md = menu.addAction("📋 Copy Pure Markdown (for ChatGPT / Claude / Codex)")
        act_ai_sum = menu.addAction("🤖 Summarize Discussion with Local AI (Ollama)...")
        chosen = menu.exec_(QCursor.pos())
        if chosen == act_copy_md:
            md = self._build_current_item_pure_markdown()
            if md:
                QApplication.clipboard().setText(md)
                notify_success(self, "Markdown Copied", "Copied clean thread Markdown to clipboard.")
        elif chosen == act_ai_sum:
            primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
            if not primary_tdoc and it.get("linked_tdocs"):
                primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
            label = primary_tdoc or str(it.get("thread_title") or it.get("subject") or "Thread")
            raw_msgs = it.get("_group_messages") if it.get("_is_group") else [it]
            emails_payload = [
                {
                    "id": m.get("id", ""),
                    "subject": m.get("subject", ""),
                    "sender_name": m.get("sender_name", ""),
                    "company": m.get("company", ""),
                    "date_received": m.get("date_str", ""),
                    "body_text": m.get("clean_body") or m.get("body_text", ""),
                }
                for m in (raw_msgs or [])
            ]
            from modules.emails.ui.email_summary_dialog import EmailSummaryDialog

            dlg = EmailSummaryDialog(
                tdoc_id=label,
                emails_data=emails_payload,
                wg=self.wg,
                save_callback=lambda tid, st, notes: self._sync_tdoc_user_data(
                    tid, new_status=st, new_notes=notes
                ),
                parent=None,
            )
            self._summary_dlg = dlg
            dlg.show()
            dlg.raise_()
            dlg.activateWindow()

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
