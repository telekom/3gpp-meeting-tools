"""Anchor dispatch, Outlook refresh workers, merge/unpin, export, and CLI launch mixin."""

from pathlib import Path
import sys
from typing import Optional
import urllib.parse

from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox

from core.ai import cli_launcher as _cli_launcher
from core.ai.cli_launcher import (
    launch_cli_in_terminal as _ORIG_LAUNCH_CLI_IN_TERMINAL,
    write_cli_context_file,
)
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    is_cacheable_document_url,
)
from modules.emails.core.tracked_threads_service import (
    cache_all_thread_documents_and_attachments,
    cache_and_open_thread_document_url,
    mark_thread_read_with_outlook,
    open_outlook_attachment,
    open_outlook_message,
    prepare_thread_llm_workspace,
    reload_outlook_message,
    reply_to_outlook_message,
    set_thread_message_flag_with_outlook,
    set_thread_message_read_with_outlook,
)
from modules.emails.core.tracked_threads_workers import SingleThreadRefreshWorker


def call_launch_cli_in_terminal(*args, **kwargs):
    """Forward calls to tracked_thread_dialog.launch_cli_in_terminal when monkeypatched in tests."""
    for mod_name in (
        "modules.emails.ui.tracked_thread_dialog",
        "src.modules.emails.ui.tracked_thread_dialog",
    ):
        mod = sys.modules.get(mod_name)
        if mod is not None:
            patched = getattr(mod, "launch_cli_in_terminal", None)
            if patched is not None and patched is not _ORIG_LAUNCH_CLI_IN_TERMINAL:
                return patched(*args, **kwargs)
    return _cli_launcher.launch_cli_in_terminal(*args, **kwargs)


class TrackedThreadActionsMixin:
    """Reader link handlers, Outlook sync actions, thread merge/delete, and LLM export."""

    def _get_thread_cache_dir(self) -> Path:
        d = self.db.db_path.parent / "ThreadAttachments" / self.thread_id
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _on_cache_all_docs(self):
        cached = cache_all_thread_documents_and_attachments(self.db, self.thread_id, self._get_thread_cache_dir())
        QMessageBox.information(
            self,
            "Thread Documents Cached",
            f"Cached {len(cached)} document(s)/attachment(s) into:\n{self._get_thread_cache_dir()}",
        )

    def _on_open_thread_folder(self):
        import os
        folder = self._get_thread_cache_dir()
        if hasattr(os, "startfile"):
            os.startfile(str(folder.resolve()))

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
            if host == "msg" and path_parts:
                if not open_outlook_message(path_parts[0]):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply" and path_parts:
                if not reply_to_outlook_message(path_parts[0], reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
            elif host == "reload" and path_parts:
                msg_id = path_parts[0]
                sb = self.browser.verticalScrollBar()
                prev_scroll = sb.value()
                if reload_outlook_message(self.db, self.thread_id, msg_id):
                    self._load_data()
                    sb.setValue(min(prev_scroll, sb.maximum()))
                    self.thread_changed.emit()
                    notify_success(self, "Email Reloaded", "Message body, tables, and links re-loaded from Outlook.")
                else:
                    notify_toast(self, "Reload Failed", "Could not re-fetch this message from Outlook (it may have been moved or deleted).", variant="warning")
            elif host in ("read", "unread", "flag", "unflag") and path_parts:
                msg_id = path_parts[0]
                sb = self.browser.verticalScrollBar()
                prev_scroll = sb.value()
                if host in ("read", "unread"):
                    set_thread_message_read_with_outlook(self.db, msg_id, is_read=(host == "read"), sync_outlook=True)
                else:
                    set_thread_message_flag_with_outlook(self.db, msg_id, is_flagged=(host == "flag"), sync_outlook=True)
                self._load_data()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.thread_changed.emit()
            elif host == "copy" and path_parts:
                target_id = path_parts[0]
                msgs = self.db.get_thread_messages(self.thread_id)
                msg = next((m for m in msgs if str(m.get("id") or "") == target_id), None)
                if msg:
                    body = (msg.get("clean_body") if self.dedup_check.isChecked() and msg.get("clean_body") else msg.get("body_text")) or ""
                    clip_txt = f"[{msg.get('date_sent', '')}] {msg.get('sender_name', '')} ({msg.get('company', '')}):\n{body.strip()}"
                    QApplication.clipboard().setText(clip_txt)
            elif host == "att" and len(path_parts) >= 2:
                out_dir = self._get_thread_cache_dir()
                if not open_outlook_attachment(path_parts[0], path_parts[1], save_dir=out_dir):
                    QMessageBox.warning(self, "Attachment", f"Could not open attachment '{path_parts[1]}' from Outlook.")
            return
        url_str = url.toString()
        if is_cacheable_document_url(url_str):
            cached_path = cache_and_open_thread_document_url(url_str, self._get_thread_cache_dir(), open_after=True)
            if cached_path:
                return
        QDesktopServices.openUrl(url)


    def _on_refresh(self):
        self._start_thread_worker(reload_cached=False)

    def _on_reload_all_cached(self):
        self._start_thread_worker(reload_cached=True)

    def _on_reload_from_msg_file(self):
        from modules.emails.core.tracked_threads_service import parse_msg_file

        paths, _ = QFileDialog.getOpenFileNames(self, "Select Outlook .msg File(s) to Reload/Merge", "", "Outlook Message (*.msg)")
        if not paths:
            return
        loaded = []
        for p in paths:
            parsed = parse_msg_file(Path(p))
            if parsed:
                loaded.append(parsed)
        if loaded:
            self.db.upsert_messages(self.thread_id, loaded)
            self._load_data()
            self.thread_changed.emit()
            notify_success(self, "Loaded .msg File(s)", f"Updated {len(loaded)} message(s) in this thread.")
        else:
            notify_toast(self, "Load Failed", "Could not parse selected .msg file(s).", variant="warning")

    def _start_thread_worker(self, reload_cached: bool = False):
        if self._refresh_worker and self._refresh_worker.isRunning():
            return
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setText("⏳ Reloading..." if reload_cached else "⏳ Syncing...")
        if reload_cached:
            msgs = [m for m in self.db.get_thread_messages(self.thread_id) if m.get("id") and not str(m.get("id")).startswith("MSG-")]
            self.sync_bar.start_busy("Reloading cached Outlook emails...", total=len(msgs))
        else:
            thread = self.db.get_thread(self.thread_id) or {}
            folders = thread.get("outlook_folders") or []
            self.sync_bar.start_busy("Scanning Outlook folders...", total=len(folders), remaining=folders)

        worker = SingleThreadRefreshWorker(self.db.db_path, self.thread_id, reload_cached=reload_cached, parent=self)
        worker.progress.connect(self.sync_bar.update_progress)
        worker.finished_sync.connect(lambda ok, cnt, err: self._on_refresh_finished(ok, cnt, err, reload_cached=reload_cached))
        self._refresh_worker = worker
        worker.start()

    def _on_cancel_refresh(self):
        if self._refresh_worker and self._refresh_worker.isRunning():
            self._refresh_worker.request_cancel()

    def _on_refresh_finished(self, ok: bool, count: int, err: str, reload_cached: bool = False):
        self.refresh_btn.setEnabled(True)
        self.refresh_btn.setText("🔄 Refresh")
        self._load_data()
        self.thread_changed.emit()
        if ok:
            self.sync_bar.finish_ok()
            if reload_cached:
                notify_success(self, "Emails Reloaded", f"Re-parsed {count} cached Outlook message(s).")
            elif count > 0:
                notify_success(self, "Thread Refreshed", f"{count} new message(s) synced.")
            else:
                notify_info(self, "Thread Up-to-Date", "No new replies found in Outlook.")
        else:
            self.sync_bar.show_error(err or "Thread sync did not finish cleanly.")
            notify_toast(self, "Thread Sync", err or "Sync warning", variant="warning")

    def _on_mark_read(self):
        mark_thread_read_with_outlook(self.db, self.thread_id, sync_outlook=True)
        self._load_data()
        self.thread_changed.emit()

    def _on_merge_thread(self):
        current = self.db.get_thread(self.thread_id)
        if not current:
            return
        candidates = [
            t for t in self.db.get_threads_for_meeting(current.get("meeting_id", ""))
            if t["thread_id"] != self.thread_id
        ]
        if not candidates:
            notify_info(self, "No Other Threads", "No other tracked threads in this meeting to merge.")
            return
        from modules.emails.ui.merge_threads_dialog import MergeThreadsDialog

        dlg = MergeThreadsDialog(target_thread=current, candidates=candidates, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            src_ids = dlg.selected_source_ids()
            if src_ids:
                self.db.merge_threads(self.thread_id, src_ids)
                self._load_data()
                self.thread_changed.emit()
                notify_success(
                    self,
                    "Threads Merged",
                    f"Merged {len(src_ids)} split thread(s) into {self.thread_id}.",
                )

    def _on_delete_thread(self):
        if QMessageBox.question(
            self,
            "Unpin Email Thread",
            "Remove this tracked email thread from the meeting?",
            QMessageBox.Yes | QMessageBox.No,
        ) == QMessageBox.Yes:
            self.db.delete_thread(self.thread_id)
            self.thread_changed.emit()
            self.accept()

    def _get_export_newest_first(self) -> bool:
        if hasattr(self, "export_chrono_check") and self.export_chrono_check.isChecked():
            return False
        return bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else False

    def _get_markdown(self) -> str:
        thread = self.db.get_thread(self.thread_id) or {}
        messages = self.db.get_thread_messages(self.thread_id)
        return export_thread_to_markdown(
            thread,
            messages,
            deduplicate_quotes=self.dedup_check.isChecked(),
            newest_first=self._get_export_newest_first(),
        )

    def _on_copy_markdown(self):
        QApplication.clipboard().setText(self._get_markdown())
        QMessageBox.information(self, "Copied", "Thread Markdown copied to clipboard.")

    def _on_save_markdown(self):
        thread = self.db.get_thread(self.thread_id) or {}
        safe_name = "".join(c for c in thread.get("title", "thread") if c.isalnum() or c in " -_").strip()[:60]
        path, _ = QFileDialog.getSaveFileName(self, "Save Thread Markdown", f"{safe_name or 'thread'}.md", "Markdown (*.md)")
        if path:
            Path(path).write_text(self._get_markdown(), encoding="utf-8")

    def _on_launch_cli(self, agent_id: Optional[str] = None):
        cli = (agent_id or "codex").strip().lower()
        prep = prepare_thread_llm_workspace(
            self.db,
            self.thread_id,
            self._get_thread_cache_dir(),
            deduplicate_quotes=self.dedup_check.isChecked(),
            newest_first=self._get_export_newest_first(),
            cache_linked_docs=True,
        )
        md_path = prep["thread_md"]
        ws_dir = prep["workspace_dir"]
        doc_mds = prep["doc_mds"]
        extra_docs = ""
        if doc_mds:
            names = ", ".join(p.name for p in doc_mds[:10])
            extra_docs = f" Companion converted draft documents are also in {ws_dir} ({names})."
        prompt = (
            f"Analyze the 3GPP email thread in {md_path} and summarize the views exchanged by company, "
            f"key disagreements, and open points.{extra_docs}"
        )
        thread = self.db.get_thread(self.thread_id) or {}
        write_cli_context_file(
            ws_dir,
            task_mode="custom",
            custom_question=prompt,
            markdown_files=[md_path] + list(doc_mds),
            briefing_title=f"Tracked Email Thread: {thread.get('title', self.thread_id)}",
            user_notes_override=str(thread.get("user_notes") or ""),
        )
        call_launch_cli_in_terminal(
            cli,
            ws_dir,
            bootstrap_prompt=prompt,
            parent=self,
            auto_execute=False,
        )

