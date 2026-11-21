"""Outlook sync, document caching, LLM CLI launch, and anchor handling mixin."""

import os
import sys
import urllib.parse
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

from core.ai.cli_detector import detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import (
    launch_cli_in_terminal,
    open_mcp_agent_setup_dialog,
    write_cli_context_file,
)
from core.ui.toast import notify_info, notify_success, notify_toast
from modules.emails.core.general_email_db import (
    reload_general_email_from_outlook,
    reload_tdoc_emails_from_outlook,
)
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    extract_thread_document_urls,
    is_cacheable_document_url,
)
from modules.emails.core.tracked_threads_service import (
    cache_and_open_thread_document_url,
    normalize_subject,
    open_outlook_attachment,
    open_outlook_message,
    reply_to_outlook_message,
)
from modules.emails.ui.email_summary_dialog import EmailSummaryDialog


def _call_launch_cli_in_terminal(*args, **kwargs):
    """Forward through tdoc_emails_dialog shim if monkeypatched in tests."""
    shim_mod = sys.modules.get("modules.emails.ui.tdoc_emails_dialog")
    launcher = getattr(shim_mod, "launch_cli_in_terminal", launch_cli_in_terminal) if shim_mod else launch_cli_in_terminal
    return launcher(*args, **kwargs)


class TDocEmailsActionsMixin:
    """Handles Outlook actions, document caching, Markdown/LLM exports, and links."""

    def _get_cache_dir(self) -> Path:
        d = self.db.db_path.parent / "TDocEmailCache" / self.target_tdoc
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _on_reload_all_from_outlook(self):
        eids = [str(e.get("id") or "") for e in self.emails if e.get("id") and not str(e.get("id")).startswith("MSG-")]
        if not eids:
            notify_info(self, "Reload Emails", "No Outlook-backed emails found to reload.")
            return
        sb = self.browser.verticalScrollBar()
        prev_scroll = sb.value()
        updated = reload_tdoc_emails_from_outlook(self.db, eids)
        self._load_emails()
        sb.setValue(min(prev_scroll, sb.maximum()))
        self.data_changed.emit()
        if updated > 0:
            notify_success(self, "Emails Reloaded", f"Re-parsed {updated} Outlook email(s) (tables, links & attachments).")
        else:
            notify_toast(
                self,
                "Reload Warning",
                "Could not reload emails from Outlook (check Outlook connection or moved items).",
                variant="warning",
            )

    def _mark_all_family_read(self):
        query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
        self.db.set_tdocs_read_status(query_set, True, sync_outlook=True)
        self.data_changed.emit()
        self._load_emails()

    def _on_open_tdoc_clicked(self):
        if self._open_tdoc_callback and self.target_tdoc:
            self._open_tdoc_callback(self.target_tdoc)

    def _on_pin_as_tracked_thread(self):
        if not self._pin_thread_callback or not self.emails:
            return
        latest = self.emails[-1]
        self._pin_thread_callback(
            {
                "subject": latest.get("subject") or f"Discussion on {self.target_tdoc}",
                "normalized_subject": normalize_subject(latest.get("subject") or ""),
                "linked_tdocs": list(self.family_tdocs),
            }
        )

    def _on_cache_all_docs(self):
        cache_dir = self._get_cache_dir()
        cached: List[Path] = []
        for doc_url in extract_thread_document_urls(self.emails):
            p = cache_and_open_thread_document_url(doc_url, cache_dir, open_after=False)
            if p:
                cached.append(p)
        ns = OutlookClient.get_namespace()
        if ns:
            for m in self.emails:
                eid = str(m.get("id") or "").strip()
                if not eid or eid.startswith("MSG-") or not m.get("attachments"):
                    continue
                try:
                    item = ns.GetItemFromID(eid)
                    atts = getattr(item, "Attachments", None)
                    if not atts:
                        continue
                    for i in range(1, atts.Count + 1):
                        att = atts.Item(i)
                        fname = getattr(att, "FileName", "") or ""
                        if not fname or fname.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                            continue
                        safe_name = "".join(c for c in fname if c not in '<>:"/\\|?*').strip()
                        if safe_name:
                            out_path = cache_dir / safe_name
                            if not out_path.exists():
                                att.SaveAsFile(str(out_path.resolve()))
                            cached.append(out_path)
                except Exception:
                    continue
        QMessageBox.information(
            self,
            "Discussion Documents Cached",
            f"Cached {len(cached)} document(s)/attachment(s) into:\n{cache_dir}",
        )

    def _on_open_cache_folder(self):
        folder = self._get_cache_dir()
        if hasattr(os, "startfile"):
            os.startfile(str(folder.resolve()))

    def _populate_llm_export_menu(self):
        menu = self.llm_export_menu
        menu.clear()
        menu.setToolTipsVisible(True)

        act_ai_sum = menu.addAction("🤖 AI Summarize Discussion (Ollama)...")
        act_ai_sum.setToolTip("Generate a structured stance summary using local Ollama and optionally append it to TDoc My Notes.")
        act_ai_sum.triggered.connect(self._on_ai_summarize_clicked)

        menu.addSeparator()
        try:
            statuses = detect_cli_agents()
        except Exception:
            statuses = {}
        installed = [st for st in statuses.values() if st.installed]
        if installed:
            for st in installed:
                badge = format_active_model_badge(st.agent_id)
                act = menu.addAction(f"{st.icon} Launch {st.label} [{badge}]")
                act.setToolTip(
                    f"Prepare discussion Markdown + cached docs and open {st.label} idle "
                    "(prompt copied to clipboard for Ctrl+V — zero auto token burn)."
                )
                act.triggered.connect(lambda _=False, aid=st.agent_id: self._on_launch_cli(aid))
        else:
            no_cli = menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
            no_cli.setEnabled(False)

        menu.addSeparator()
        act_copy = menu.addAction("📋 Copy Discussion Markdown")
        act_copy.triggered.connect(self._on_copy_markdown)
        act_save = menu.addAction("💾 Save Discussion .md...")
        act_save.triggered.connect(self._on_save_markdown)
        menu.addSeparator()
        menu.addAction(self.export_chrono_check)
        menu.addSeparator()
        act_cfg = menu.addAction("⚙️ Configure CLI Agents && MCP...")
        act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(self))

    def _get_export_newest_first(self) -> bool:
        if hasattr(self, "export_chrono_check") and self.export_chrono_check.isChecked():
            return False
        return bool(self.order_combo.currentData()) if hasattr(self, "order_combo") else False

    def _get_markdown(self) -> str:
        return export_thread_to_markdown(
            self._build_synthetic_thread_dict(),
            self.emails,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=False,
            newest_first=self._get_export_newest_first(),
            include_metadata=True,
        )

    def _on_copy_markdown(self):
        QApplication.clipboard().setText(self._get_markdown())
        notify_success(self, "Copied", "Discussion Markdown copied to clipboard.")

    def _on_save_markdown(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc Emails Markdown",
            f"{self.target_tdoc}_emails.md",
            "Markdown (*.md)",
        )
        if path:
            Path(path).write_text(self._get_markdown(), encoding="utf-8")

    def _on_launch_cli(self, agent_id: Optional[str] = None):
        cli = (agent_id or "codex").strip().lower()
        cache_dir = self._get_cache_dir()
        ws_dir = cache_dir / "llm_workspace"
        ws_dir.mkdir(parents=True, exist_ok=True)
        md_path = ws_dir / f"{self.target_tdoc}_emails.md"
        md_path.write_text(self._get_markdown(), encoding="utf-8")

        doc_mds: List[Path] = []
        for existing in cache_dir.iterdir():
            if existing.is_file() and existing.suffix.lower() in (".docx", ".docm"):
                out_md = ws_dir / f"{existing.stem}.md"
                if not out_md.exists():
                    try:
                        from modules.meetings.core.tdoc_markdown_converter import convert_document_to_markdown
                        convert_document_to_markdown(str(existing), str(out_md))
                    except Exception:
                        pass
                if out_md.exists():
                    doc_mds.append(out_md)

        extra_docs = ""
        if doc_mds:
            extra_docs = f" Companion converted documents are also in {ws_dir} ({', '.join(p.name for p in doc_mds[:10])})."
        prompt = (
            f"Analyze the 3GPP email discussion referencing {self.target_tdoc} in {md_path} and summarize "
            f"company positions, key technical objections, revisions, and open points.{extra_docs}"
        )
        write_cli_context_file(
            ws_dir,
            task_mode="custom",
            custom_question=prompt,
            markdown_files=[md_path] + doc_mds,
            briefing_title=f"TDoc Related Emails: {self.target_tdoc}",
        )
        _call_launch_cli_in_terminal(
            cli,
            ws_dir,
            bootstrap_prompt=prompt,
            parent=self,
            auto_execute=False,
        )

    def _on_ai_summarize_clicked(self):
        target_emails = list(self._filtered_emails or self.emails)
        if not target_emails:
            QMessageBox.warning(self, "No Emails", "There are no emails to summarize for this document.")
            return
        if self._summary_dialog and self._summary_dialog.isVisible():
            self._summary_dialog.setWindowState(
                self._summary_dialog.windowState() & ~Qt.WindowMinimized | Qt.WindowActive
            )
            self._summary_dialog.raise_()
            self._summary_dialog.activateWindow()
            return
        self._summary_dialog = EmailSummaryDialog(
            tdoc_id=self.target_tdoc,
            family_tdocs=self.family_tdocs,
            emails_data=target_emails,
            save_callback=self.save_callback,
            parent=None,
        )
        self._summary_dialog.show()

    def _on_anchor_clicked(self, url: QUrl):
        scheme = url.scheme().lower()
        if scheme == "tdoc":
            target_tdoc = url.path().strip() or url.toString().split(":", 1)[1].strip()
            if target_tdoc:
                self.tdoc_selected.emit(target_tdoc.upper())
            return

        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if not path_parts:
                return
            msg_id = path_parts[0]
            sb = self.browser.verticalScrollBar()
            prev_scroll = sb.value()

            if host == "msg":
                if not open_outlook_message(msg_id):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply":
                if not reply_to_outlook_message(msg_id, reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
            elif host == "reload":
                if reload_general_email_from_outlook(self.db, msg_id):
                    self._load_emails()
                    sb.setValue(min(prev_scroll, sb.maximum()))
                    self.data_changed.emit()
                    notify_success(self, "Email Reloaded", "Message body, tables, links, and attachments re-loaded from Outlook.")
                else:
                    notify_toast(self, "Reload Failed", "Could not re-fetch this message from Outlook.", variant="warning")
            elif host in ("read", "unread"):
                self.db.set_emails_read_status([msg_id], is_read=(host == "read"), sync_outlook=True)
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host in ("flag", "unflag"):
                self.db.set_emails_flag_status([msg_id], is_flagged=(host == "flag"), sync_outlook=True)
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host in ("ignore", "unignore"):
                self.db.set_emails_ignored_status([msg_id], is_ignored=(host == "ignore"))
                self._load_emails()
                sb.setValue(min(prev_scroll, sb.maximum()))
                self.data_changed.emit()
            elif host == "copy":
                msg = next((m for m in self.emails if str(m.get("id") or "") == msg_id), None)
                if msg:
                    body = (msg.get("clean_body") if self.dedup_check.isChecked() and msg.get("clean_body") else msg.get("body_text")) or ""
                    clip_txt = f"[{msg.get('date_received', '')}] {msg.get('sender_name', '')} ({msg.get('company', '')}):\n{body.strip()}"
                    QApplication.clipboard().setText(clip_txt)
                    notify_info(self, "Copied", "Message copied to clipboard.")
            elif host == "att" and len(path_parts) >= 2:
                att_name = "/".join(path_parts[1:])
                if not open_outlook_attachment(msg_id, att_name, save_dir=self._get_cache_dir()):
                    QMessageBox.warning(self, "Attachment", f"Could not open attachment '{att_name}' from Outlook.")
            return

        url_str = url.toString()
        if is_cacheable_document_url(url_str):
            cached_path = cache_and_open_thread_document_url(url_str, self._get_cache_dir(), open_after=True)
            if cached_path:
                notify_info(self, "3GPP Document Opened", f"Cached and opened {cached_path.name}.")
                return
            notify_toast(
                self,
                "Direct Download Unavailable",
                "Could not download automatically — opening link in browser.",
                variant="warning",
            )
        QDesktopServices.openUrl(url)
