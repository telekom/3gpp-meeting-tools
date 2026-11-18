# --- File: src/modules/emails/ui/tracked_thread_dialog.py ---
import subprocess
from pathlib import Path
from typing import Callable, List, Optional
import urllib.parse

from PyQt5.QtCore import QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    is_cacheable_document_url,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    cache_all_thread_documents_and_attachments,
    cache_and_open_thread_document_url,
    list_recent_threads_in_folders,
    open_outlook_attachment,
    open_outlook_message,
    pin_thread_by_subject_or_cid,
    prepare_thread_llm_workspace,
    refresh_thread,
    reply_to_outlook_message,
)
from modules.meetings.ui.topic_pills import TagCloudWidget


class TrackedThreadDialog(QDialog):
    """Integrated viewer for a tracked email thread with Markdown export and LLM CLI launch."""

    thread_changed = pyqtSignal()

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        thread_id: str,
        parent: Optional[QWidget] = None,
        ai_suggestions: Optional[List[str]] = None,
        wi_suggestions: Optional[List[str]] = None,
        tdoc_suggestions: Optional[List[str]] = None,
        tdoc_family_expander: Optional[Callable[[List[str]], List[str]]] = None,
    ):
        super().__init__(
            parent,
            Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint,
        )
        self.db = db
        self.thread_id = thread_id
        self._ai_suggestions = ai_suggestions or []
        self._wi_suggestions = wi_suggestions or []
        self._tdoc_suggestions = tdoc_suggestions or []
        self._tdoc_family_expander = tdoc_family_expander
        self._loading = False
        self._settings = QSettings("3GPP_Tools", "TrackedThreadDialog")
        self.setWindowTitle("Tracked Email Thread")
        self.setMinimumSize(520, 420)
        self.resize(1020, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        self.title_label = QLabel("Email Thread")
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)

        self.meta_label = QLabel("")
        self.meta_label.setStyleSheet("color: #666;")
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

        # Links pill editors (auto-save + autocomplete)
        links_grid = QGridLayout()
        links_grid.setSpacing(6)

        links_grid.addWidget(QLabel("<b>AIs:</b>"), 0, 0, Qt.AlignTop)
        self.ais_pills = TagCloudWidget(
            placeholder="Type AI # or topic (e.g. 20.6.1)...",
            pill_colors={"bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
        )
        self.ais_pills.set_suggestions(self._ai_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        self.ais_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.ais_pills, 0, 1)

        links_grid.addWidget(QLabel("<b>WIs:</b>"), 1, 0, Qt.AlignTop)
        self.wis_pills = TagCloudWidget(
            placeholder="Type Work Item (e.g. FS_6G_ARC)...",
            pill_colors={"bg": "#EEF2FF", "fg": "#4338CA", "border": "#C7D2FE"},
        )
        self.wis_pills.set_suggestions(self._wi_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        self.wis_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.wis_pills, 1, 1)

        links_grid.addWidget(QLabel("<b>TDocs:</b>"), 2, 0, Qt.AlignTop)
        self.tdocs_pills = TagCloudWidget(
            placeholder="Type TDoc # (e.g. S2-2612673) — revisions & merges auto-included...",
            pill_colors={"bg": "#ECFDF5", "fg": "#047857", "border": "#A7F3D0"},
        )
        self.tdocs_pills.set_suggestions(self._tdoc_suggestions, value_extractor=lambda s: s.split(" — ", 1)[0].strip().upper())
        self.tdocs_pills.tagsChanged.connect(self._on_links_changed)
        links_grid.addWidget(self.tdocs_pills, 2, 1)

        layout.addLayout(links_grid)

        # Responsive wrapping action toolbar
        toolbar_widget = QWidget(self)
        toolbar_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        toolbar_sp.setHeightForWidth(True)
        toolbar_widget.setSizePolicy(toolbar_sp)
        toolbar = FlowLayout(toolbar_widget, margin=0, h_spacing=6, v_spacing=6)

        self.order_combo = QComboBox()
        self.order_combo.setToolTip("Choose display order of messages within this thread.")
        self.order_combo.addItem("⬇️ Newest First", True)
        self.order_combo.addItem("⬆️ Oldest First", False)
        saved_newest = self._settings.value("newest_first", True, type=bool)
        self.order_combo.setCurrentIndex(0 if saved_newest else 1)
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        toolbar.addWidget(self.order_combo)

        self.export_chrono_check = QCheckBox("Export Oldest First (for LLMs)")
        self.export_chrono_check.setToolTip("When checked, Copy/Save Markdown and LLM CLI export chronologically (Oldest First) even when viewing Newest First.")
        self.export_chrono_check.setChecked(self._settings.value("export_oldest_first", True, type=bool))
        self.export_chrono_check.toggled.connect(lambda v: self._settings.setValue("export_oldest_first", bool(v)))
        toolbar.addWidget(self.export_chrono_check)

        self.dedup_check = QCheckBox("Deduplicate trailing quotes")
        self.dedup_check.setChecked(True)
        self.dedup_check.toggled.connect(self._render_messages)
        toolbar.addWidget(self.dedup_check)

        refresh_btn = QPushButton("Refresh Now")
        refresh_btn.clicked.connect(self._on_refresh)
        toolbar.addWidget(refresh_btn)

        read_btn = QPushButton("Mark All Read")
        read_btn.clicked.connect(self._on_mark_read)
        toolbar.addWidget(read_btn)

        merge_btn = QPushButton("Merge Split Thread...")
        merge_btn.clicked.connect(self._on_merge_thread)
        toolbar.addWidget(merge_btn)

        del_btn = QPushButton("Unpin")
        del_btn.clicked.connect(self._on_delete_thread)
        toolbar.addWidget(del_btn)

        cache_docs_btn = QPushButton("📥 Cache Docs")
        cache_docs_btn.setToolTip("Download and cache all documents (e.g. 10.10.10.10 / 3GPP FTP) and attachments linked in this thread.")
        cache_docs_btn.clicked.connect(self._on_cache_all_docs)
        toolbar.addWidget(cache_docs_btn)

        folder_btn = QPushButton("📂 Thread Folder")
        folder_btn.setToolTip("Open the local cache directory for this email thread's documents and attachments.")
        folder_btn.clicked.connect(self._on_open_thread_folder)
        toolbar.addWidget(folder_btn)

        copy_md_btn = QPushButton("Copy Markdown")
        copy_md_btn.clicked.connect(self._on_copy_markdown)
        toolbar.addWidget(copy_md_btn)

        save_md_btn = QPushButton("Save .md...")
        save_md_btn.clicked.connect(self._on_save_markdown)
        toolbar.addWidget(save_md_btn)

        self.cli_combo = QComboBox()
        self.cli_combo.addItems(["codex", "claude"])
        toolbar.addWidget(self.cli_combo)

        launch_cli_btn = QPushButton("Launch LLM CLI")
        launch_cli_btn.setToolTip("Prepare thread.md + convert cached .docx drafts to Markdown, then launch Codex/Claude CLI.")
        launch_cli_btn.clicked.connect(self._on_launch_cli)
        toolbar.addWidget(launch_cli_btn)

        layout.addWidget(toolbar_widget)

        # In-thread search & company filter bar
        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        self.msg_search_edit = QLineEdit()
        self.msg_search_edit.setClearButtonEnabled(True)
        self.msg_search_edit.setPlaceholderText("🔍 Filter messages by keyword, TDoc, or sender...")
        self.msg_search_edit.textChanged.connect(self._render_messages)
        filter_row.addWidget(self.msg_search_edit, 1)

        self.company_filter_combo = QComboBox()
        self.company_filter_combo.addItem("All Companies", "")
        self.company_filter_combo.currentIndexChanged.connect(self._render_messages)
        filter_row.addWidget(self.company_filter_combo)

        self.msg_count_lbl = QLabel("")
        self.msg_count_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
        filter_row.addWidget(self.msg_count_lbl)
        layout.addLayout(filter_row)

        # Timeline browser
        self.browser = QTextBrowser()
        self.browser.setOpenLinks(False)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        layout.addWidget(self.browser, 1)

    def _on_order_changed(self):
        self._settings.setValue("newest_first", bool(self.order_combo.currentData()))
        self._render_messages()

    def closeEvent(self, event):
        try:
            self._settings.setValue("geometry", self.saveGeometry())
        except Exception:
            pass
        super().closeEvent(event)

    def _populate_company_filter(self, messages: list):
        cur_val = self.company_filter_combo.currentData() or ""
        companies = sorted({
            str(m.get("company") or "").strip()
            for m in messages
            if str(m.get("company") or "").strip() and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
        }, key=str.lower)
        self.company_filter_combo.blockSignals(True)
        self.company_filter_combo.clear()
        self.company_filter_combo.addItem("All Companies", "")
        for c in companies:
            self.company_filter_combo.addItem(c, c)
        idx = self.company_filter_combo.findData(cur_val)
        self.company_filter_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.company_filter_combo.blockSignals(False)

    def _load_data(self):
        thread = self.db.get_thread(self.thread_id)
        if not thread:
            return
        self._loading = True
        try:
            self.title_label.setText(f"📧 {thread.get('title', 'Email Thread')}")
            self.meta_label.setText(
                f"Meeting: {thread.get('meeting_id') or 'N/A'} | "
                f"Messages: {thread.get('message_count', 0)} | "
                f"Unread: {thread.get('unread_count', 0)} | "
                f"Last synced: {thread.get('last_synced_at') or 'Never'}"
            )
            self.ais_pills.setText(", ".join(thread.get("linked_ais") or []))
            self.wis_pills.setText(", ".join(thread.get("linked_wis") or []))
            self.tdocs_pills.setText(", ".join(thread.get("linked_tdocs") or []))
            self._populate_company_filter(self.db.get_thread_messages(self.thread_id))
        finally:
            self._loading = False
        self._render_messages()

    def _render_messages(self):
        thread = self.db.get_thread(self.thread_id)
        all_messages = self.db.get_thread_messages(self.thread_id)
        if not thread:
            return

        query = (self.msg_search_edit.text() if hasattr(self, "msg_search_edit") else "").strip().lower()
        comp_filter = (self.company_filter_combo.currentData() if hasattr(self, "company_filter_combo") else "") or ""

        filtered = []
        for m in all_messages:
            if comp_filter and (m.get("company") or "").strip().lower() != comp_filter.lower():
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
        md = export_thread_to_markdown(
            thread,
            filtered,
            deduplicate_quotes=self.dedup_check.isChecked(),
            interactive_links=True,
            newest_first=newest_first,
        )
        self.browser.setMarkdown(md)

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
        if scheme == "outlook":
            host = url.host().lower()
            path_parts = [urllib.parse.unquote(p) for p in url.path().strip("/").split("/") if p]
            if host == "msg" and path_parts:
                if not open_outlook_message(path_parts[0]):
                    QMessageBox.warning(self, "Outlook Message", "Could not open the selected message in Microsoft Outlook.")
            elif host == "reply" and path_parts:
                if not reply_to_outlook_message(path_parts[0], reply_all=True):
                    QMessageBox.warning(self, "Outlook Reply", "Could not open a Reply window in Microsoft Outlook.")
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

    def _on_links_changed(self):
        if self._loading:
            return
        ais = self.ais_pills.tags()
        wis = self.wis_pills.tags()
        tdocs = self.tdocs_pills.tags()
        if self._tdoc_family_expander:
            expanded = self._tdoc_family_expander(tdocs)
            if [t.upper() for t in expanded] != [t.upper() for t in tdocs]:
                self._loading = True
                try:
                    self.tdocs_pills.setText(", ".join(expanded))
                finally:
                    self._loading = False
                tdocs = expanded
        self.db.update_thread_links(self.thread_id, linked_ais=ais, linked_wis=wis, linked_tdocs=tdocs)
        self.thread_changed.emit()
        self._render_messages()

    def _on_refresh(self):
        new_count = refresh_thread(self.db, self.thread_id)
        self._load_data()
        self.thread_changed.emit()
        QMessageBox.information(self, "Thread Refreshed", f"Sync complete. {new_count} new message(s) found.")

    def _on_mark_read(self):
        self.db.mark_thread_read(self.thread_id)
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
            QMessageBox.information(self, "No Other Threads", "No other tracked threads in this meeting to merge.")
            return
        items = [f"{t['thread_id']} — {t['title']} ({t.get('message_count', 0)} msgs)" for t in candidates]
        choice, ok = QInputDialog.getItem(
            self,
            "Merge Split Thread",
            "Select another tracked thread to merge into this thread:",
            items,
            0,
            False,
        )
        if ok and choice:
            src_id = choice.split(" — ", 1)[0].strip()
            self.db.merge_threads(self.thread_id, [src_id])
            self._load_data()
            self.thread_changed.emit()

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

    def _on_launch_cli(self):
        cli = self.cli_combo.currentText()
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
        cmd = f'start powershell -NoExit -Command "cd \\"{ws_dir}\\"; {cli} \\"{prompt}\\""'
        subprocess.Popen(cmd, shell=True)


class PinFromOutlookDialog(QDialog):
    """Allows browsing and pinning email threads from configured WG Outlook folders."""

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        meeting_id: str,
        wg: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.meeting_id = meeting_id
        self.wg = wg
        self.pinned_thread_ids = []
        self._threads = []
        self.setWindowTitle(f"Pin Email Thread from Configured {wg} Outlook Folders")
        self.resize(780, 520)
        self._build_ui()
        self._load_threads()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Filter threads:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Type KI, WI, TDoc, or subject keywords (e.g. KI#1_1, NAS Routing)...")
        self.search_edit.textChanged.connect(self._apply_filter)
        search_row.addWidget(self.search_edit, 1)

        reload_btn = QPushButton("Rescan Folders")
        reload_btn.clicked.connect(self._load_threads)
        search_row.addWidget(reload_btn)
        layout.addLayout(search_row)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        layout.addWidget(self.list_widget, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        pin_btn = QPushButton("📌 Pin Selected Thread(s)")
        pin_btn.clicked.connect(self._on_pin_selected)
        btn_row.addWidget(pin_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _load_threads(self):
        self._threads = list_recent_threads_in_folders(self.wg)
        self._apply_filter()

    def _apply_filter(self):
        q = self.search_edit.text().strip().lower()
        self.list_widget.clear()
        for t in self._threads:
            title = t.get("title", "")
            sender = t.get("last_sender", "")
            if q and q not in title.lower() and q not in sender.lower():
                continue
            label = f"📧 {title}  ({t.get('count', 1)} msgs • Last: {t.get('last_date', '')} by {sender})"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, t)
            self.list_widget.addItem(item)

    def _on_pin_selected(self):
        selected = self.list_widget.selectedItems()
        if not selected:
            QMessageBox.information(self, "No Thread Selected", "Please select at least one email thread to pin.")
            return
        for item in selected:
            t = item.data(Qt.UserRole)
            tid = pin_thread_by_subject_or_cid(
                self.db,
                title=t["title"],
                normalized_subject=t["normalized_subject"],
                conversation_id=t.get("conversation_id", ""),
                meeting_id=self.meeting_id,
                wg=self.wg,
            )
            self.pinned_thread_ids.append(tid)
        self.accept()

