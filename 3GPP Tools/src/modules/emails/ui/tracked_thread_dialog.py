# --- File: src/modules/emails/ui/tracked_thread_dialog.py ---
import subprocess
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import refresh_thread


class TrackedThreadDialog(QDialog):
    """Integrated viewer for a tracked email thread with Markdown export and LLM CLI launch."""

    def __init__(self, db: TrackedThreadsDatabase, thread_id: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db = db
        self.thread_id = thread_id
        self.setWindowTitle("Tracked Email Thread")
        self.resize(920, 680)
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

        # Links editor row
        links_row = QHBoxLayout()
        links_row.addWidget(QLabel("AIs:"))
        self.ais_edit = QLineEdit()
        self.ais_edit.setPlaceholderText("e.g. KI#1_1, AI 6.2")
        links_row.addWidget(self.ais_edit)

        links_row.addWidget(QLabel("WIs:"))
        self.wis_edit = QLineEdit()
        self.wis_edit.setPlaceholderText("e.g. FS_6G_ARC")
        links_row.addWidget(self.wis_edit)

        links_row.addWidget(QLabel("TDocs:"))
        self.tdocs_edit = QLineEdit()
        self.tdocs_edit.setPlaceholderText("e.g. S2-2601234")
        links_row.addWidget(self.tdocs_edit)

        save_links_btn = QPushButton("Save Links")
        save_links_btn.clicked.connect(self._save_links)
        links_row.addWidget(save_links_btn)
        layout.addLayout(links_row)

        # Toolbar row
        toolbar = QHBoxLayout()
        self.dedup_check = QCheckBox("Deduplicate trailing quotes")
        self.dedup_check.setChecked(True)
        self.dedup_check.toggled.connect(self._render_messages)
        toolbar.addWidget(self.dedup_check)

        toolbar.addStretch()

        refresh_btn = QPushButton("Refresh Now")
        refresh_btn.clicked.connect(self._on_refresh)
        toolbar.addWidget(refresh_btn)

        read_btn = QPushButton("Mark All Read")
        read_btn.clicked.connect(self._on_mark_read)
        toolbar.addWidget(read_btn)

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
        launch_cli_btn.clicked.connect(self._on_launch_cli)
        toolbar.addWidget(launch_cli_btn)

        layout.addLayout(toolbar)

        # Timeline browser
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)
        layout.addWidget(self.browser, 1)

    def _load_data(self):
        thread = self.db.get_thread(self.thread_id)
        if not thread:
            return
        self.title_label.setText(f"📧 {thread.get('title', 'Email Thread')}")
        self.meta_label.setText(
            f"Meeting: {thread.get('meeting_id') or 'N/A'} | "
            f"Messages: {thread.get('message_count', 0)} | "
            f"Unread: {thread.get('unread_count', 0)} | "
            f"Last synced: {thread.get('last_synced_at') or 'Never'}"
        )
        self.ais_edit.setText(", ".join(thread.get("linked_ais") or []))
        self.wis_edit.setText(", ".join(thread.get("linked_wis") or []))
        self.tdocs_edit.setText(", ".join(thread.get("linked_tdocs") or []))
        self._render_messages()

    def _render_messages(self):
        thread = self.db.get_thread(self.thread_id)
        messages = self.db.get_thread_messages(self.thread_id)
        if not thread:
            return
        md = export_thread_to_markdown(thread, messages, deduplicate_quotes=self.dedup_check.isChecked())
        self.browser.setMarkdown(md)

    def _save_links(self):
        ais = [x.strip() for x in self.ais_edit.text().split(",") if x.strip()]
        wis = [x.strip() for x in self.wis_edit.text().split(",") if x.strip()]
        tdocs = [x.strip() for x in self.tdocs_edit.text().split(",") if x.strip()]
        self.db.update_thread_links(self.thread_id, linked_ais=ais, linked_wis=wis, linked_tdocs=tdocs)
        QMessageBox.information(self, "Links Saved", "Updated AI/WI/TDoc links for this thread.")

    def _on_refresh(self):
        new_count = refresh_thread(self.db, self.thread_id)
        self._load_data()
        QMessageBox.information(self, "Thread Refreshed", f"Sync complete. {new_count} new message(s) found.")

    def _on_mark_read(self):
        self.db.mark_thread_read(self.thread_id)
        self._load_data()

    def _get_markdown(self) -> str:
        thread = self.db.get_thread(self.thread_id) or {}
        messages = self.db.get_thread_messages(self.thread_id)
        return export_thread_to_markdown(thread, messages, deduplicate_quotes=self.dedup_check.isChecked())

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
        out_dir = Path.home() / ".3gpp_tools_tmp"
        out_dir.mkdir(parents=True, exist_ok=True)
        md_path = out_dir / f"{self.thread_id}.md"
        md_path.write_text(self._get_markdown(), encoding="utf-8")
        prompt = f"Analyze the 3GPP email thread in {md_path} and summarize the views exchanged by company, key disagreements, and open points."
        cmd = f'start powershell -NoExit -Command "{cli} \\"{prompt}\\""'
        subprocess.Popen(cmd, shell=True)
