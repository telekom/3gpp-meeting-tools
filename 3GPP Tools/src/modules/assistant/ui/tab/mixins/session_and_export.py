"""Ollama status, new-session reset, clipboard copy, and Markdown export mixin."""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

from core.ui.toast import notify_info, notify_success
from modules.assistant.ui.assistant_components import (
    STATUS_PILL_NEUTRAL,
    STATUS_PILL_OFFLINE,
    STATUS_PILL_ONLINE,
)

logger = logging.getLogger(__name__)


class AssistantSessionExportMixin:
    """Handles Ollama status polling/settings and session lifecycle/export actions."""

    def _refresh_ollama_status(self) -> None:
        try:
            from core.ai.ollama_client import load_ollama_config, is_ollama_port_open

            cfg = load_ollama_config()
            host = cfg.get("host", "http://127.0.0.1:11434")
            model = cfg.get("selected_model", "")
            is_open = is_ollama_port_open(host)
            if is_open and model:
                self.model_status_btn.setText(f"🟢 Ollama: {model}")
                self.model_status_btn.setStyleSheet(STATUS_PILL_ONLINE)
                self.model_status_btn.setToolTip(
                    f"Host: {host}\nModel: {model}\nClick to configure Ollama settings."
                )
            elif is_open:
                self.model_status_btn.setText("🟡 Ollama: No Model Selected")
                self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
                self.model_status_btn.setToolTip(f"Host: {host}\nNo model selected.\nClick to configure.")
            else:
                self.model_status_btn.setText("🔴 Ollama: Offline")
                self.model_status_btn.setStyleSheet(STATUS_PILL_OFFLINE)
                self.model_status_btn.setToolTip(
                    f"Host: {host}\nDaemon not responding.\nClick to configure or start service."
                )
        except Exception as exc:
            logger.debug("Failed to query Ollama status: %s", exc)
            self.model_status_btn.setText("⚪ Ollama: Unknown")
            self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        if hasattr(self, "dock_model_status_btn"):
            self.dock_model_status_btn.setText(self.model_status_btn.text())
            self.dock_model_status_btn.setStyleSheet(self.model_status_btn.styleSheet())
            self.dock_model_status_btn.setToolTip(self.model_status_btn.toolTip())

    def _on_open_ollama_settings(self) -> None:
        if self.receivers(self.open_ollama_settings_requested) > 0:
            self.open_ollama_settings_requested.emit()
        else:
            try:
                from core.ai.ollama_dialog import OllamaConfigDialog

                dlg = OllamaConfigDialog(self)
                dlg.exec_()
            except Exception as exc:
                logger.error("Could not open Ollama settings dialog: %s", exc)
        QTimer.singleShot(1500, self._refresh_ollama_status)

    def _on_new_chat(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Active Research in Progress",
                "A research query is currently running. Do you want to stop it and begin a new session?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self._on_stop()
            else:
                return

        self._conversation.clear()
        self._activities.clear()
        self._evidence.clear()
        self._last_result = None
        self._exchange_events.clear()
        self.conversation_browser.clear()
        self.activity_widget.clear()
        self.sources_tree.clear()
        self.evidence_inspector.set_evidence(None)
        self.diagnostics_widget.clear()
        self._update_activity_count()
        self._update_source_counts()
        self.cockpit_tabs.setTabText(1, "⚡ Activity (0)")
        self._set_status_text("Ready")
        self.starter_chips.setVisible(True)
        self.question_input.clear()
        self.question_input.setFocus()

    def _build_transcript_markdown(self, include_timestamp: bool = False) -> str:
        """Render the current conversation and referenced sources as Markdown."""
        lines = ["# 3GPP Research Assistant Transcript\n"]
        if include_timestamp:
            lines.append(f"*Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n")
        for msg in self._conversation:
            role = "### User" if msg.get("role") == "user" else "### Assistant"
            content = msg.get("content", "")
            lines.append(f"{role}\n\n{content}\n")
        if self._evidence:
            lines.append("## Referenced Sources & Evidence\n")
            for eid, ev in sorted(self._evidence.items()):
                label = self._source_label(ev)
                lines.append(f"- **[{eid}]** {label}")
        return "\n".join(lines)

    def _on_copy_conversation(self) -> None:
        if not self._conversation:
            notify_info(self, "Copy Chat", "Conversation transcript is currently empty.")
            return
        full_text = self._build_transcript_markdown(include_timestamp=False)
        QApplication.clipboard().setText(full_text)
        self.status_label.setText("Transcript copied!")
        notify_success(self, "Transcript Copied", "Conversation transcript copied to clipboard.")
        QTimer.singleShot(
            2500,
            lambda: self.status_label.setText(
                "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
            ),
        )

    def _on_export_markdown(self) -> None:
        if not self._conversation:
            notify_info(self, "Export Transcript", "Conversation transcript is currently empty.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Research Transcript",
            f"3gpp_assistant_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
            "Markdown Files (*.md);;All Files (*)",
        )
        if not filename:
            return
        content = self._build_transcript_markdown(include_timestamp=True)
        try:
            Path(filename).write_text(content, encoding="utf-8")
            self.status_label.setText("Exported successfully!")
            notify_success(
                self,
                "Transcript Exported",
                f"Saved Markdown transcript to {Path(filename).name}",
            )
            QTimer.singleShot(
                2500,
                lambda: self.status_label.setText(
                    "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
                ),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Failed to save transcript: {exc}")
