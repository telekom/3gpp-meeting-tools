"""Research-oriented Qt UI for the local 3GPP Assistant."""

from __future__ import annotations

import html
import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QFrame, QGroupBox, QHBoxLayout, QLabel, QMessageBox,
    QPushButton, QPlainTextEdit, QSplitter, QTextBrowser, QToolButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    TREE_STYLE_CLEAN,
)
from modules.assistant.core.agent_controller import AgentControllerConfig
from modules.assistant.core.agent_models import (
    AgentResult, Evidence, EvidenceSourceType, ResearchActivity,
)
from modules.assistant.core.ollama_agent_adapter import OllamaAgentAdapter
from modules.assistant.core.tools import build_read_only_tool_registry
from modules.assistant.ui.assistant_components import (
    CARD_FRAME_STYLE,
    CARD_HEADER_STYLE,
    MultilinePromptEdit,
    StarterChipsWidget,
    STATUS_PILL_OFFLINE,
    STATUS_PILL_ONLINE,
    STATUS_PILL_NEUTRAL,
)
from modules.assistant.ui.assistant_worker import AgentResearchWorker

logger = logging.getLogger(__name__)


class AssistantTab(QWidget):
    """Conversation + objective research activity + application-owned sources."""

    open_ollama_settings_requested = pyqtSignal()
    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)

    def __init__(
        self,
        *,
        specs_db_path: Path,
        spec_search_db_path: Path,
        protocol_db_path: Path,
        parent=None,
    ):
        super().__init__(parent)
        self.specs_db_path = Path(specs_db_path)
        self.spec_search_db_path = Path(spec_search_db_path)
        self.protocol_db_path = Path(protocol_db_path)

        self.tool_registry = build_read_only_tool_registry(
            specs_db_path=self.specs_db_path,
            spec_search_db_path=self.spec_search_db_path,
            protocol_db_path=self.protocol_db_path,
        )
        self._worker: Optional[AgentResearchWorker] = None
        self._conversation: List[Dict[str, str]] = []
        self._activities: List[ResearchActivity] = []
        self._evidence: Dict[str, Evidence] = {}
        self._last_result: Optional[AgentResult] = None
        self._exchange_events: List[dict] = []
        self._setup_ui()
        # Explicit idle state; independent of Ollama monitor state.
        self._set_running(False)
        self.status_label.setText("Ready")
        self._refresh_ollama_status()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # Top Header Ribbon (Model Status, Session Actions, Layout Switcher)
        # ---------------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("assistantCardHeader")
        header_card.setStyleSheet("""
            QFrame#assistantCardHeader {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px;
            }
        """)
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(8, 4, 8, 4)
        header_layout.setSpacing(8)

        title = QLabel("🤖 3GPP Assistant")
        title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        header_layout.addWidget(title)

        # Model status pill
        self.model_status_btn = QPushButton("⚪ Ollama: Checking...")
        self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        self.model_status_btn.setCursor(Qt.PointingHandCursor)
        self.model_status_btn.setToolTip("Click to configure Ollama settings.")
        self.model_status_btn.clicked.connect(self._on_open_ollama_settings)
        header_layout.addWidget(self.model_status_btn)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #64748B; font-weight: 500; font-size: 12px;")
        header_layout.addWidget(self.status_label)

        header_layout.addStretch()

        # Session Action Buttons
        self.new_chat_btn = QPushButton("➕ New Chat")
        self.new_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.new_chat_btn.setToolTip("Clear the conversation and begin a new research session.")
        self.new_chat_btn.clicked.connect(self._on_new_chat)
        header_layout.addWidget(self.new_chat_btn)

        self.copy_chat_btn = QPushButton("📋 Copy Chat")
        self.copy_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_chat_btn.setToolTip("Copy the full conversation transcript and citations to the clipboard.")
        self.copy_chat_btn.clicked.connect(self._on_copy_conversation)
        header_layout.addWidget(self.copy_chat_btn)

        self.export_chat_btn = QPushButton("💾 Export...")
        self.export_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.export_chat_btn.setToolTip("Export the research conversation to a Markdown (.md) file.")
        self.export_chat_btn.clicked.connect(self._on_export_markdown)
        header_layout.addWidget(self.export_chat_btn)

        # Layout Switcher Button
        self.layout_toggle_btn = QPushButton("⬒ Bottom Split")
        self.layout_toggle_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.layout_toggle_btn.setCheckable(True)
        self.layout_toggle_btn.setChecked(True)  # True = Side-by-Side (Horizontal)
        self.layout_toggle_btn.setToolTip("Toggle between Side-by-Side and Bottom Split layouts.")
        self.layout_toggle_btn.clicked.connect(self._toggle_layout)
        header_layout.addWidget(self.layout_toggle_btn)

        layout.addWidget(header_card)

        # ---------------------------------------------------------------------
        # Main Splitter (Left: Chat & Input, Right: Research Cockpit Drawer)
        # ---------------------------------------------------------------------
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)

        # Left / Chat Widget
        chat_widget = QWidget()
        chat_layout = QVBoxLayout(chat_widget)
        chat_layout.setContentsMargins(0, 0, 0, 0)
        chat_layout.setSpacing(6)

        # Conversation Card Frame
        conv_frame = QFrame()
        conv_frame.setObjectName("assistantCardFrame")
        conv_frame.setStyleSheet(CARD_FRAME_STYLE)
        conv_frame_layout = QVBoxLayout(conv_frame)
        conv_frame_layout.setContentsMargins(8, 8, 8, 8)
        conv_frame_layout.setSpacing(6)

        # Starter Chips for empty state
        self.starter_chips = StarterChipsWidget()
        self.starter_chips.prompt_selected.connect(self._on_starter_chip_selected)
        conv_frame_layout.addWidget(self.starter_chips)

        self.conversation_browser = QTextBrowser()
        self.conversation_browser.setOpenExternalLinks(False)
        self.conversation_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #FFFFFF;
                border: none;
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 13px;
                color: #1E293B;
            }
        """)
        self.conversation_browser.setPlaceholderText(
            "Ask a 3GPP question. The Assistant will research local specification "
            "and protocol knowledge before answering."
        )
        conv_frame_layout.addWidget(self.conversation_browser, stretch=1)
        chat_layout.addWidget(conv_frame, stretch=1)

        # Input Card Frame
        input_frame = QFrame()
        input_frame.setObjectName("assistantCardFrame")
        input_frame.setStyleSheet(CARD_FRAME_STYLE)
        input_frame_layout = QVBoxLayout(input_frame)
        input_frame_layout.setContentsMargins(8, 8, 8, 8)
        input_frame_layout.setSpacing(6)

        self.question_input = MultilinePromptEdit()
        self.question_input.submit_requested.connect(self._on_send)
        self.question_input.textChanged.connect(self._on_prompt_changed)
        input_frame_layout.addWidget(self.question_input)

        # Action / Control bar under input
        input_action_bar = QHBoxLayout()
        input_action_bar.setContentsMargins(0, 0, 0, 0)
        input_action_bar.setSpacing(6)

        self.char_count_label = QLabel("0 chars")
        self.char_count_label.setStyleSheet("color: #94A3B8; font-size: 11px;")
        input_action_bar.addWidget(self.char_count_label)

        self.clear_input_btn = QPushButton("✕ Clear")
        self.clear_input_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_input_btn.setEnabled(False)
        self.clear_input_btn.clicked.connect(self.question_input.clear)
        input_action_bar.addWidget(self.clear_input_btn)

        input_action_bar.addStretch()

        self.send_btn = QPushButton("Send ↵")
        self.send_btn.setObjectName("primaryBtn")
        self.send_btn.clicked.connect(self._on_send)
        input_action_bar.addWidget(self.send_btn)

        self.stop_btn = QPushButton("■ Stop")
        self.stop_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop)
        input_action_bar.addWidget(self.stop_btn)

        input_frame_layout.addLayout(input_action_bar)
        chat_layout.addWidget(input_frame)

        self.main_splitter.addWidget(chat_widget)

        # ---------------------------------------------------------------------
        # Right / Research Cockpit Drawer
        # ---------------------------------------------------------------------
        self.detail_widget = QWidget()
        detail_layout = QVBoxLayout(self.detail_widget)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(6)

        # Research Activity Card
        activity_group = QGroupBox()
        activity_layout = QVBoxLayout(activity_group)
        activity_header = QHBoxLayout()
        self.activity_toggle = QToolButton()
        self.activity_toggle.setText("▼ Research activity")
        self.activity_toggle.setCheckable(True)
        self.activity_toggle.setChecked(True)
        self.activity_toggle.clicked.connect(self._toggle_activity)
        activity_header.addWidget(self.activity_toggle)
        activity_header.addStretch()
        self.activity_count = QLabel("0 steps")
        self.activity_count.setStyleSheet("color: #64748B;")
        activity_header.addWidget(self.activity_count)
        activity_layout.addLayout(activity_header)
        self.activity_browser = QTextBrowser()
        activity_layout.addWidget(self.activity_browser)
        detail_layout.addWidget(activity_group, stretch=1)

        # Sources Tree Card
        sources_group = QGroupBox("Sources")
        sources_layout = QVBoxLayout(sources_group)
        self.sources_tree = QTreeWidget()
        self.sources_tree.setStyleSheet(TREE_STYLE_CLEAN)
        self.sources_tree.setHeaderLabels(["ID", "Source", "Version / Clause"])
        self.sources_tree.setAlternatingRowColors(True)
        self.sources_tree.setRootIsDecorated(False)
        self.sources_tree.itemDoubleClicked.connect(self._show_source_detail)
        sources_layout.addWidget(self.sources_tree)
        detail_layout.addWidget(sources_group, stretch=1)

        # Exchange Details Card
        exchange_group = QGroupBox()
        exchange_layout = QVBoxLayout(exchange_group)
        exchange_header = QHBoxLayout()
        self.exchange_toggle = QToolButton()
        self.exchange_toggle.setText("▶ Ollama exchange details")
        self.exchange_toggle.setCheckable(True)
        self.exchange_toggle.setChecked(False)
        self.exchange_toggle.clicked.connect(self._toggle_exchange)
        exchange_header.addWidget(self.exchange_toggle)
        exchange_header.addStretch()
        self.copy_exchange_btn = QPushButton("Copy Trace")
        self.copy_exchange_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_exchange_btn.setVisible(False)
        self.copy_exchange_btn.clicked.connect(self._copy_exchange_trace)
        exchange_header.addWidget(self.copy_exchange_btn)
        self.clear_exchange_btn = QPushButton("Clear")
        self.clear_exchange_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_exchange_btn.setVisible(False)
        self.clear_exchange_btn.clicked.connect(self._clear_exchange_trace)
        exchange_header.addWidget(self.clear_exchange_btn)
        exchange_layout.addLayout(exchange_header)
        self.exchange_browser = QPlainTextEdit()
        self.exchange_browser.setReadOnly(True)
        self.exchange_browser.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.exchange_browser.setPlaceholderText(
            "Raw requests sent to Ollama and responses returned by Ollama appear here."
        )
        self.exchange_browser.setVisible(False)
        exchange_layout.addWidget(self.exchange_browser)
        detail_layout.addWidget(exchange_group)

        self.main_splitter.addWidget(self.detail_widget)

        # Default side-by-side sizes (65% Chat, 35% Drawer)
        self.main_splitter.setSizes([680, 420])
        layout.addWidget(self.main_splitter, stretch=1)

    def _toggle_layout(self, checked: bool) -> None:
        """Toggle between Side-by-Side (Horizontal) and Bottom Split (Vertical)."""
        if checked:
            self.main_splitter.setOrientation(Qt.Horizontal)
            self.main_splitter.setSizes([680, 420])
            self.layout_toggle_btn.setText("⬒ Bottom Split")
            self.layout_toggle_btn.setToolTip("Switch to stacked top/bottom layout")
        else:
            self.main_splitter.setOrientation(Qt.Vertical)
            self.main_splitter.setSizes([450, 280])
            self.layout_toggle_btn.setText("◫ Side Drawer")
            self.layout_toggle_btn.setToolTip("Switch to side-by-side drawer layout")

    def _on_prompt_changed(self) -> None:
        text = self.question_input.toPlainText()
        count = len(text)
        self.char_count_label.setText(f"{count} char" + ("s" if count != 1 else ""))
        self.clear_input_btn.setEnabled(count > 0)

    def _on_starter_chip_selected(self, prompt: str) -> None:
        self.question_input.setPlainText(prompt)
        self._on_send()

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
                self.model_status_btn.setToolTip(f"Host: {host}\nModel: {model}\nClick to configure Ollama settings.")
            elif is_open:
                self.model_status_btn.setText("🟡 Ollama: No Model Selected")
                self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
                self.model_status_btn.setToolTip(f"Host: {host}\nNo model selected.\nClick to configure.")
            else:
                self.model_status_btn.setText("🔴 Ollama: Offline")
                self.model_status_btn.setStyleSheet(STATUS_PILL_OFFLINE)
                self.model_status_btn.setToolTip(f"Host: {host}\nDaemon not responding.\nClick to configure or start service.")
        except Exception as exc:
            logger.debug("Failed to query Ollama status: %s", exc)
            self.model_status_btn.setText("⚪ Ollama: Unknown")
            self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)

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
        self.activity_browser.clear()
        self.sources_tree.clear()
        if hasattr(self, "exchange_browser"):
            self.exchange_browser.clear()
        self._update_activity_count()
        self.status_label.setText("Ready")
        self.starter_chips.setVisible(True)
        self.question_input.clear()
        self.question_input.setFocus()

    def _on_copy_conversation(self) -> None:
        if not self._conversation:
            QMessageBox.information(self, "Copy Chat", "Conversation transcript is currently empty.")
            return
        lines = ["# 3GPP Research Assistant Transcript\n"]
        for msg in self._conversation:
            role = "### User" if msg.get("role") == "user" else "### Assistant"
            content = msg.get("content", "")
            lines.append(f"{role}\n\n{content}\n")
        if self._evidence:
            lines.append("## Referenced Sources & Evidence\n")
            for eid, ev in sorted(self._evidence.items()):
                label = self._source_label(ev)
                lines.append(f"- **[{eid}]** {label}")
        full_text = "\n".join(lines)
        QApplication.clipboard().setText(full_text)
        self.status_label.setText("Transcript copied!")
        QTimer.singleShot(
            2500,
            lambda: self.status_label.setText(
                "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
            ),
        )

    def _on_export_markdown(self) -> None:
        if not self._conversation:
            QMessageBox.information(self, "Export Transcript", "Conversation transcript is currently empty.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Research Transcript",
            f"3gpp_assistant_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
            "Markdown Files (*.md);;All Files (*)",
        )
        if not filename:
            return
        lines = [
            "# 3GPP Research Assistant Transcript\n",
            f"*Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n",
        ]
        for msg in self._conversation:
            role = "### User" if msg.get("role") == "user" else "### Assistant"
            content = msg.get("content", "")
            lines.append(f"{role}\n\n{content}\n")
        if self._evidence:
            lines.append("## Referenced Sources & Evidence\n")
            for eid, ev in sorted(self._evidence.items()):
                label = self._source_label(ev)
                lines.append(f"- **[{eid}]** {label}")
        try:
            Path(filename).write_text("\n".join(lines), encoding="utf-8")
            self.status_label.setText("Exported successfully!")
            QTimer.singleShot(
                2500,
                lambda: self.status_label.setText(
                    "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
                ),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Failed to save transcript: {exc}")

    def _on_send(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        question = self.question_input.toPlainText().strip()
        if not question:
            return

        self.question_input.clear()
        self.starter_chips.setVisible(False)
        self._append_conversation("user", question)
        self._activities.clear()
        self._evidence.clear()
        self._last_result = None
        self.activity_browser.clear()
        self.sources_tree.clear()
        self._clear_exchange_trace()
        self._update_activity_count()

        try:
            adapter = OllamaAgentAdapter()
        except Exception as exc:
            self._append_system_message(f"Could not initialize the local model: {exc}")
            return

        worker = AgentResearchWorker(
            question=question,
            tool_registry=self.tool_registry,
            adapter=adapter,
            conversation_context=self._conversation[:-1],
            controller_config=AgentControllerConfig(),
            parent=self,
        )
        self._worker = worker
        worker.research_started.connect(self._on_research_started)
        worker.activity_updated.connect(self._on_activity)
        worker.evidence_added.connect(self._on_evidence)
        worker.exchange_observed.connect(self._on_exchange)
        worker.result_ready.connect(self._on_result)
        worker.research_cancelled.connect(self._on_cancelled)
        worker.research_failed.connect(self._on_failed)
        worker.finished.connect(self._on_worker_finished)
        self._set_running(True)
        worker.start()

    def _on_stop(self) -> None:
        worker = self._worker
        if worker is None or not worker.isRunning():
            return
        self.status_label.setText("Cancelling…")
        self.stop_btn.setEnabled(False)
        worker.request_cancel()

    def _on_research_started(self) -> None:
        self.status_label.setText("Researching…")

    def _on_activity(self, activity: ResearchActivity) -> None:
        self._activities.append(activity)
        self._update_activity_count()
        self.activity_browser.append(
            f"<div style='margin:2px 0;'>• {html.escape(activity.message)}</div>"
        )

    def _on_evidence(self, evidence: Evidence) -> None:
        self._evidence[evidence.id] = evidence
        self._add_source_row(evidence)

    def _on_exchange(self, event: object) -> None:
        if not isinstance(event, dict):
            return
        event = dict(event)
        self._exchange_events.append(event)
        stamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        direction = str(event.get("direction") or "event")
        labels = {
            "request": "APP -> OLLAMA",
            "response_headers": "OLLAMA -> APP [HEADERS]",
            "response": "OLLAMA -> APP",
            "response_chunk": "OLLAMA -> APP [STREAM CHUNK]",
        }
        label = labels.get(direction, direction.upper())
        if direction == "request":
            body = {k: event.get(k) for k in ("method", "endpoint", "payload") if k in event}
        elif direction == "response_headers":
            body = {k: event.get(k) for k in ("status_code", "reason", "headers") if k in event}
        elif direction == "response_chunk":
            raw = event.get("raw", "")
            try:
                body = json.loads(raw)
            except Exception:
                body = raw
        else:
            body = event.get("payload", event)
        if isinstance(body, (dict, list)):
            rendered = json.dumps(body, indent=2, ensure_ascii=False, default=str)
        else:
            rendered = str(body)
        self.exchange_browser.appendPlainText(f"[{stamp}] {label}\n{rendered}\n")

    def _toggle_exchange(self, checked: bool) -> None:
        self.exchange_browser.setVisible(checked)
        self.copy_exchange_btn.setVisible(checked)
        self.clear_exchange_btn.setVisible(checked)
        self.exchange_toggle.setText(
            ("▼" if checked else "▶") + " Ollama exchange details"
        )

    def _copy_exchange_trace(self) -> None:
        text = self.exchange_browser.toPlainText()
        if text:
            QApplication.clipboard().setText(text)

    def _clear_exchange_trace(self) -> None:
        self._exchange_events.clear()
        if hasattr(self, "exchange_browser"):
            self.exchange_browser.clear()

    def _on_result(self, result: AgentResult) -> None:
        self._last_result = result
        # Final result is authoritative in case an evidence signal was missed.
        for evidence in result.evidence:
            if evidence.id not in self._evidence:
                self._evidence[evidence.id] = evidence
                self._add_source_row(evidence)

        rendered = self._render_answer(result.answer)
        self.conversation_browser.append(
            "<div style='margin:10px 0 4px 0;'><b>Assistant</b></div>" + rendered
        )
        self._conversation.append({"role": "assistant", "content": result.answer})
        if result.warnings:
            self.conversation_browser.append(
                "<div style='color:#B45309; margin-bottom:8px;'>"
                + "<br>".join(html.escape(w) for w in result.warnings)
                + "</div>"
            )
        self.status_label.setText("Ready")

    def _on_cancelled(self, result) -> None:
        self._append_system_message("Research stopped.")
        self.status_label.setText("Stopped")

    def _on_failed(self, message: str, result) -> None:
        logger.error("Assistant research failed: %s", message)
        self._append_system_message(message, error=True)
        self.status_label.setText("Failed")

    def _on_worker_finished(self) -> None:
        sender = self.sender()
        if sender is not None and sender is not self._worker:
            return
        worker = self._worker
        if worker is None:
            return
        self._set_running(False)
        worker.deleteLater()
        self._worker = None
        if self.status_label.text() == "Cancelling…":
            self.status_label.setText("Stopped")

    def shutdown(self, wait_ms: int = 500) -> None:
        """Request cooperative cancellation during application shutdown."""
        worker = self._worker
        if worker is None or not worker.isRunning():
            return
        worker.request_cancel()
        if not worker.wait(wait_ms):
            logger.warning(
                "Assistant worker is still waiting for an active operation during shutdown; "
                "it was not force-terminated."
            )

    def _set_running(self, running: bool) -> None:
        self.send_btn.setEnabled(not running)
        self.question_input.setEnabled(not running)
        self.stop_btn.setEnabled(running)
        if not running:
            self.question_input.setFocus()

    def _append_conversation(self, role: str, content: str) -> None:
        label = "You" if role == "user" else "Assistant"
        self.conversation_browser.append(
            f"<div style='margin:10px 0 4px 0;'><b>{label}</b></div>"
            f"<div style='white-space:pre-wrap;'>{html.escape(content)}</div>"
        )
        self._conversation.append({"role": role, "content": content})

    def _append_system_message(self, message: str, error: bool = False) -> None:
        color = "#B91C1C" if error else "#64748B"
        self.conversation_browser.append(
            f"<div style='color:{color}; margin:8px 0;'>{html.escape(message)}</div>"
        )

    def _render_answer(self, answer: str) -> str:
        escaped = html.escape(answer).replace("\n", "<br>")
        # Only IDs present in the application evidence store are made visually
        # citation-like. Unknown IDs remain ordinary escaped text.
        valid_ids = set(self._evidence)

        def _highlight_bracket(match: re.Match) -> str:
            inner = match.group(1)
            if not any(eid in valid_ids for eid in re.findall(r"\bE\d+\b", inner)):
                return match.group(0)
            highlighted = re.sub(
                r"\b(E\d+)\b",
                lambda m: (
                    f"<span style='color:#0369A1; font-weight:bold;'>{m.group(1)}</span>"
                    if m.group(1) in valid_ids
                    else m.group(1)
                ),
                inner,
            )
            return f"<span style='color:#0369A1; font-weight:bold;'>[</span>{highlighted}<span style='color:#0369A1; font-weight:bold;'>]</span>"

        escaped = re.sub(r"\[([^\[\]]*?\bE\d+\b[^\[\]]*?)\]", _highlight_bracket, escaped)
        return f"<div style='white-space:pre-wrap;'>{escaped}</div>"

    def _add_source_row(self, evidence: Evidence) -> None:
        source = self._source_label(evidence)
        detail_parts = []
        if evidence.version:
            detail_parts.append(f"v{evidence.version}")
        if evidence.clause:
            detail_parts.append(f"Clause {evidence.clause}")
        item = QTreeWidgetItem(
            [evidence.id, source, " — ".join(detail_parts)]
        )
        item.setData(0, Qt.UserRole, evidence.id)
        tooltip = evidence.title or source
        item.setToolTip(1, tooltip)
        self.sources_tree.addTopLevelItem(item)
        for column in range(3):
            self.sources_tree.resizeColumnToContents(column)

    @staticmethod
    def _source_label(evidence: Evidence) -> str:
        if evidence.source_type == EvidenceSourceType.SPECIFICATION_TEXT:
            prefix = f"TS {evidence.specification}" if evidence.specification else evidence.source_name
            return f"{prefix}: {evidence.title or 'indexed clause'}"
        if evidence.source_type == EvidenceSourceType.SPECIFICATION_CATALOGUE:
            prefix = f"TS {evidence.specification}" if evidence.specification else evidence.source_name
            return f"{prefix}: catalogue"
        if evidence.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE:
            return evidence.title or "Structured protocol knowledge"
        return evidence.source_name

    def _show_source_detail(self, item: QTreeWidgetItem, column: int) -> None:
        evidence_id = item.data(0, Qt.UserRole)
        evidence = self._evidence.get(str(evidence_id))
        if evidence is None:
            return
        content = evidence.content
        if isinstance(content, dict):
            import json
            text = json.dumps(content, indent=2, ensure_ascii=False, default=str)
        else:
            text = str(content or "")
        header = self._source_label(evidence)
        QMessageBox.information(
            self,
            f"{evidence.id} — {header}",
            (text[:12000] + ("\n\n[Content shortened for display]" if len(text) > 12000 else "")),
        )

    def _toggle_activity(self, checked: bool) -> None:
        self.activity_browser.setVisible(checked)
        self.activity_toggle.setText(
            ("▼" if checked else "▶") + " Research activity"
        )

    def _update_activity_count(self) -> None:
        count = len(self._activities)
        self.activity_count.setText(f"{count} step" + ("" if count == 1 else "s"))
