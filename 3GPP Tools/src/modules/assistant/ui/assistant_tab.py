"""Research-oriented Qt UI for the local 3GPP Assistant."""

from __future__ import annotations

import html
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QPlainTextEdit, QSplitter, QTextBrowser, QToolButton, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget,
)

from modules.assistant.core.agent_controller import AgentControllerConfig
from modules.assistant.core.agent_models import (
    AgentResult, Evidence, EvidenceSourceType, ResearchActivity,
)
from modules.assistant.core.ollama_agent_adapter import OllamaAgentAdapter
from modules.assistant.core.tools import build_read_only_tool_registry
from modules.assistant.ui.assistant_worker import AgentResearchWorker

logger = logging.getLogger(__name__)


class AssistantTab(QWidget):
    """Conversation + objective research activity + application-owned sources."""

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

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(7)

        header = QHBoxLayout()
        title = QLabel("🤖 3GPP Assistant")
        title.setStyleSheet("font-weight: bold; font-size: 14px; color: #1E293B;")
        header.addWidget(title)
        header.addStretch()
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #64748B;")
        header.addWidget(self.status_label)
        layout.addLayout(header)

        splitter = QSplitter(Qt.Vertical)

        conversation_group = QGroupBox("Conversation")
        conversation_layout = QVBoxLayout(conversation_group)
        self.conversation_browser = QTextBrowser()
        self.conversation_browser.setOpenExternalLinks(False)
        self.conversation_browser.setPlaceholderText(
            "Ask a 3GPP question. The Assistant will research local specification "
            "and protocol knowledge before answering."
        )
        conversation_layout.addWidget(self.conversation_browser)
        splitter.addWidget(conversation_group)

        detail_splitter = QSplitter(Qt.Horizontal)

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
        detail_splitter.addWidget(activity_group)

        sources_group = QGroupBox("Sources")
        sources_layout = QVBoxLayout(sources_group)
        self.sources_tree = QTreeWidget()
        self.sources_tree.setHeaderLabels(["ID", "Source", "Version / Clause"])
        self.sources_tree.setAlternatingRowColors(True)
        self.sources_tree.setRootIsDecorated(False)
        self.sources_tree.itemDoubleClicked.connect(self._show_source_detail)
        sources_layout.addWidget(self.sources_tree)
        detail_splitter.addWidget(sources_group)

        detail_splitter.setSizes([430, 520])
        splitter.addWidget(detail_splitter)

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
        self.copy_exchange_btn.clicked.connect(self._copy_exchange_trace)
        exchange_header.addWidget(self.copy_exchange_btn)
        self.clear_exchange_btn = QPushButton("Clear")
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
        splitter.addWidget(exchange_group)

        splitter.setSizes([480, 240, 36])
        layout.addWidget(splitter)

        input_row = QHBoxLayout()
        self.question_input = QLineEdit()
        self.question_input.setPlaceholderText("Ask about 3GPP specifications, procedures, messages or IEs...")
        self.question_input.returnPressed.connect(self._on_send)
        input_row.addWidget(self.question_input, stretch=1)

        self.send_btn = QPushButton("Send")
        self.send_btn.clicked.connect(self._on_send)
        input_row.addWidget(self.send_btn)

        self.stop_btn = QPushButton("■ Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop)
        input_row.addWidget(self.stop_btn)
        layout.addLayout(input_row)

    def _on_send(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        question = self.question_input.text().strip()
        if not question:
            return

        self.question_input.clear()
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
        for evidence_id in sorted(self._evidence, key=len, reverse=True):
            escaped = escaped.replace(
                f"[{evidence_id}]",
                f"<span style='color:#0369A1; font-weight:bold;'>[{evidence_id}]</span>",
            )
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
