"""Worker execution, live activity/evidence ingestion, and shutdown mixin."""

from __future__ import annotations

import logging
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTreeWidgetItem

from modules.assistant.core.agent_controller import AgentControllerConfig
from modules.assistant.core.agent_models import (
    AgentResult,
    Evidence,
    EvidenceSourceType,
    ResearchActivity,
)
from modules.assistant.core.ollama_agent_adapter import OllamaAgentAdapter
from modules.assistant.ui.assistant_components import (
    render_assistant_message_card,
    render_markdown_to_html,
    render_system_message_card,
    render_user_message_card,
)
from modules.assistant.ui.assistant_worker import AgentResearchWorker

logger = logging.getLogger(__name__)


class AssistantResearchRunnerMixin:
    """Manages background research worker lifecycle, telemetry, and message rendering."""

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
        self.activity_widget.clear()
        self.sources_tree.clear()
        self.evidence_inspector.set_evidence(None)
        self.diagnostics_widget.clear()
        self._exchange_events.clear()
        self._update_activity_count()
        self._update_source_counts()
        self.cockpit_tabs.setTabText(1, "⚡ Activity (0)")

        try:
            adapter = OllamaAgentAdapter()
        except Exception as exc:
            self._append_system_message(f"Could not initialize the local model: {exc}", error=True)
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
        self.activity_widget.set_running(True)
        # Switch to Research Activity tab during research so user sees live execution
        self.cockpit_tabs.setCurrentIndex(1)
        worker.start()

    def _on_stop(self) -> None:
        worker = self._worker
        if worker is None or not worker.isRunning():
            return
        self._set_status_text("Cancelling…")
        self.stop_btn.setEnabled(False)
        worker.request_cancel()

    def _on_research_started(self) -> None:
        self._set_status_text("Researching…")

    def _on_activity(self, activity: ResearchActivity) -> None:
        self._activities.append(activity)
        self._update_activity_count()
        self.activity_widget.add_activity(activity)

    def _on_evidence(self, evidence: Evidence) -> None:
        self._evidence[evidence.id] = evidence
        self._add_source_row(evidence)
        self._update_source_counts()
        # Automatically display the first retrieved evidence in the inspector
        if len(self._evidence) == 1:
            self.evidence_inspector.set_evidence(evidence)

    def _on_exchange(self, event: object) -> None:
        if not isinstance(event, dict):
            return
        event = dict(event)
        event["timestamp"] = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self._exchange_events.append(event)
        self.diagnostics_widget.add_exchange_event(event)

    def _on_result(self, result: AgentResult) -> None:
        self._last_result = result
        # Final result is authoritative in case an evidence signal was missed.
        for evidence in result.evidence:
            if evidence.id not in self._evidence:
                self._evidence[evidence.id] = evidence
                self._add_source_row(evidence)

        self._update_source_counts()
        self.activity_widget.set_running(False)

        # Switch to Sources tab so user can explore evidence citations immediately
        if self._evidence:
            self.cockpit_tabs.setCurrentIndex(0)
            if self.sources_tree.topLevelItemCount() > 0:
                first_item = self.sources_tree.topLevelItem(0)
                self.sources_tree.setCurrentItem(first_item)
                self._on_source_tree_item_clicked(first_item)

        # Compute KPI badge and update Diagnostics metrics
        stamp = datetime.now().strftime("%H:%M")
        msg_idx = len(self._conversation)

        total_duration = 0.0
        prompt_tokens = 0
        eval_tokens = 0
        eval_duration = 0.0

        if result.trace and result.trace.events:
            for ev in result.trace.events:
                if ev.get("type") == "model_call_finished":
                    total_duration += float(ev.get("duration_seconds", 0.0))
                    metrics = ev.get("metrics") or {}
                    prompt_tokens += int(metrics.get("prompt_eval_count") or 0)
                    eval_tokens += int(metrics.get("eval_count") or 0)
                    eval_duration += float(metrics.get("eval_duration") or 0) / 1e9

        self.diagnostics_widget.update_metrics(
            total_seconds=total_duration,
            prompt_tokens=prompt_tokens,
            eval_tokens=eval_tokens,
            eval_seconds=eval_duration,
        )

        kpi_parts = []
        if total_duration > 0:
            kpi_parts.append(f"⏱️ {total_duration:.1f}s")
        if self._activities:
            kpi_parts.append(f"{len(self._activities)} steps")
        if self._evidence:
            kpi_parts.append(f"{len(self._evidence)} sources")
        kpi_str = " · ".join(kpi_parts)

        card_html = render_assistant_message_card(
            result.answer,
            timestamp=stamp,
            valid_evidence_ids=set(self._evidence),
            msg_index=msg_idx,
            kpi_badge=kpi_str,
        )
        self.conversation_browser.append(card_html)
        self._conversation.append({
            "role": "assistant",
            "content": result.answer,
            "timestamp": stamp,
            "kpi": kpi_str,
        })

        if result.warnings:
            for w in result.warnings:
                self.conversation_browser.append(render_system_message_card(w, error=False))

        self._set_status_text("Ready")

    def _on_cancelled(self, result) -> None:
        self.activity_widget.set_running(False)
        self._append_system_message("Research stopped.")
        self._set_status_text("Stopped")

    def _on_failed(self, message: str, result) -> None:
        self.activity_widget.set_running(False)
        logger.error("Assistant research failed: %s", message)
        self._append_system_message(message, error=True)
        self._set_status_text("Failed")

    def _on_worker_finished(self) -> None:
        sender = self.sender()
        if sender is not None and sender is not self._worker:
            return
        worker = self._worker
        if worker is None:
            return
        self._set_running(False)
        self.activity_widget.set_running(False)
        worker.deleteLater()
        self._worker = None
        if self.status_label.text() == "Cancelling…":
            self._set_status_text("Stopped")

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
        stamp = datetime.now().strftime("%H:%M")
        if role == "user":
            card_html = render_user_message_card(content, timestamp=stamp)
            self.conversation_browser.append(card_html)
            self._conversation.append({"role": role, "content": content, "timestamp": stamp})

    def _append_system_message(self, message: str, error: bool = False) -> None:
        card_html = render_system_message_card(message, error=error)
        self.conversation_browser.append(card_html)

    def _render_answer(self, answer: str) -> str:
        return render_markdown_to_html(answer, set(self._evidence))

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

    def _update_activity_count(self) -> None:
        count = len(self._activities)
        self.cockpit_tabs.setTabText(1, f"⚡ Activity ({count})")
        self.activity_widget.update_count(count)
