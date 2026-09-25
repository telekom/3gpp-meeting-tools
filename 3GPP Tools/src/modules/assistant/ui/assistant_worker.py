"""QThread execution adapter for one Assistant research run."""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Mapping, Optional, Sequence

from PyQt5.QtCore import QThread, pyqtSignal

from modules.assistant.core.agent_controller import (
    AgentController,
    AgentControllerConfig,
)
from modules.assistant.core.agent_models import (
    Evidence,
    ResearchActivity,
    ResearchOutcome,
)
from modules.assistant.core.ollama_agent_adapter import OllamaAgentAdapter
from modules.assistant.core.tool_registry import ToolRegistry

logger = logging.getLogger(__name__)


class AgentResearchWorker(QThread):
    """Runs exactly one sequential AgentController research run.

    The worker is intentionally thin.  Agent policy remains in AgentController;
    this class owns Qt signals, lifecycle, and the exception boundary.
    """

    research_started = pyqtSignal()
    activity_updated = pyqtSignal(object)   # ResearchActivity
    evidence_added = pyqtSignal(object)     # Evidence
    exchange_observed = pyqtSignal(object)  # Raw Ollama request/response event
    result_ready = pyqtSignal(object)       # AgentResult
    research_cancelled = pyqtSignal(object) # AgentResult
    research_failed = pyqtSignal(str, object)  # user-facing error, AgentResult|None

    def __init__(
        self,
        *,
        question: str,
        tool_registry: ToolRegistry,
        adapter: Optional[OllamaAgentAdapter] = None,
        conversation_context: Optional[Sequence[Mapping[str, str]]] = None,
        controller_config: Optional[AgentControllerConfig] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.question = str(question or "").strip()
        self.tool_registry = tool_registry
        self.adapter = adapter or OllamaAgentAdapter()
        # OllamaClient invokes this callback in this worker thread. Re-emitting
        # through a Qt signal safely queues the event for the GUI thread.
        self.adapter.client.exchange_callback = self._on_exchange
        self.conversation_context: List[Dict[str, str]] = [
            dict(item) for item in (conversation_context or [])
        ]
        self.controller_config = controller_config or AgentControllerConfig()
        self._controller: Optional[AgentController] = None
        self._cancel_requested = False

    def request_cancel(self) -> None:
        """Cooperatively cancel this run.

        If an Ollama HTTP request is currently blocking, the request is not
        force-terminated.  Its result will be discarded at the controller's
        next cancellation checkpoint.
        """
        self._cancel_requested = True
        controller = self._controller
        if controller is not None:
            controller.cancel_active()

    @property
    def cancellation_requested(self) -> bool:
        return self._cancel_requested

    def run(self) -> None:
        self.research_started.emit()
        try:
            controller = AgentController(
                adapter=self.adapter,
                tools=self.tool_registry,
                config=self.controller_config,
                activity_callback=self._on_activity,
                evidence_callback=self._on_evidence,
            )
            self._controller = controller

            # Cover cancellation requested after start() but before the
            # controller has created its ResearchRun.
            if self._cancel_requested:
                self._emit_pre_run_cancelled()
                return

            result = controller.run(
                self.question,
                conversation_context=self.conversation_context,
            )

            if result.outcome == ResearchOutcome.COMPLETED:
                self.result_ready.emit(result)
            elif result.outcome == ResearchOutcome.CANCELLED:
                self.research_cancelled.emit(result)
            else:
                self.research_failed.emit(
                    result.error or "Assistant research failed.",
                    result,
                )

        except Exception as exc:
            # Nothing from Ollama/tools/controller may escape into the Qt event
            # loop.  Full detail goes to the application log.
            logger.exception("Unhandled exception in AgentResearchWorker")
            self.research_failed.emit(
                f"Assistant research failed: {exc}",
                None,
            )
        finally:
            self._controller = None

    def _on_activity(self, activity: ResearchActivity) -> None:
        self.activity_updated.emit(activity)

    def _on_evidence(self, evidence: Evidence) -> None:
        self.evidence_added.emit(evidence)

    def _on_exchange(self, event: object) -> None:
        self.exchange_observed.emit(event)

    def _emit_pre_run_cancelled(self) -> None:
        # No ResearchRun exists yet, so there is no AgentResult to expose.
        self.research_cancelled.emit(None)
