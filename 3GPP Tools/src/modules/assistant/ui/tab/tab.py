"""Main AssistantTab class assembling UI, docking, session, and runner mixins."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QWidget

from modules.assistant.core.agent_models import AgentResult, Evidence, ResearchActivity
from modules.assistant.core.tools import build_read_only_tool_registry
from modules.assistant.ui.assistant_worker import AgentResearchWorker
from modules.assistant.ui.tab.mixins import (
    AssistantDockingLayoutMixin,
    AssistantResearchRunnerMixin,
    AssistantSessionExportMixin,
    AssistantUiSetupMixin,
)


class AssistantTab(
    AssistantUiSetupMixin,
    AssistantDockingLayoutMixin,
    AssistantSessionExportMixin,
    AssistantResearchRunnerMixin,
    QWidget,
):
    """Conversation + objective research activity + application-owned sources."""

    open_ollama_settings_requested = pyqtSignal()
    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)
    undock_requested = pyqtSignal()
    docked_changed = pyqtSignal(bool)

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
        self._is_docked: bool = False
        self._setup_ui()
        # Explicit idle state; independent of Ollama monitor state.
        self._set_running(False)
        self._set_status_text("Ready")
        self._refresh_ollama_status()
