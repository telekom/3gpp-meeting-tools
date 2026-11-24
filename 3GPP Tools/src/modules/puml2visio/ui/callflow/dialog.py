"""Modeless CallFlowDialog orchestrator combining UI setup and generation mixins."""

from __future__ import annotations

from typing import Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog

from core.ai.ollama_client import OllamaClient
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.callflow_thread import (
    CallFlowGeneratorThread,
    CallFlowValidationThread,
)
from modules.puml2visio.core.prompt_manager import PromptManager
from modules.puml2visio.ui.callflow.mixins import (
    CallFlowGenerationMixin,
    CallFlowUiSetupMixin,
)
from modules.puml2visio.utils.paths import get_puml2visio_asset_path


class CallFlowDialog(
    QDialog,
    CallFlowUiSetupMixin,
    CallFlowGenerationMixin,
):
    """Modeless dialog for generating PlantUML diagrams from specification call flow text."""

    diagram_generated = pyqtSignal(str, bool)  # (puml_code, should_replace)

    def __init__(self, parent=None):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.setWindowTitle("🤖 Generate PlantUML from 3GPP Call Flow")
        self.resize(1120, 740)

        # Modeless behavior & top-level window controls
        self.setModal(False)
        self.setWindowFlags(
            Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint
        )

        self.client = OllamaClient()
        self.prompt_mgr = PromptManager()
        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)

        self.worker: Optional[CallFlowGeneratorThread] = None
        self.validator: Optional[CallFlowValidationThread] = None
        self.conversation_history: List[Dict[str, str]] = []
        self.last_clean_puml: str = ""
        self.last_syntax_error: str = ""
        self.last_error_line: int = 0

        self._setup_ui()
        self._refresh_models()

    def closeEvent(self, event):
        worker = getattr(self, "worker", None)
        if worker and worker.isRunning():
            worker.cancel()
            if not worker.wait(500):
                worker.terminate()
                worker.wait(200)
        validator = getattr(self, "validator", None)
        if validator and validator.isRunning():
            if not validator.wait(500):
                validator.terminate()
                validator.wait(200)
        event.accept()
