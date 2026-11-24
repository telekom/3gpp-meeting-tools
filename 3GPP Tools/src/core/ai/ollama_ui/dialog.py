"""Configuration dialog for Ollama local LLM connection, proxy mode, and daemon control."""

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QDialog

from core.ai.ollama_client import OllamaClient, load_ollama_config
from core.ui.theme import C

from .mixins import OllamaActionsDiagnosticsMixin, OllamaUiSetupMixin


class OllamaConfigDialog(OllamaUiSetupMixin, OllamaActionsDiagnosticsMixin, QDialog):
    """Modal configuration dialog for local/remote Ollama connectivity, GPU flags, and diagnostics."""

    settings_saved = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🦙 Ollama LLM Configuration")
        self.setMinimumWidth(560)
        self.setStyleSheet(f"background-color: {C.BG_ALT_LIGHT};")

        self.cfg = load_ollama_config()
        self.client = OllamaClient(
            host=self.cfg.get("host"),
            proxy_mode=self.cfg.get("proxy_mode", "direct"),
        )
        self.service_worker = None
        self.diagnostics_worker = None
        self._close_after_service_action = False

        self._setup_ui()
        self._load_current_values()
