# --- File: src/core/ai/ollama_dialog.py ---
"""Backward-compatible facade re-exporting `core.ai.ollama_ui`."""

from core.ai.ollama_ui import OllamaConfigDialog

__all__ = ["OllamaConfigDialog"]
