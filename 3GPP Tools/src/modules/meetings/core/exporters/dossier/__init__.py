"""Modular Unified Navigable Dossier Exporter package (HTML + PDF + AI Prompt)."""

from .css import DOSSIER_CSS
from .exporter_thread import DossierExporterThread, logger as _thread_logger
from .html_builder import build_html_dossier, convert_markdown_to_html
from .pdf_renderer import render_bookmarked_pdf
from .prompt_builder import generate_ai_prompt

logger = _thread_logger

__all__ = [
    "DOSSIER_CSS",
    "DossierExporterThread",
    "build_html_dossier",
    "convert_markdown_to_html",
    "generate_ai_prompt",
    "render_bookmarked_pdf",
]
