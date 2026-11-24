# --- File: src/modules/meetings/core/session_plan/parser/parser.py ---
"""Composed SessionPlanParser combining Markdown, OpenXML, and Agenda Item parsing."""

from .agenda_items import AgendaItemsParserMixin
from .docx_table import DocxTableParserMixin
from .markdown_parser import MarkdownParserMixin


class SessionPlanParser(MarkdownParserMixin, DocxTableParserMixin, AgendaItemsParserMixin):
    """Parses meeting session plan tables and extracts semantic schedule models."""
