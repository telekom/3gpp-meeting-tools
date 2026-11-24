"""Prompt editor, starter query chips, and chat markdown/card renderers."""

from __future__ import annotations

import html
import re
from typing import Optional, Set

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import QHBoxLayout, QLabel, QPlainTextEdit, QPushButton, QVBoxLayout, QWidget

from core.ui.theme import C
from modules.assistant.ui_components.styles import (
    PROMPT_INPUT_STYLE,
    STARTER_CHIP_STYLE,
    STARTER_PROMPTS,
)


class MultilinePromptEdit(QPlainTextEdit):
    """Auto-wrapping multi-line prompt editor.

    Submits on Enter (Return); Shift+Enter inserts a newline.
    """

    submit_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setStyleSheet(PROMPT_INPUT_STYLE)
        self.setPlaceholderText(
            "Ask about 3GPP specifications, procedures, messages or IEs... "
            "(Press Enter to send, Shift+Enter for newline)"
        )
        self.setMinimumHeight(48)
        self.setMaximumHeight(140)
        self.setTabChangesFocus(True)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if bool(event.modifiers() & Qt.ShiftModifier):
                super().keyPressEvent(event)
            else:
                event.accept()
                self.submit_requested.emit()
            return
        super().keyPressEvent(event)


class StarterChipsWidget(QWidget):
    """Clickable prompt suggestion chips shown when the conversation is empty."""

    prompt_selected = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 16, 12, 16)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)

        title = QLabel("💡 <b>Suggested 3GPP Research Queries</b>")
        title.setStyleSheet(f"font-size: 13px; color: {C.TEXT_SECONDARY};")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        chip_container = QHBoxLayout()
        chip_container.setSpacing(8)
        chip_container.setAlignment(Qt.AlignCenter)

        for label, query in STARTER_PROMPTS:
            btn = QPushButton(label)
            btn.setStyleSheet(STARTER_CHIP_STYLE)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip(f'Click to ask:\n"{query}"')
            btn.clicked.connect(lambda checked=False, q=query: self.prompt_selected.emit(q))
            chip_container.addWidget(btn)

        layout.addLayout(chip_container)


def render_markdown_to_html(md: str, valid_evidence_ids: Optional[Set[str]] = None) -> str:
    """Renders basic technical Markdown to sanitized HTML with interactive citations."""
    valid_ids = set(valid_evidence_ids or set())

    # 1. Extract fenced code blocks
    code_blocks = []

    def _save_code(match: re.Match) -> str:
        lang = match.group(1) or ""
        code = match.group(2)
        code_blocks.append((lang, code))
        return f"__CODE_BLOCK_{len(code_blocks)-1}__"

    text = re.sub(r"```([a-zA-Z0-9_-]*)\n?(.*?)```", _save_code, md or "", flags=re.DOTALL)

    # 2. Escape HTML for ordinary text
    text = html.escape(text)

    # 3. Tables (simple markdown table handling)
    def _render_table(match: re.Match) -> str:
        raw_table = match.group(0).strip()
        rows = [r.strip() for r in raw_table.split("\n") if r.strip()]
        if len(rows) < 2:
            return raw_table
        headers = [c.strip() for c in rows[0].strip("|").split("|")]
        data_rows = rows[2:] if ("---" in rows[1] or ":---" in rows[1]) else rows[1:]

        th_html = "".join(
            f'<th style="background:{C.BG_SURFACE}; color:{C.TEXT_SECONDARY}; font-weight:bold; font-size:12px; border:1px solid {C.BORDER_STRONG}; padding:5px 8px; text-align:left;">{h}</th>'
            for h in headers
        )
        tr_rows = []
        for r in data_rows:
            cols = [c.strip() for c in r.strip("|").split("|")]
            td_html = "".join(
                f'<td style="border:1px solid {C.BORDER_SUBTLE}; padding:4px 8px; font-size:12px; color:{C.TEXT_STRONG};">{c}</td>'
                for c in cols
            )
            tr_rows.append(f"<tr>{td_html}</tr>")

        return f'<table style="border-collapse:collapse; width:100%; margin:8px 0;"><thead><tr>{th_html}</tr></thead><tbody>{"".join(tr_rows)}</tbody></table>'

    text = re.sub(
        r"(?:^\|.+\|\r?\n(?:^\|[-:\s|]+\|\r?\n)?(?:^\|.+\|\r?\n?)+)",
        _render_table,
        text,
        flags=re.MULTILINE,
    )

    # 4. Headings
    text = re.sub(
        r"^###\s+(.+)$",
        fr'<h4 style="margin:8px 0 4px 0; color:{C.TEXT_STRONG}; font-size:13px; font-weight:bold;">\1</h4>',
        text,
        flags=re.MULTILINE,
    )
    text = re.sub(
        r"^##\s+(.+)$",
        fr'<h3 style="margin:10px 0 4px 0; color:{C.TEXT_STRONG}; font-size:14px; font-weight:bold;">\1</h3>',
        text,
        flags=re.MULTILINE,
    )
    text = re.sub(
        r"^#\s+(.+)$",
        fr'<h2 style="margin:12px 0 6px 0; color:{C.TEXT_STRONG}; font-size:15px; font-weight:bold;">\1</h2>',
        text,
        flags=re.MULTILINE,
    )

    # 5. Bold, Italic, and Inline code
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*([^*]+?)\*(?!\*)", r"<i>\1</i>", text)
    text = re.sub(
        r"`([^`]+)`",
        fr'<code style="background:{C.BG_CANVAS}; color:{C.TEXT_PRIMARY}; padding:1px 4px; border-radius:3px; font-family:Consolas,monospace; font-size:12px;">\1</code>',
        text,
    )

    # 6. Unordered and Ordered Lists
    lines = text.split("\n")
    out_lines = []
    in_ul = False
    in_ol = False
    for line in lines:
        m_ul = re.match(r"^\s*[-*]\s+(.+)$", line)
        m_ol = re.match(r"^\s*\d+\.\s+(.+)$", line)
        if m_ul:
            if not in_ul:
                out_lines.append('<ul style="margin:4px 0 6px 18px; padding:0;">')
                in_ul = True
            out_lines.append(f"<li>{m_ul.group(1)}</li>")
        elif m_ol:
            if not in_ol:
                out_lines.append('<ol style="margin:4px 0 6px 18px; padding:0;">')
                in_ol = True
            out_lines.append(f"<li>{m_ol.group(1)}</li>")
        else:
            if in_ul:
                out_lines.append("</ul>")
                in_ul = False
            if in_ol:
                out_lines.append("</ol>")
                in_ol = False
            out_lines.append(line)
    if in_ul:
        out_lines.append("</ul>")
    if in_ol:
        out_lines.append("</ol>")
    text = "\n".join(out_lines)

    # 7. Interactive Evidence Citations
    def _highlight_citations(match: re.Match) -> str:
        inner = match.group(1)
        found_eids = re.findall(r"\bE\d+\b", inner)
        if not any(eid in valid_ids for eid in found_eids):
            return match.group(0)

        h = re.sub(
            r"\b(E\d+)\b",
            lambda m: (
                f'<a href="evidence:{m.group(1)}" style="color:{C.ACCENT_SKY_DARK}; text-decoration:none; font-weight:bold; background:{C.SKY_SOFT}; padding:1px 5px; border-radius:4px; border:1px solid {C.SKY_BORDER};">{m.group(1)}</a>'
                if m.group(1) in valid_ids
                else m.group(1)
            ),
            inner,
        )
        return f'<span style="color:{C.ACCENT_SKY_DARK}; font-weight:bold;">[</span>{h}<span style="color:{C.ACCENT_SKY_DARK}; font-weight:bold;">]</span>'

    text = re.sub(r"\[([^\[\]]*?\bE\d+\b[^\[\]]*?)\]", _highlight_citations, text)

    # 8. Restore Code Blocks
    for i, (_lang, code) in enumerate(code_blocks):
        escaped_code = html.escape(code.strip())
        block_html = (
            f'<div style="margin:8px 0; background:{C.BG_SURFACE}; border-radius:6px; padding:8px 12px; border:1px solid {C.BORDER_STRONG};">'
            f'<pre style="margin:0; font-family:Consolas,monospace; font-size:12px; color:{C.TEXT_STRONG}; white-space:pre-wrap;">{escaped_code}</pre>'
            f"</div>"
        )
        text = text.replace(f"__CODE_BLOCK_{i}__", block_html)

    # 9. Clean Linebreaks
    return text.replace("\n", "<br>")


def render_user_message_card(content: str, timestamp: str = "") -> str:
    """Renders a user question into a soft blue message card."""
    escaped = html.escape(content).replace("\n", "<br>")
    time_str = timestamp or ""
    return f"""
    <div style="margin: 8px 0; padding: 10px 14px; background-color: {C.BLUE_TINT_SOFT}; border: 1px solid {C.SKY_BORDER}; border-radius: 8px;">
      <div style="margin-bottom: 5px;">
        <span style="font-weight: bold; color: {C.ACCENT_SKY_DARK}; font-size: 12px;">👤 You</span>
        <span style="color: {C.TEXT_FAINT}; font-size: 11px; float: right;">{time_str}</span>
      </div>
      <div style="color: {C.TEXT_PRIMARY}; font-size: 13px; line-height: 1.45; white-space: pre-wrap;">{escaped}</div>
    </div>
    """


def render_assistant_message_card(
    answer: str,
    timestamp: str = "",
    valid_evidence_ids: Optional[Set[str]] = None,
    msg_index: int = 0,
    kpi_badge: str = "",
) -> str:
    """Renders a grounded Assistant response into a styled card with markdown and clickable citations."""
    rendered_body = render_markdown_to_html(answer, valid_evidence_ids)
    time_str = timestamp or ""
    kpi_html = (
        f'<span style="background:{C.BG_CANVAS}; color:{C.TEXT_SECONDARY}; font-size:11px; padding:1px 6px; border-radius:4px; margin-left:8px; border:1px solid {C.BORDER_SUBTLE};">{kpi_badge}</span>'
        if kpi_badge
        else ""
    )

    return f"""
    <div style="margin: 8px 0; padding: 12px 16px; background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px;">
      <div style="margin-bottom: 8px; border-bottom: 1px solid {C.BG_CANVAS}; padding-bottom: 6px;">
        <span style="font-weight: bold; color: {C.TEXT_STRONG}; font-size: 13px;">🤖 3GPP Assistant</span>
        {kpi_html}
        <span style="color: {C.TEXT_FAINT}; font-size: 11px; float: right;">{time_str}</span>
      </div>
      <div style="color: {C.TEXT_STRONG}; font-size: 13px; line-height: 1.5;">{rendered_body}</div>
      <div style="margin-top: 10px; padding-top: 6px; border-top: 1px solid {C.BG_SURFACE}; text-align: right;">
        <a href="copy_msg:{msg_index}" style="color: {C.TEXT_MUTED}; font-size: 11px; text-decoration: none; background-color: {C.BG_SURFACE}; padding: 3px 8px; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px;">📋 Copy Answer</a>
      </div>
    </div>
    """


def render_system_message_card(message: str, error: bool = False) -> str:
    """Renders system notifications or errors as clean colored notice banners."""
    escaped = html.escape(message)
    bg = C.DANGER_SOFT if error else C.BG_SURFACE
    border = C.DANGER_BORDER if error else C.BORDER_SUBTLE
    color = C.DANGER if error else C.TEXT_SECONDARY
    icon = "⚠️" if error else "ℹ️"
    return f"""
    <div style="margin: 6px 0; padding: 8px 12px; background-color: {bg}; border: 1px solid {border}; border-radius: 6px; color: {color}; font-size: 12px;">
      {icon} {escaped}
    </div>
    """
