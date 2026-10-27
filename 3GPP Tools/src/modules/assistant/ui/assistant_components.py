"""Reusable UI components, prompt editor, and style presets for the 3GPP Assistant."""

from __future__ import annotations

import html
import json
import re
from typing import Any, Dict, List, Optional, Sequence, Set
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
    TREE_STYLE_CLEAN,
)
from modules.assistant.core.agent_models import Evidence, EvidenceSourceType, ResearchActivity

# =========================================================================
# Stylesheet Constants
# =========================================================================

PROMPT_INPUT_STYLE = """
QPlainTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 6px 8px;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
    color: #1E293B;
}
QPlainTextEdit:focus {
    border: 1.5px solid #1E5C99;
}
"""

STATUS_PILL_ONLINE = """
QPushButton {
    background-color: #ECFDF5;
    color: #047857;
    border: 1px solid #A7F3D0;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #D1FAE5;
    border-color: #6EE7B7;
}
"""

STATUS_PILL_OFFLINE = """
QPushButton {
    background-color: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #FEE2E2;
    border-color: #FCA5A5;
}
"""

STATUS_PILL_NEUTRAL = """
QPushButton {
    background-color: #F1F5F9;
    color: #475569;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #E2E8F0;
}
"""

STARTER_CHIP_STYLE = """
QPushButton {
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    text-align: left;
}
QPushButton:hover {
    background-color: #EBF3FC;
    color: #1E5C99;
    border-color: #1E5C99;
}
QPushButton:pressed {
    background-color: #DFEDFA;
}
"""

CARD_FRAME_STYLE = """
QFrame#assistantCardFrame {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
}
"""

CARD_HEADER_STYLE = """
QFrame#assistantCardHeader {
    background-color: #F8FAFC;
    border-bottom: 1px solid #E2E8F0;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
}
"""

BADGE_SPEC_TEXT = "background-color: #EEF2FF; color: #4338CA; border: 1px solid #C7D2FE; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"
BADGE_PROTOCOL = "background-color: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"
BADGE_CATALOGUE = "background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"


# =========================================================================
# Multi-line Prompt Editor
# =========================================================================

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


# =========================================================================
# Starter Prompt Chips Widget (Empty State)
# =========================================================================

STARTER_PROMPTS = [
    ("🔍 TS 23.501 Registration", "What is the Registration procedure in TS 23.501?"),
    ("📨 NGAP PDU Session Setup", "What IEs are in the NGAP PDU Session Resource Setup Request?"),
    ("🔬 TS 29.244 (PFCP) Messages", "List all messages defined for TS 29.244 (PFCP)"),
    ("📋 TS 38.331 RRCReconfiguration", "What is the structure of the RRCReconfiguration message in TS 38.331?"),
]


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
        title.setStyleSheet("font-size: 13px; color: #475569;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        chip_container = QHBoxLayout()
        chip_container.setSpacing(8)
        chip_container.setAlignment(Qt.AlignCenter)

        for label, query in STARTER_PROMPTS:
            btn = QPushButton(label)
            btn.setStyleSheet(STARTER_CHIP_STYLE)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip(f"Click to ask:\n\"{query}\"")
            btn.clicked.connect(lambda checked=False, q=query: self.prompt_selected.emit(q))
            chip_container.addWidget(btn)

        layout.addLayout(chip_container)


# =========================================================================
# Markdown & Message Card Rendering Engine
# =========================================================================

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
            f'<th style="background:#F8FAFC; color:#475569; font-weight:bold; font-size:12px; border:1px solid #CBD5E1; padding:5px 8px; text-align:left;">{h}</th>'
            for h in headers
        )
        tr_rows = []
        for r in data_rows:
            cols = [c.strip() for c in r.strip("|").split("|")]
            td_html = "".join(
                f'<td style="border:1px solid #E2E8F0; padding:4px 8px; font-size:12px; color:#1E293B;">{c}</td>'
                for c in cols
            )
            tr_rows.append(f"<tr>{td_html}</tr>")

        return f'<table style="border-collapse:collapse; width:100%; margin:8px 0;"><thead><tr>{th_html}</tr></thead><tbody>{"".join(tr_rows)}</tbody></table>'

    text = re.sub(r"(?:^\|.+\|\r?\n(?:^\|[-:\s|]+\|\r?\n)?(?:^\|.+\|\r?\n?)+)", _render_table, text, flags=re.MULTILINE)

    # 4. Headings
    text = re.sub(r"^###\s+(.+)$", r'<h4 style="margin:8px 0 4px 0; color:#1E293B; font-size:13px; font-weight:bold;">\1</h4>', text, flags=re.MULTILINE)
    text = re.sub(r"^##\s+(.+)$", r'<h3 style="margin:10px 0 4px 0; color:#1E293B; font-size:14px; font-weight:bold;">\1</h3>', text, flags=re.MULTILINE)
    text = re.sub(r"^#\s+(.+)$", r'<h2 style="margin:12px 0 6px 0; color:#1E293B; font-size:15px; font-weight:bold;">\1</h2>', text, flags=re.MULTILINE)

    # 5. Bold, Italic, and Inline code
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*([^*]+?)\*(?!\*)", r"<i>\1</i>", text)
    text = re.sub(
        r"`([^`]+)`",
        r'<code style="background:#F1F5F9; color:#0F172A; padding:1px 4px; border-radius:3px; font-family:Consolas,monospace; font-size:12px;">\1</code>',
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
                f'<a href="evidence:{m.group(1)}" style="color:#0369A1; text-decoration:none; font-weight:bold; background:#E0F2FE; padding:1px 5px; border-radius:4px; border:1px solid #BAE6FD;">{m.group(1)}</a>'
                if m.group(1) in valid_ids
                else m.group(1)
            ),
            inner,
        )
        return f'<span style="color:#0369A1; font-weight:bold;">[</span>{h}<span style="color:#0369A1; font-weight:bold;">]</span>'

    text = re.sub(r"\[([^\[\]]*?\bE\d+\b[^\[\]]*?)\]", _highlight_citations, text)

    # 8. Restore Code Blocks
    for i, (lang, code) in enumerate(code_blocks):
        escaped_code = html.escape(code.strip())
        block_html = (
            f'<div style="margin:8px 0; background:#F8FAFC; border-radius:6px; padding:8px 12px; border:1px solid #CBD5E1;">'
            f'<pre style="margin:0; font-family:Consolas,monospace; font-size:12px; color:#1E293B; white-space:pre-wrap;">{escaped_code}</pre>'
            f"</div>"
        )
        text = text.replace(f"__CODE_BLOCK_{i}__", block_html)

    # 9. Clean Linebreaks
    text = text.replace("\n", "<br>")
    return text


def render_user_message_card(content: str, timestamp: str = "") -> str:
    """Renders a user question into a soft blue message card."""
    escaped = html.escape(content).replace("\n", "<br>")
    time_str = timestamp or ""
    return f"""
    <div style="margin: 8px 0; padding: 10px 14px; background-color: #F0F7FF; border: 1px solid #BAE6FD; border-radius: 8px;">
      <div style="margin-bottom: 5px;">
        <span style="font-weight: bold; color: #0369A1; font-size: 12px;">👤 You</span>
        <span style="color: #94A3B8; font-size: 11px; float: right;">{time_str}</span>
      </div>
      <div style="color: #0F172A; font-size: 13px; line-height: 1.45; white-space: pre-wrap;">{escaped}</div>
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
    kpi_html = f'<span style="background:#F1F5F9; color:#475569; font-size:11px; padding:1px 6px; border-radius:4px; margin-left:8px; border:1px solid #E2E8F0;">{kpi_badge}</span>' if kpi_badge else ""

    return f"""
    <div style="margin: 8px 0; padding: 12px 16px; background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px;">
      <div style="margin-bottom: 8px; border-bottom: 1px solid #F1F5F9; padding-bottom: 6px;">
        <span style="font-weight: bold; color: #1E293B; font-size: 13px;">🤖 3GPP Assistant</span>
        {kpi_html}
        <span style="color: #94A3B8; font-size: 11px; float: right;">{time_str}</span>
      </div>
      <div style="color: #1E293B; font-size: 13px; line-height: 1.5;">{rendered_body}</div>
      <div style="margin-top: 10px; padding-top: 6px; border-top: 1px solid #F8FAFC; text-align: right;">
        <a href="copy_msg:{msg_index}" style="color: #64748B; font-size: 11px; text-decoration: none; background-color: #F8FAFC; padding: 3px 8px; border: 1px solid #E2E8F0; border-radius: 4px;">📋 Copy Answer</a>
      </div>
    </div>
    """


def render_system_message_card(message: str, error: bool = False) -> str:
    """Renders system notifications or errors as clean colored notice banners."""
    escaped = html.escape(message)
    bg = "#FEF2F2" if error else "#F8FAFC"
    border = "#FECACA" if error else "#E2E8F0"
    color = "#B91C1C" if error else "#475569"
    icon = "⚠️" if error else "ℹ️"
    return f"""
    <div style="margin: 6px 0; padding: 8px 12px; background-color: {bg}; border: 1px solid {border}; border-radius: 6px; color: {color}; font-size: 12px;">
      {icon} {escaped}
    </div>
    """


# =========================================================================
# Embedded Evidence Inspector Pane
# =========================================================================

class EvidenceInspectorWidget(QFrame):
    """Embedded inspector previewing retrieved clause/protocol evidence with deep links."""

    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("assistantCardFrame")
        self.setStyleSheet(CARD_FRAME_STYLE)
        self._current_evidence: Optional[Evidence] = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header toolbar
        hdr = QHBoxLayout()
        hdr.setContentsMargins(0, 0, 0, 0)
        hdr.setSpacing(6)

        self.badge = QLabel("Evidence")
        self.badge.setStyleSheet(BADGE_CATALOGUE)
        hdr.addWidget(self.badge)

        self.title_label = QLabel("No evidence selected")
        self.title_label.setStyleSheet("font-weight: bold; color: #1E293B; font-size: 12px;")
        hdr.addWidget(self.title_label, stretch=1)

        self.spec_search_btn = QPushButton("🔎 Spec Search")
        self.spec_search_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.spec_search_btn.setToolTip("Open and navigate to this clause in Spec Search tab")
        self.spec_search_btn.clicked.connect(self._on_open_spec_search)
        self.spec_search_btn.setVisible(False)
        hdr.addWidget(self.spec_search_btn)

        self.protocol_btn = QPushButton("🔬 In Protocols")
        self.protocol_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.protocol_btn.setToolTip("Inspect this message in Protocols tab")
        self.protocol_btn.clicked.connect(self._on_open_protocol)
        self.protocol_btn.setVisible(False)
        hdr.addWidget(self.protocol_btn)

        self.specs_tab_btn = QPushButton("📚 In Specs")
        self.specs_tab_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.specs_tab_btn.setToolTip("View specification metadata in Specifications tab")
        self.specs_tab_btn.clicked.connect(self._on_open_specs)
        self.specs_tab_btn.setVisible(False)
        hdr.addWidget(self.specs_tab_btn)

        self.copy_btn = QPushButton("📋 Copy Excerpt")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._on_copy_excerpt)
        self.copy_btn.setEnabled(False)
        hdr.addWidget(self.copy_btn)

        layout.addLayout(hdr)

        # Excerpt Browser
        self.content_browser = QTextBrowser()
        self.content_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                color: #1E293B;
                padding: 6px;
            }
        """)
        self.content_browser.setPlaceholderText(
            "Select an evidence item from the list above or click an [E#] citation in the chat to view its retrieved content."
        )
        layout.addWidget(self.content_browser, stretch=1)

    def set_evidence(self, ev: Optional[Evidence]) -> None:
        self._current_evidence = ev
        if ev is None:
            self.badge.setText("Evidence")
            self.badge.setStyleSheet(BADGE_CATALOGUE)
            self.title_label.setText("No evidence selected")
            self.content_browser.clear()
            self.spec_search_btn.setVisible(False)
            self.protocol_btn.setVisible(False)
            self.specs_tab_btn.setVisible(False)
            self.copy_btn.setEnabled(False)
            return

        # Badge
        if ev.source_type == EvidenceSourceType.SPECIFICATION_TEXT:
            self.badge.setText("Spec Text")
            self.badge.setStyleSheet(BADGE_SPEC_TEXT)
        elif ev.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE:
            self.badge.setText("Protocol")
            self.badge.setStyleSheet(BADGE_PROTOCOL)
        else:
            self.badge.setText("Catalogue")
            self.badge.setStyleSheet(BADGE_CATALOGUE)

        # Title
        parts = [ev.id]
        if ev.specification:
            parts.append(f"TS {ev.specification}")
        if ev.version:
            parts.append(f"v{ev.version}")
        if ev.clause:
            parts.append(f"Clause {ev.clause}")
        elif ev.title:
            parts.append(ev.title)
        self.title_label.setText(" — ".join(parts))

        # Deep link buttons visibility
        self.spec_search_btn.setVisible(bool(ev.specification and ev.clause))
        self.specs_tab_btn.setVisible(bool(ev.specification))
        self.protocol_btn.setVisible(bool(ev.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE))
        self.copy_btn.setEnabled(True)

        # Content rendering
        content = ev.content
        if isinstance(content, dict):
            text = json.dumps(content, indent=2, ensure_ascii=False, default=str)
        else:
            text = str(content or "")
        self.content_browser.setPlainText(text)

    def _on_copy_excerpt(self) -> None:
        text = self.content_browser.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            orig = self.copy_btn.text()
            self.copy_btn.setText("✓ Copied!")
            QTimer.singleShot(1500, lambda: self.copy_btn.setText(orig))

    def _on_open_spec_search(self) -> None:
        if self._current_evidence and self._current_evidence.specification:
            spec = self._current_evidence.specification
            ver = self._current_evidence.version or ""
            clause = self._current_evidence.clause or ""
            self.open_spec_search_requested.emit(spec, ver, clause)

    def _on_open_specs(self) -> None:
        if self._current_evidence and self._current_evidence.specification:
            spec = self._current_evidence.specification
            ver = self._current_evidence.version or ""
            self.open_specs_requested.emit(spec, ver)

    def _on_open_protocol(self) -> None:
        if self._current_evidence and self._current_evidence.content:
            proto = ""
            msg = ""
            if isinstance(self._current_evidence.content, dict):
                proto = str(self._current_evidence.content.get("protocol") or "")
                msg = str(self._current_evidence.content.get("message") or "")
            self.open_protocol_requested.emit(proto, msg)


# =========================================================================
# Visual Research Activity Timeline Widget
# =========================================================================

TOOL_ICONS = {
    "search_spec_text": "🔍",
    "get_spec_clause": "📖",
    "list_protocols": "🔬",
    "list_protocol_messages": "📨",
    "find_protocol_message": "🔬",
    "find_protocol_ie": "🏷️",
    "find_specifications": "📚",
    "get_specification": "📚",
    "research_started": "🚀",
    "answer_ready": "✅",
    "research_cancelled": "⏹️",
    "research_failed": "❌",
}


class ActivityTimelineWidget(QWidget):
    """Visual timeline displaying real-time research steps with semantic icons."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        # Header with pulse bar and step count
        header = QHBoxLayout()
        header.setContentsMargins(2, 0, 2, 0)
        self.step_label = QLabel("0 steps executed")
        self.step_label.setStyleSheet("font-weight: 600; color: #475569; font-size: 11px;")
        header.addWidget(self.step_label)
        header.addStretch()

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_btn.clicked.connect(self.clear)
        header.addWidget(self.clear_btn)
        layout.addLayout(header)

        # Progress bar (visible during active research)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 0)  # Indeterminate pulse
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #E2E8F0;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #1E5C99;
                border-radius: 2px;
            }
        """)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Browser
        self.browser = QTextBrowser()
        self.browser.setStyleSheet("""
            QTextBrowser {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 12px;
                color: #1E293B;
                padding: 6px;
            }
        """)
        layout.addWidget(self.browser, stretch=1)

    def set_running(self, running: bool) -> None:
        self.progress_bar.setVisible(running)

    def add_activity(self, activity: ResearchActivity) -> None:
        icon = TOOL_ICONS.get(activity.kind, "•")
        tool = activity.metadata.get("tool", "")
        if tool in TOOL_ICONS:
            icon = TOOL_ICONS[tool]

        msg = html.escape(activity.message)
        eids = activity.metadata.get("evidence_ids") or []
        eid_html = ""
        if eids:
            eid_chips = " ".join(
                f'<span style="background:#E0F2FE; color:#0369A1; font-weight:bold; padding:1px 4px; border-radius:3px; font-size:10px;">{eid}</span>'
                for eid in eids
            )
            eid_html = f" &nbsp;{eid_chips}"

        item_html = f"""
        <div style="margin: 3px 0; padding: 4px 6px; border-left: 2px solid #CBD5E1; background-color: #F8FAFC;">
          <span style="font-size: 13px;">{icon}</span> <b>{msg}</b>{eid_html}
        </div>
        """
        self.browser.append(item_html)

    def update_count(self, count: int) -> None:
        self.step_label.setText(f"{count} step" + ("s" if count != 1 else "") + " executed")

    def clear(self) -> None:
        self.browser.clear()
        self.update_count(0)
        self.set_running(False)


# =========================================================================
# Diagnostics & Ollama Trace Cockpit Widget
# =========================================================================

class DiagnosticsTraceWidget(QWidget):
    """Detailed diagnostics panel for raw Ollama exchange trace with KPI metrics."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        # KPI Metrics Strip
        kpi_card = QFrame()
        kpi_card.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px 8px;
            }
        """)
        kpi_layout = QHBoxLayout(kpi_card)
        kpi_layout.setContentsMargins(4, 2, 4, 2)
        kpi_layout.setSpacing(12)

        self.kpi_latency = QLabel("⏱️ Latency: 0.0s")
        self.kpi_latency.setStyleSheet("color: #475569; font-size: 11px; font-weight: 600;")
        kpi_layout.addWidget(self.kpi_latency)

        self.kpi_prompt_tokens = QLabel("📥 Prompt: 0 tok")
        self.kpi_prompt_tokens.setStyleSheet("color: #475569; font-size: 11px;")
        kpi_layout.addWidget(self.kpi_prompt_tokens)

        self.kpi_eval_tokens = QLabel("📤 Eval: 0 tok")
        self.kpi_eval_tokens.setStyleSheet("color: #475569; font-size: 11px;")
        kpi_layout.addWidget(self.kpi_eval_tokens)

        self.kpi_speed = QLabel("⚡ Speed: 0 tok/s")
        self.kpi_speed.setStyleSheet("color: #047857; font-size: 11px; font-weight: 600;")
        kpi_layout.addWidget(self.kpi_speed)

        kpi_layout.addStretch()

        self.copy_btn = QPushButton("Copy Trace")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._on_copy_trace)
        kpi_layout.addWidget(self.copy_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_btn.clicked.connect(self.clear)
        kpi_layout.addWidget(self.clear_btn)

        layout.addWidget(kpi_card)

        # Raw Trace Browser
        self.browser = QPlainTextEdit()
        self.browser.setReadOnly(True)
        self.browser.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.browser.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1E293B;
                color: #E2E8F0;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 11px;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        self.browser.setPlaceholderText("Raw Ollama REST requests and responses will be recorded here.")
        layout.addWidget(self.browser, stretch=1)

    def add_exchange_event(self, event: dict) -> None:
        stamp = event.get("timestamp") or ""
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

        self.browser.appendPlainText(f"[{stamp}] {label}\n{rendered}\n")

    def update_metrics(self, total_seconds: float, prompt_tokens: int, eval_tokens: int, eval_seconds: float) -> None:
        self.kpi_latency.setText(f"⏱️ Latency: {total_seconds:.1f}s")
        self.kpi_prompt_tokens.setText(f"📥 Prompt: {prompt_tokens} tok")
        self.kpi_eval_tokens.setText(f"📤 Eval: {eval_tokens} tok")
        speed = (eval_tokens / eval_seconds) if eval_seconds > 0 else 0.0
        self.kpi_speed.setText(f"⚡ Speed: {speed:.1f} tok/s")

    def _on_copy_trace(self) -> None:
        text = self.browser.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            orig = self.copy_btn.text()
            self.copy_btn.setText("✓ Copied!")
            QTimer.singleShot(1500, lambda: self.copy_btn.setText(orig))

    def clear(self) -> None:
        self.browser.clear()
        self.kpi_latency.setText("⏱️ Latency: 0.0s")
        self.kpi_prompt_tokens.setText("📥 Prompt: 0 tok")
        self.kpi_eval_tokens.setText("📤 Eval: 0 tok")
        self.kpi_speed.setText("⚡ Speed: 0 tok/s")
