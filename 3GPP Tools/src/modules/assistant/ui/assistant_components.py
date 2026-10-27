"""Reusable UI components, prompt editor, and style presets for the 3GPP Assistant."""

from __future__ import annotations

from typing import Optional
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QKeyEvent
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_DANGER,
)

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
QLabel {
    background-color: #ECFDF5;
    color: #047857;
    border: 1px solid #A7F3D0;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}
"""

STATUS_PILL_OFFLINE = """
QLabel {
    background-color: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}
"""

STATUS_PILL_NEUTRAL = """
QLabel {
    background-color: #F1F5F9;
    color: #475569;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
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
                # Shift+Enter: newline
                super().keyPressEvent(event)
            else:
                # Plain Enter: submit
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
