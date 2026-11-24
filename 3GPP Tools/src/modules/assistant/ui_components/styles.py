"""QSS stylesheet presets, starter prompts, and activity tool icons for 3GPP Assistant."""

from __future__ import annotations

from core.ui.theme import C

PROMPT_INPUT_STYLE = f"""
QPlainTextEdit {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: 6px;
    padding: 6px 8px;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
    color: {C.TEXT_STRONG};
}}
QPlainTextEdit:focus {{
    border: 1.5px solid {C.ACCENT_PRIMARY};
}}
"""

STATUS_PILL_ONLINE = f"""
QPushButton {{
    background-color: {C.WG_RAN_BG};
    color: {C.EMERALD_DARK};
    border: 1px solid {C.WG_RAN_BORDER};
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}}
QPushButton:hover {{
    background-color: {C.EMERALD_SOFT};
    border-color: {C.EMERALD_BORDER};
}}
"""

STATUS_PILL_OFFLINE = f"""
QPushButton {{
    background-color: {C.DANGER_SOFT};
    color: {C.DANGER};
    border: 1px solid {C.DANGER_BORDER};
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    border-color: {C.DANGER_BORDER_ALT};
}}
"""

STATUS_PILL_NEUTRAL = f"""
QPushButton {{
    background-color: {C.BG_CANVAS};
    color: {C.TEXT_SECONDARY};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: 12px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
}}
QPushButton:hover {{
    background-color: {C.BORDER_SUBTLE};
}}
"""

STARTER_CHIP_STYLE = f"""
QPushButton {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    text-align: left;
}}
QPushButton:hover {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    border-color: {C.ACCENT_PRIMARY};
}}
QPushButton:pressed {{
    background-color: {C.ACCENT_SOFT_HOVER};
}}
"""

CARD_FRAME_STYLE = f"""
QFrame#assistantCardFrame {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: 8px;
}}
"""

CARD_HEADER_STYLE = f"""
QFrame#assistantCardHeader {{
    background-color: {C.BG_SURFACE};
    border-bottom: 1px solid {C.BORDER_SUBTLE};
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
}}
"""

BADGE_SPEC_TEXT = (
    f"background-color: {C.WG_SA_BG}; color: {C.INDIGO_DARK}; "
    f"border: 1px solid {C.WG_SA_BORDER}; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"
)
BADGE_PROTOCOL = (
    f"background-color: {C.WG_RAN_BG}; color: {C.EMERALD_DARK}; "
    f"border: 1px solid {C.WG_RAN_BORDER}; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"
)
BADGE_CATALOGUE = (
    f"background-color: {C.BG_CANVAS}; color: {C.TEXT_SECONDARY}; "
    f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 6px; font-weight: bold; font-size: 11px;"
)

STARTER_PROMPTS = [
    ("🔍 TS 23.501 Registration", "What is the Registration procedure in TS 23.501?"),
    ("📨 NGAP PDU Session Setup", "What IEs are in the NGAP PDU Session Resource Setup Request?"),
    ("🔬 TS 29.244 (PFCP) Messages", "List all messages defined for TS 29.244 (PFCP)"),
    ("📋 TS 38.331 RRCReconfiguration", "What is the structure of the RRCReconfiguration message in TS 38.331?"),
]

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
