"""Regex patterns and toolbar/button QSS constants for NASTab."""

from __future__ import annotations

import re

from core.ui.theme import C

RE_SPEC_FROM_COL = re.compile(r"(?:24|25|29|36|37|38)\.[0-9]{3}")
RE_CLAUSE_FROM_REF = re.compile(r"((?:9|8|7|6|5|4|D\.6)(?:\.[0-9A-Za-z]+)+)")
RE_CTX_SRB = re.compile(
    r"\bSignalling\s+radio\s+bearer\s*:\s*(.+?)(?=\s+(?:RLC-SAP|Logical\s+channels?|Direction)\s*:|$)",
    re.IGNORECASE,
)
RE_CTX_RLC = re.compile(
    r"\bRLC-SAP\s*:\s*(.+?)(?=\s+(?:Logical\s+channels?|Direction)\s*:|$)",
    re.IGNORECASE,
)
RE_CTX_LCH = re.compile(
    r"\bLogical\s+channels?\s*:\s*(.+?)(?=\s+Direction\s*:|$)",
    re.IGNORECASE,
)
RE_CTX_TRAILING_CAPTION = re.compile(
    r"(?:\s+|^)(?:[A-Z][A-Za-z0-9]*(?:-[A-Za-z0-9]+|[A-Z0-9][A-Za-z0-9]*)[A-Za-z0-9\-]*)\s+"
    r"(?:information\s+elements?|messages?)\s*$"
)

SEGMENTED_MODE_STYLE = f"""
    QPushButton {{
        font-size: 11px;
        font-weight: 500;
        color: {C.TEXT_SECONDARY};
        background-color: transparent;
        border: 1px solid transparent;
        border-radius: 4px;
        padding: 2px 8px;
    }}
    QPushButton:hover {{
        color: {C.TEXT_PRIMARY};
        background-color: {C.BG_SURFACE};
    }}
    QPushButton:checked {{
        font-weight: 600;
        color: {C.ACCENT_PRIMARY};
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
    }}
"""

TOGGLE_BUTTON_STYLE = f"""
    QPushButton {{
        font-size: 11px;
        font-weight: bold;
        color: {C.TEXT_MUTED};
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 3px 8px;
    }}
    QPushButton:hover {{
        color: {C.ACCENT_PRIMARY};
        background-color: {C.ACCENT_SOFT};
        border-color: {C.ACCENT_PRIMARY};
    }}
    QPushButton:checked {{
        color: {C.ACCENT_PRIMARY};
        background-color: {C.ACCENT_SOFT};
        border: 1.5px solid {C.ACCENT_PRIMARY};
    }}
    QPushButton:checked:hover {{
        color: {C.ACCENT_PRIMARY};
        background-color: {C.ACCENT_SOFT_PRESSED};
        border: 1.5px solid {C.ACCENT_PRIMARY};
    }}
"""

TINY_HEADER_BTN_STYLE = f"""
    QPushButton {{
        font-size: 10px;
        font-weight: 500;
        color: {C.TEXT_SECONDARY};
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 3px;
        padding: 1px 7px;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        color: {C.TEXT_PRIMARY};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:pressed {{
        background-color: {C.BORDER_SUBTLE};
    }}
"""

RIBBON_BADGE_STYLE = f"""
    background-color: {C.BG_CANVAS};
    color: {C.TEXT_BODY};
    font-size: 11px;
    font-weight: 600;
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: 11px;
    padding: 3px 10px;
"""
