"""Stylesheets and natural sorting helpers for PreMeetingExportDialog."""

import re

from core.ui.theme import C

STYLE_TOGGLE_ACTIVE = f"""
    QPushButton {{
        background-color: {C.TEAL_DARK};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.TEAL_DARK};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }}
"""

STYLE_TOGGLE_INACTIVE = f"""
    QPushButton {{
        background-color: {C.BG_CARD};
        color: {C.TEXT_SECONDARY};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }}
    QPushButton:hover {{
        background-color: {C.TEAL_SURFACE};
        color: {C.TEAL_DARK};
    }}
"""

STYLE_COMPACT_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:pressed {{
        background-color: {C.BORDER_SUBTLE};
    }}
"""

STYLE_TREE = f"""
    QTreeWidget {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 4px;
        font-size: 11px;
        color: {C.TEXT_STRONG};
    }}
    QTreeWidget::item {{
        padding: 3px 2px;
        border-radius: 3px;
    }}
    QTreeWidget::item:hover {{
        background-color: {C.TEAL_SURFACE};
    }}
    QTreeWidget::item:selected {{
        background-color: {C.TEAL_SOFT};
        color: {C.TEAL_DARK};
    }}
"""


def natural_ai_sort_key(ai_str: str) -> list:
    """Natural numerical sort key for 3GPP Agenda Item numbers (e.g. 5.1 before 5.10)."""
    parts = []
    for chunk in re.split(r'(\d+)', str(ai_str).strip()):
        if chunk.isdigit():
            parts.append(int(chunk))
        else:
            parts.append(chunk.lower())
    return parts
