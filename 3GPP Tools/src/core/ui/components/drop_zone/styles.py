"""Centralized QSS style definitions for ModernDropZone."""

from __future__ import annotations

from core.ui.theme import C

from .metadata import (
    COLOR_TEXT_MUTED,
    COLOR_TEXT_PRIMARY,
    COLOR_TEXT_SECONDARY,
)

EMPTY_ICON_COMPACT_STYLE = (
    f"font-size: 20px; background-color: {C.BG_HOVER}; border-radius: 18px; padding: 4px;"
)
EMPTY_ICON_CANVAS_STYLE = (
    f"font-size: 20px; background-color: {C.BG_HOVER}; border-radius: 18px; padding: 2px;"
)

TITLE_LABEL_STYLE = f"font-size: 13px; font-weight: 600; color: {COLOR_TEXT_PRIMARY};"
SUBTITLE_COMPACT_STYLE = f"font-size: 11px; color: {COLOR_TEXT_MUTED};"
SUBTITLE_CANVAS_STYLE = f"font-size: 12px; color: {COLOR_TEXT_SECONDARY};"

DROP_ZONE_ACTION_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 5px;
        padding: 5px 12px;
        font-size: 12px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.INFO_SOFT};
        border-color: {C.INFO_ACCENT};
        color: {C.INFO};
    }}
"""

EXT_BADGE_STYLE = f"""
    QLabel {{
        background-color: {C.BG_CANVAS};
        color: {C.TEXT_BODY};
        font-size: 11px;
        font-weight: 600;
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 2px 7px;
    }}
"""

FILE_SIZE_BADGE_STYLE = f"""
    QLabel {{
        background-color: {C.BG_CANVAS};
        color: {C.TEXT_BODY};
        font-size: 10px;
        font-weight: 600;
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 4px;
        padding: 1px 6px;
    }}
"""

SENSITIVITY_BADGE_STYLE = f"""
    QLabel {{
        background-color: {C.WG_RAN_BG};
        color: {C.WG_RAN_TEXT};
        font-size: 10px;
        font-weight: 600;
        border: 1px solid {C.WG_RAN_BORDER};
        border-radius: 4px;
        padding: 1px 6px;
    }}
"""

CLEAR_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_MUTED};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 14px;
        font-size: 12px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {C.DANGER_SOFT_HOVER};
        border-color: {C.DANGER_ACCENT};
        color: {C.DANGER};
    }}
"""


def build_file_icon_badge_style(bg: str, fg: str, border: str) -> str:
    """Build QSS styling for the loaded-file archetype badge."""
    return f"""
        QLabel {{
            background-color: {bg};
            color: {fg};
            border: 1px solid {border};
            border-radius: 8px;
            font-size: 20px;
        }}
    """


MULTI_FILE_ICON_BADGE_STYLE = build_file_icon_badge_style(
    C.INFO_SOFT,
    C.INFO,
    C.INFO_BORDER,
)
