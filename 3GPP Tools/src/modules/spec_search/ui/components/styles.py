"""Inspector button stylesheets using centralized ThemeColors tokens."""

from core.ui.theme import C


INSPECTOR_BTN_BASE = f"""
QPushButton {{
    font-size: 11px;
    font-weight: 500;
    color: {C.TEXT_BODY};
    background-color: {C.BG_SURFACE};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    color: {C.TEXT_PRIMARY};
    border-color: {C.TEXT_FAINT};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

INSPECTOR_BTN_SUCCESS = f"""
QPushButton {{
    font-size: 11px;
    font-weight: 600;
    color: {C.SUCCESS_DARK};
    background-color: {C.SUCCESS_SOFT};
    border: 1px solid {C.SUCCESS_BORDER_MEDIUM};
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.SUCCESS_SOFT_ALT};
    color: {C.SUCCESS_DARK};
    border-color: {C.EMERALD_ACCENT};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

INSPECTOR_BTN_AI = f"""
QPushButton {{
    font-size: 11px;
    font-weight: 600;
    color: {C.VIOLET_PRIMARY};
    background-color: {C.VIOLET_SURFACE};
    border: 1px solid {C.VIOLET_BORDER};
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.VIOLET_SOFT};
    color: {C.VIOLET_DARK};
    border-color: {C.VIOLET_BORDER_STRONG};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

INSPECTOR_BTN_WARNING = f"""
QPushButton {{
    font-size: 11px;
    font-weight: 500;
    color: {C.WARNING_DARK};
    background-color: {C.WARNING_SOFT};
    border: 1px solid {C.WARNING_BORDER};
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.WARNING_SOFT_ALT};
    color: {C.WARNING_BROWN_DEEP};
    border-color: {C.WARNING_ACCENT};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

INSPECTOR_BTN_INFO = f"""
QPushButton {{
    font-size: 11px;
    font-weight: 500;
    color: {C.INFO};
    background-color: {C.INFO_SOFT};
    border: 1px solid {C.INFO_BORDER};
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.INFO_SOFT_STRONG};
    color: {C.INFO_DARK};
    border-color: {C.INFO_BORDER_MEDIUM};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""
