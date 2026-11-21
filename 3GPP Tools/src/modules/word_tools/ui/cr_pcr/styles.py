"""Shared button styles for CrPcrDialog."""

from core.ui.theme import C

BUTTON_STYLE_PRIMARY = f"""
    QPushButton {{
        font-weight: bold;
        background-color: {C.ACCENT_BRIGHT};
        color: {C.TEXT_ON_ACCENT};
        padding: 7px 18px;
        border-radius: 4px;
        border: 1px solid {C.ACCENT_BRIGHT_BORDER};
    }}
    QPushButton:hover {{
        background-color: {C.ACCENT_BRIGHT_BORDER};
    }}
    QPushButton:disabled {{
        background-color: {C.BORDER_STRONG};
        color: {C.TEXT_FAINT};
        border: 1px solid {C.BORDER_STRONG};
    }}
"""
