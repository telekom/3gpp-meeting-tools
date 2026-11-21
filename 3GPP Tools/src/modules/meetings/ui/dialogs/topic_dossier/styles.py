"""Working-group presets and shared button/card QSS for TopicDossierDialog."""

from core.ui.theme import C

ALL_COMMON_WGS = ["SA2", "RAN2", "RAN3", "CT1", "RAN1", "SA3", "CT3", "CT4", "SA1", "SA4", "SA5", "RAN4"]

STYLE_CARD = f"""
    QFrame#CardFrame {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 8px;
    }}
"""

STYLE_PRIMARY_BTN = f"""
    QPushButton {{
        background-color: {C.EMERALD_ACCENT};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.EMERALD_PRIMARY};
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 12px;
    }}
    QPushButton:hover {{
        background-color: {C.EMERALD_PRIMARY};
    }}
    QPushButton:disabled {{
        background-color: {C.TEXT_FAINT};
        border-color: {C.TEXT_FAINT};
    }}
"""

STYLE_SECONDARY_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:disabled {{
        color: {C.TEXT_FAINT};
        background-color: {C.BG_SURFACE};
        border-color: {C.BORDER_SUBTLE};
    }}
"""
