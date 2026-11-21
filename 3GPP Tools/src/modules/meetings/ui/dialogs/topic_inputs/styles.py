"""Stylesheets and decision-type badge color palette for TopicInputsViewerDialog."""

from core.ui.theme import C

STYLE_PRIMARY_BTN = f"""
    QPushButton {{
        background-color: {C.EMERALD_ACCENT};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.EMERALD_PRIMARY};
        border-radius: 6px;
        padding: 6px 14px;
        font-size: 11px;
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
    }}
"""

STYLE_TABLE = f"""
    QTableWidget {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        gridline-color: {C.BG_CANVAS};
        font-size: 11px;
    }}
    QHeaderView::section {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_SECONDARY};
        font-weight: bold;
        font-size: 11px;
        border: none;
        border-bottom: 1px solid {C.BORDER_SUBTLE};
        padding: 6px;
    }}
    QTableWidget::item {{
        padding: 4px 6px;
    }}
    QTableWidget::item:selected {{
        background-color: {C.INFO_SOFT};
        color: {C.INFO};
    }}
"""

BADGE_COLORS = {
    "Text Agreement": (C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK),
    "Approved pCR": (C.WG_RAN_BG, C.WG_RAN_TEXT),
    "Approved CR": (C.EMERALD_SOFT, C.EMERALD_DARK),
    "Endorsed WF": (C.INFO_SOFT_STRONG, C.INFO_DARK),
    "Working Assumption": (C.WARNING_SOFT_ALT, C.WARNING_DARK),
    "Agreed LS": (C.VIOLET_SOFT, C.VIOLET_DARK),
    "LS in": (C.VIOLET_SOFT, C.PURPLE_DARK),
    "LS out": (C.INDIGO_SOFT, C.WG_SA_TEXT),
    "Discussion": (C.BG_CANVAS, C.TEXT_BODY),
    "Work Plan": (C.DANGER_SOFT_HOVER, C.DANGER_DARK),
}
