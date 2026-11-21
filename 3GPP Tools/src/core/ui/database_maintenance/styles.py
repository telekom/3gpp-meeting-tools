"""QSS styles for the SQLite Database Maintenance & Compaction Hub."""

from __future__ import annotations

from core.ui.theme import C

DIALOG_BG_STYLE = f"background-color: {C.BG_ALT_LIGHT};"

KPI_CARD_STYLE = f"""
    QFrame#kpiCard {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 8px;
    }}
    QLabel {{
        border: none;
        background: transparent;
    }}
"""

KPI_TITLE_STYLE = (
    f"font-size: 10px; font-weight: bold; color: {C.TEXT_MUTED}; letter-spacing: 0.5px;"
)
KPI_SUB_STYLE = f"font-size: 10px; color: {C.TEXT_FAINT};"


def build_kpi_value_style(color: str = C.TEXT_PRIMARY) -> str:
    """Return CSS for a KPI card's main numeric value label."""
    return f"font-size: 16px; font-weight: bold; color: {color};"


HEADER_TITLE_STYLE = f"font-size: 15px; font-weight: bold; color: {C.TEXT_PRIMARY};"
HEADER_DESC_STYLE = f"color: {C.TEXT_MUTED}; font-size: 11px;"

TOP_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 5px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: bold;
        color: {C.TEXT_BODY};
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.BORDER_HOVER};
    }}
"""

STORAGE_TABLE_STYLE = f"""
    QTableWidget {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 6px;
        gridline-color: {C.BG_SURFACE};
        outline: none;
    }}
    QTableWidget::item {{
        border-bottom: 1px solid {C.BORDER_ROW};
        padding: 2px 4px;
    }}
    QTableWidget::item:selected {{
        background-color: {C.BG_SURFACE};
    }}
    QHeaderView::section {{
        background-color: {C.BG_SURFACE};
        padding: 8px 6px;
        font-weight: bold;
        color: {C.TEXT_SECONDARY};
        border: none;
        border-bottom: 2px solid {C.BORDER_SUBTLE};
        font-size: 11px;
    }}
"""

VACUUM_ALL_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.ACCENT_PRIMARY};
        color: {C.TEXT_ON_ACCENT};
        font-weight: bold;
        font-size: 12px;
        border-radius: 5px;
        padding: 4px 16px;
    }}
    QPushButton:hover {{
        background-color: {C.ACCENT_HOVER};
    }}
    QPushButton:disabled {{
        background-color: {C.TEXT_FAINT};
    }}
"""

PROGRESS_BAR_STYLE = f"""
    QProgressBar {{
        background-color: {C.BORDER_SUBTLE};
        border: none;
        border-radius: 4px;
    }}
    QProgressBar::chunk {{
        background-color: {C.ACCENT_SKY};
        border-radius: 4px;
    }}
"""

STATUS_LABEL_STYLE = f"font-size: 11px; color: {C.ACCENT_SKY}; font-weight: bold;"

CLOSE_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 5px;
        padding: 4px 16px;
        font-size: 11px;
        font-weight: bold;
        color: {C.TEXT_SECONDARY};
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.BORDER_HOVER};
    }}
"""

CELL_TITLE_STYLE = f"font-size: 12px; font-weight: bold; color: {C.TEXT_PRIMARY};"
CELL_SUB_STYLE = f"font-size: 10px; color: {C.TEXT_MUTED};"

INSPECT_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        font-size: 11px;
        font-weight: bold;
        color: {C.ACCENT_SKY_DARK};
        padding: 2px 10px;
        min-width: 66px;
    }}
    QPushButton:hover {{
        background-color: {C.SKY_SOFT};
        border-color: {C.ACCENT_SKY};
    }}
"""

COMPACT_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        font-size: 11px;
        font-weight: bold;
        color: {C.TEXT_PRIMARY};
        padding: 2px 10px;
        min-width: 70px;
    }}
    QPushButton:hover {{
        background-color: {C.BG_MUTED};
        border-color: {C.BORDER_HOVER};
    }}
"""

WIPE_BTN_STYLE = f"""
    QPushButton {{
        background-color: {C.DANGER_SOFT};
        border: 1px solid {C.DANGER_BORDER};
        border-radius: 4px;
        font-size: 11px;
        font-weight: bold;
        color: {C.DANGER_ALT};
        padding: 2px 10px;
        min-width: 58px;
    }}
    QPushButton:hover {{
        background-color: {C.DANGER_SOFT_HOVER};
        border-color: {C.DANGER_ACCENT};
    }}
"""
