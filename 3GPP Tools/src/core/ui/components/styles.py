"""Shared QSS stylesheet constants for cards, tables, trees, inputs, badges, and buttons."""

from core.ui.theme import C, R, T


CARD_FRAME_STYLE = f"""
QFrame#cardFrame {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.LG}px;
}}
"""

TABLE_STYLE_CLEAN = f"""
QTableWidget {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.MD}px;
    gridline-color: {C.BORDER_ROW};
    font-size: {T.SIZE_TABLE}px;
}}
QTableWidget::item {{
    padding: 4px 8px;
    border-bottom: 1px solid {C.BORDER_ROW};
}}
QTableWidget::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.TEXT_STRONG};
}}
QHeaderView::section {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_SECONDARY};
    font-weight: bold;
    font-size: {T.SIZE_CAPTION}px;
    border: none;
    border-bottom: 1px solid {C.BORDER_STRONG};
    padding: 6px 8px;
}}
"""

TREE_STYLE_CLEAN = f"""
QTreeWidget {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.MD}px;
    alternate-background-color: {C.BG_SURFACE};
    font-size: {T.SIZE_TABLE}px;
    outline: none;
}}
QTreeWidget::item {{
    min-height: 28px;
    padding: 2px 6px;
    border-bottom: 1px solid {C.BORDER_ROW};
}}
QTreeWidget::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.TEXT_STRONG};
}}
QTreeWidget::item:hover {{
    background-color: {C.BG_SURFACE};
}}
QHeaderView::section {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_SECONDARY};
    font-weight: bold;
    font-size: {T.SIZE_CAPTION}px;
    border: none;
    border-bottom: 1px solid {C.BORDER_STRONG};
    padding: 6px 8px;
}}
"""

SEARCH_INPUT_STYLE = f"""
QLineEdit {{
    padding: 4px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
}}
QLineEdit:focus {{
    border: 1px solid {C.ACCENT_PRIMARY};
}}
"""

BADGE_STYLE_PRIMARY = f"""
background-color: {C.SUCCESS_SOFT_MINT};
color: {C.SUCCESS_ALT};
font-weight: bold;
font-size: {T.SIZE_CAPTION}px;
border: 1px solid {C.SUCCESS_BORDER_MINT};
border-radius: {R.SM}px;
padding: 2px 8px;
"""

BADGE_STYLE_MUTED = f"""
background-color: {C.BG_CANVAS};
color: {C.TEXT_SECONDARY};
font-size: {T.SIZE_CAPTION}px;
border: 1px solid {C.BORDER_SUBTLE};
border-radius: {R.SM}px;
padding: 2px 8px;
"""

BADGE_STYLE_INFO = f"""
background-color: {C.INFO_SOFT_ALT};
color: {C.INFO_ALT};
border: 1px solid {C.INFO_BORDER_ALT};
border-radius: {R.SM}px;
padding: 2px 8px;
font-size: {T.SIZE_CAPTION}px;
font-weight: bold;
"""

BUTTON_STYLE_TOOLBAR_SECONDARY = f"""
QPushButton {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_STRONG};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:pressed {{
    background-color: {C.BG_MUTED};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    border: 1.5px solid {C.ACCENT_PRIMARY};
    font-weight: 600;
}}
QPushButton:checked:hover {{
    background-color: {C.ACCENT_SOFT_HOVER};
    border-color: {C.ACCENT_HOVER};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

BUTTON_STYLE_TOOLBAR_DANGER = f"""
QPushButton {{
    background-color: {C.DANGER_SOFT};
    color: {C.DANGER_ALT};
    border: 1px solid {C.DANGER_BORDER};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    border-color: {C.DANGER_BORDER_HOVER};
    color: {C.DANGER};
}}
QPushButton:pressed {{
    background-color: {C.DANGER_BORDER};
    border-color: {C.DANGER_ACCENT};
}}
QPushButton:checked {{
    background-color: {C.DANGER_SOFT_HOVER};
    color: {C.DANGER_DARK};
    border: 1.5px solid {C.DANGER_ACCENT};
    font-weight: bold;
}}
QPushButton:checked:hover {{
    background-color: {C.DANGER_BORDER};
    border-color: {C.DANGER_ALT};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

BUTTON_STYLE_TOOLBAR_WARNING = f"""
QPushButton {{
    background-color: {C.WARNING_SOFT};
    color: {C.WARNING};
    border: 1px solid {C.WARNING_BORDER};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {C.WARNING_SOFT_ALT};
    border-color: {C.WARNING_ACCENT};
}}
QPushButton:pressed {{
    background-color: {C.WARNING_BORDER};
}}
QPushButton:checked {{
    background-color: {C.WARNING_SOFT_ALT};
    color: {C.WARNING_DARK};
    border: 1.5px solid {C.WARNING_ACCENT};
    font-weight: 600;
}}
QPushButton:checked:hover {{
    background-color: {C.WARNING_BORDER};
    border-color: {C.WARNING};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

COMBOBOX_STYLE_TOOLBAR = f"""
QComboBox {{
    combobox-popup: 0;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 600;
    color: {C.TEXT_STRONG};
    background-color: {C.BG_SURFACE};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    padding: 3px 22px 3px 10px;
    min-height: 22px;
}}
QComboBox:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QComboBox:focus, QComboBox:on {{
    background-color: {C.BG_CARD};
    border-color: {C.ACCENT_PRIMARY};
}}
QComboBox QAbstractItemView {{
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.MD}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
    selection-background-color: {C.ACCENT_SOFT};
    selection-color: {C.ACCENT_PRIMARY};
    padding: 4px;
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    min-height: 22px;
    padding: 2px 8px;
    border-radius: 3px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {C.BG_CANVAS};
    color: {C.TEXT_PRIMARY};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    font-weight: bold;
}}
QScrollBar:vertical {{
    border: none;
    background: {C.BG_SURFACE};
    width: 8px;
    margin: 4px 0;
    border-radius: {R.SM}px;
}}
QScrollBar::handle:vertical {{
    background: {C.BORDER_STRONG};
    min-height: 20px;
    border-radius: {R.SM}px;
}}
QScrollBar::handle:vertical:hover {{
    background: {C.BORDER_HOVER};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
"""

SEGMENT_BTN_STYLE = f"""
QPushButton {{
    padding: 3px 9px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_BRIGHT};
    color: {C.TEXT_ON_ACCENT};
    font-weight: bold;
    border: 1px solid {C.ACCENT_BRIGHT_BORDER};
}}
"""

TOGGLE_PILL_STYLE = f"""
QPushButton {{
    padding: 3px 12px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.PILL}px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    font-weight: bold;
    border: 1.5px solid {C.ACCENT_PRIMARY};
    border-radius: {R.PILL}px;
}}
QPushButton:checked:hover {{
    background-color: {C.ACCENT_SOFT_HOVER};
}}
"""

FAV_PILL_STYLE = f"""
QPushButton {{
    padding: 3px 10px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.PILL}px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.WARNING_SOFT_ALT};
    color: {C.WARNING_DARK};
    font-weight: bold;
    border: 1.5px solid {C.WARNING_ACCENT};
}}
QPushButton:checked:hover {{
    background-color: {C.WARNING_BORDER};
}}
"""

DISMISSIBLE_TAG_STYLE = f"""
QPushButton {{
    padding: 2px 8px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.INFO_SOFT};
    color: {C.INFO_DARK};
    border: 1px solid {C.INFO_BORDER};
    border-radius: {R.XL}px;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    color: {C.DANGER_DARK};
    border-color: {C.DANGER_BORDER_ALT};
}}
"""

CLEAR_ALL_BTN_STYLE = f"""
QPushButton {{
    padding: 2px 8px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: bold;
    background-color: {C.DANGER_SOFT};
    color: {C.DANGER};
    border: 1px solid {C.DANGER_BORDER};
    border-radius: {R.SM}px;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    border-color: {C.DANGER_BORDER_HOVER};
}}
"""

SPEC_FILTER_COMBO_STYLE = f"""
QComboBox {{
    combobox-popup: 0;
    padding: 3px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
}}
QComboBox:hover {{
    border-color: {C.BORDER_HOVER};
}}
QComboBox:focus {{
    border-color: {C.ACCENT_PRIMARY};
}}
"""

SPEC_FILTER_INPUT_STYLE = f"""
QLineEdit {{
    padding: 4px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
}}
QLineEdit:focus {{
    border-color: {C.ACCENT_PRIMARY};
}}
"""
