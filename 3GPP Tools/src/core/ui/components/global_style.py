"""Application-wide Qt global stylesheet (All-Blue Theme)."""

from core.ui.theme import C, R, T


def _combo_chevron_url() -> str:
    """Return cached SVG URL for the global QComboBox dropdown chevron."""
    try:
        from core.ui.icons import lucide_svg_path
        return lucide_svg_path("chevron-down", C.TEXT_MUTED)
    except Exception:
        return ""


_GLOBAL_COMBO_CHEVRON = _combo_chevron_url()
_GLOBAL_COMBO_ARROW_QSS = (
    f'image: url("{_GLOBAL_COMBO_CHEVRON}"); width: 12px; height: 12px;'
    if _GLOBAL_COMBO_CHEVRON
    else ""
)

GLOBAL_STYLE = f"""
    QWidget {{
        font-family: {T.FONT_FAMILY_UI};
        font-size: {T.SIZE_BODY}px;
        color: {C.TEXT_BODY};
    }}
    QToolTip {{
        color: {C.TEXT_BODY};
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.SM}px;
        padding: 4px;
    }}
    QTabWidget::pane {{
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.LG}px;
        background: {C.BG_CARD};
        top: -1px;
    }}
    QTabBar::tab {{
        background: {C.BG_CANVAS};
        border: 1px solid {C.BORDER_STRONG};
        padding: 8px 16px;
        margin-right: 2px;
        border-top-left-radius: {R.MD}px;
        border-top-right-radius: {R.MD}px;
    }}
    QTabBar::tab:selected {{
        background: {C.BG_CARD};
        border-bottom-color: {C.BG_CARD};
        font-weight: bold;
        color: {C.ACCENT_PRIMARY};
    }}
    QTabBar::tab:hover:!selected {{
        background: {C.BG_SURFACE};
    }}
    QPushButton {{
        padding: 8px 16px;
        border-radius: {R.MD}px;
        border: 1px solid {C.BORDER_STRONG};
        background-color: {C.BG_SURFACE};
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
    }}
    QPushButton:disabled {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_FAINT};
        border: 1px solid {C.BORDER_SUBTLE};
    }}
    QPushButton:checked {{
        background-color: {C.ACCENT_SOFT};
        border: 2px solid {C.ACCENT_PRIMARY};
        color: {C.ACCENT_PRIMARY};
    }}
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {{
        background-color: {C.ACCENT_PRIMARY};
        color: {C.TEXT_ON_ACCENT};
        border: none;
    }}
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {{
        background-color: {C.ACCENT_HOVER};
    }}
    QPushButton#primaryBtn:disabled, QPushButton#pptBtn:disabled, QPushButton#svgBtn:disabled {{
        background-color: {C.BG_MUTED};
        color: {C.TEXT_FAINT};
        border: 1px solid {C.BORDER_STRONG};
    }}
    QMenu {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.MD}px;
        padding: 4px;
    }}
    QMenu::item {{
        padding: 8px 24px 8px 12px;
        border-radius: {R.SM}px;
        font-size: {T.SIZE_BODY}px;
        color: {C.TEXT_BODY};
    }}
    QMenu::item:selected {{
        background-color: {C.ACCENT_SOFT};
        color: {C.ACCENT_PRIMARY};
        font-weight: bold;
    }}
    QPushButton::menu-indicator {{
        width: 0px;
    }}
    QSplitter::handle {{
        background-color: {C.BORDER_SUBTLE};
        height: 2px;
        margin: 4px 0px;
    }}
    QSplitter::handle:hover {{
        background-color: {C.ACCENT_PRIMARY};
    }}
    QStatusBar {{
        background-color: {C.BG_SURFACE};
        border-top: 1px solid {C.BORDER_STRONG};
        color: {C.TEXT_BODY};
    }}
    QComboBox {{
        combobox-popup: 0;
        padding: 4px 26px 4px 10px;
        border-radius: {R.SM}px;
        border: 1px solid {C.BORDER_STRONG};
        background-color: {C.BG_CARD};
        font-weight: bold;
        color: {C.TEXT_BODY};
    }}
    QComboBox:hover {{
        border-color: {C.BORDER_HOVER};
    }}
    QComboBox:focus, QComboBox:on {{
        border-color: {C.ACCENT_PRIMARY};
    }}
    QComboBox:disabled {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_MUTED};
        border-color: {C.BORDER_SUBTLE};
    }}
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 22px;
        border: none;
    }}
    QComboBox::down-arrow {{
        {_GLOBAL_COMBO_ARROW_QSS}
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
        min-height: 24px;
        padding: 3px 8px;
        border-radius: 3px;
    }}
    QTextEdit#console, QListWidget#queueList {{
        background-color: {C.CONSOLE_LEGACY_BG};
        color: {C.CONSOLE_TEXT};
        font-family: {T.FONT_FAMILY_MONO};
        font-size: {T.SIZE_BODY}px;
        border-radius: {R.LG}px;
        padding: 8px;
        border: 1px solid {C.CONSOLE_BORDER};
    }}
    QListWidget#queueList::item {{
        padding: 4px;
        border-bottom: 1px solid {C.CONSOLE_ROW_DIVIDER};
    }}
    QListWidget#queueList::item:selected {{
        background-color: {C.CONSOLE_SEL_BG};
        color: {C.TEXT_ON_ACCENT};
        border-radius: {R.SM}px;
    }}
"""
