"""Centralized design-token QSS styling for QuickSpecSearchPanel."""

from __future__ import annotations

from PyQt5.QtWidgets import QWidget

from core.ui.theme import C, R, T


def apply_quick_search_styles(widget: QWidget) -> None:
    """Apply token-based QSS stylesheet to QuickSpecSearchPanel."""
    widget.setStyleSheet(f"""
        QWidget#quickSpecSearchPanel {{
            background-color: {C.BG_CANVAS};
        }}
        QFrame#quickSearchFilterCard, QFrame#quickSearchPreviewFrame {{
            background-color: {C.BG_CARD};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: {R.LG}px;
        }}
        QLineEdit#quickSearchInput, QLineEdit#quickSearchSubInput {{
            background-color: {C.BG_SURFACE};
            color: {C.TEXT_PRIMARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: {R.MD}px;
            padding: 5px 8px;
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_BODY}px;
        }}
        QLineEdit#quickSearchInput:focus, QLineEdit#quickSearchSubInput:focus {{
            border-color: {C.ACCENT_PRIMARY};
            background-color: {C.BG_CARD};
        }}
        QPushButton#quickSearchTogglePill {{
            background-color: {C.BG_SURFACE};
            color: {C.TEXT_SECONDARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: {R.MD}px;
            padding: 4px 8px;
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_CAPTION}px;
            font-weight: 600;
        }}
        QPushButton#quickSearchTogglePill:checked {{
            background-color: {C.ACCENT_SOFT};
            color: {C.ACCENT_PRIMARY};
            border-color: {C.ACCENT_PRIMARY};
        }}
        QLabel#quickSearchStatusLabel {{
            color: {C.TEXT_MUTED};
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_CAPTION}px;
            padding: 0 2px;
        }}
        QListWidget#quickSearchResultsList {{
            background-color: transparent;
            border: none;
            outline: none;
        }}
        QLabel#quickSearchPreviewTitle {{
            color: {C.TEXT_STRONG};
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_TABLE}px;
            font-weight: 600;
        }}
        QTextBrowser#quickSearchPreviewBrowser {{
            background-color: {C.BG_SURFACE};
            color: {C.TEXT_PRIMARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: {R.MD}px;
            padding: 6px;
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_TABLE}px;
        }}
        QPushButton#quickSearchActionBtn {{
            background-color: {C.BG_SURFACE};
            color: {C.TEXT_PRIMARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: {R.MD}px;
            padding: 4px 8px;
            font-family: {T.FONT_FAMILY_UI};
            font-size: {T.SIZE_CAPTION}px;
            font-weight: 600;
        }}
        QPushButton#quickSearchActionBtn:hover {{
            background-color: {C.ACCENT_SOFT};
            border-color: {C.ACCENT_PRIMARY};
            color: {C.ACCENT_PRIMARY};
        }}
        QPushButton#quickSearchIconBtn {{
            background-color: transparent;
            border: none;
            border-radius: {R.SM}px;
            padding: 2px;
        }}
        QPushButton#quickSearchIconBtn:hover {{
            background-color: {C.BG_SURFACE};
        }}
    """)
