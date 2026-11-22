"""Dialog displaying all specifications linked to a 3GPP Work Item."""

from __future__ import annotations

import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY


class LinkedSpecsDialog(QDialog):
    """Dialog displaying all specifications linked to a given Work Item."""

    def __init__(self, wi_code: str, acronym: str, specs: list, parent=None):
        super().__init__(parent)
        title_str = (
            f"Linked Specifications for WI {acronym} ({wi_code})"
            if acronym
            else f"Linked Specifications for WI #{wi_code}"
        )
        self.setWindowTitle(title_str)
        self.setMinimumWidth(580)
        self.setMinimumHeight(380)
        self.setStyleSheet(f"""
            QFrame#specCard {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}
            QFrame#specCard:hover {{ border-color: {C.TEXT_FAINT}; background-color: {C.BLUE_TINT_SOFT}; }}
            QLabel {{ color: {C.TEXT_STRONG}; }}
            QScrollArea {{ border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; background-color: {C.BG_CARD}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header_text = (
            f"Specifications Impacted by Work Item <b>{acronym}</b> ({wi_code}):"
            if acronym
            else f"Specifications Impacted by Work Item <b>#{wi_code}</b>:"
        )
        header = QLabel(header_text)
        header.setStyleSheet(f"font-size: 14px; color: {C.TEXT_STRONG};")
        header.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        cards_layout = QVBoxLayout(scroll_content)
        cards_layout.setContentsMargins(8, 8, 8, 8)
        cards_layout.setSpacing(8)

        if specs:
            for sp in specs:
                num = sp.get("number", "Unknown")
                title = sp.get("title", "No Title Available")
                spec_type = sp.get("type", "TS")
                series_name = sp.get("series_name", num.split(".")[0] if "." in num else "")
                is_pri = bool(sp.get("is_primary", False))

                card = self._create_spec_card(num, title, spec_type, series_name, is_pri)
                cards_layout.addWidget(card)
        else:
            no_data = QLabel("No linked specifications recorded in the local database for this Work Item.")
            no_data.setStyleSheet(f"color: {C.TEXT_MUTED}; font-style: italic; padding: 25px;")
            no_data.setAlignment(Qt.AlignCenter)
            cards_layout.addWidget(no_data)

        cards_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        footer_layout = QHBoxLayout()
        count_lbl = QLabel(f"<b>Total:</b> {len(specs)} specification(s)")
        count_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 12px;")
        footer_layout.addWidget(count_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        layout.addLayout(footer_layout)

    def _create_spec_card(
        self, num: str, title: str, spec_type: str, series_name: str, is_primary: bool
    ) -> QFrame:
        card = QFrame()
        card.setObjectName("specCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        type_badge = QLabel(f"<b>{spec_type}</b>")
        type_badge.setStyleSheet(f"""
            background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; border: 1px solid {C.INFO_BORDER};
            border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
        """)
        top_row.addWidget(type_badge)

        num_lbl = QLabel(f"<b>{num}</b>")
        num_lbl.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_STRONG};")
        num_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top_row.addWidget(num_lbl)

        if is_primary:
            pri_badge = QLabel("⭐ Primary Spec")
            pri_badge.setStyleSheet(f"""
                background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER};
                border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
            """)
            top_row.addWidget(pri_badge)

        top_row.addStretch()

        clean_num = num.split("-")[0].replace(".", "").strip()
        dyna_url = f"https://www.3gpp.org/DynaReport/{clean_num}.htm" if clean_num else ""
        ftp_url = f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/{num}/" if series_name else ""

        if dyna_url:
            dyna_btn = QPushButton("🌐 DynaReport")
            dyna_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            dyna_btn.setCursor(Qt.PointingHandCursor)
            dyna_btn.clicked.connect(lambda: webbrowser.open(dyna_url))
            top_row.addWidget(dyna_btn)

        if ftp_url:
            ftp_btn = QPushButton("📂 FTP Archive")
            ftp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(ftp_url))
            top_row.addWidget(ftp_btn)

        card_layout.addLayout(top_row)

        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_SECONDARY}; line-height: 1.4;")
        title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        card_layout.addWidget(title_lbl)

        return card
