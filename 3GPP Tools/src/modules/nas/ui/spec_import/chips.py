"""Shared chip / badge widgets for the Specification Import Studio."""

from __future__ import annotations

from typing import Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QLabel, QPushButton, QWidget

from core.ui.theme import C
from core.ui.ui_components import escape_qt_mnemonic

_BADGE_COLORS = {
    "wg": (C.INFO_DARK, C.INFO_SOFT),
    "gen": (C.TEAL_DARK, C.TEAL_SOFT),
    "domain": (C.ACCENT_SKY_DARK, C.SKY_SOFT),
    "plane": (C.VIOLET_DARK, C.VIOLET_SOFT),
    "layer": (C.ORANGE_DARK, C.ORANGE_SOFT),
    "ok": (C.SUCCESS_DARK, C.SUCCESS_SOFT),
    "queued": (C.ACCENT_PRIMARY, C.ACCENT_SOFT),
    "muted": (C.TEXT_MUTED, C.BG_MUTED),
}


def badge_colors(kind: str) -> Tuple[str, str]:
    return _BADGE_COLORS.get(kind, _BADGE_COLORS["muted"])


def make_badge(text: str, kind: str = "muted", parent: Optional[QWidget] = None) -> QLabel:
    """Small rounded pill label (`kind` selects the theme-token colour pair)."""
    fg, bg = badge_colors(kind)
    lbl = QLabel(text, parent)
    lbl.setStyleSheet(
        f"background-color: {bg}; color: {fg}; border-radius: 7px; "
        f"padding: 1px 7px; font-size: 10px; font-weight: 600;"
    )
    lbl.setAlignment(Qt.AlignCenter)
    return lbl


def chip_qss() -> str:
    """QSS for checkable filter chips (defines `:checked` and `:checked:hover`)."""
    return (
        f"QPushButton {{ font-size: 11px; font-weight: 600; color: {C.TEXT_SECONDARY}; "
        f"background-color: {C.BG_CARD}; border: 1.5px solid {C.BORDER_STRONG}; "
        f"border-radius: 11px; padding: 2px 10px; min-height: 18px; }} "
        f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.ACCENT_SKY}; }} "
        f"QPushButton:checked {{ color: {C.ACCENT_PRIMARY}; background-color: {C.ACCENT_SOFT}; "
        f"border: 1.5px solid {C.ACCENT_PRIMARY}; font-weight: 600; }} "
        f"QPushButton:checked:hover {{ background-color: {C.SKY_SOFT}; }}"
    )


def make_chip(text: str, tooltip: str = "", *, checkable: bool = True, parent: Optional[QWidget] = None) -> QPushButton:
    btn = QPushButton(escape_qt_mnemonic(text), parent)
    btn.setCheckable(checkable)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setStyleSheet(chip_qss())
    if tooltip:
        btn.setToolTip(tooltip)
    return btn


def primary_button_qss() -> str:
    return (
        f"QPushButton {{ background-color: {C.ACCENT_PRIMARY}; color: {C.BG_CARD}; border: none; "
        f"border-radius: 5px; padding: 7px 16px; font-size: 12px; font-weight: 600; }} "
        f"QPushButton:hover {{ background-color: {C.ACCENT_SKY_DARK}; }} "
        f"QPushButton:disabled {{ background-color: {C.BG_MUTED}; color: {C.TEXT_FAINT}; }}"
    )
