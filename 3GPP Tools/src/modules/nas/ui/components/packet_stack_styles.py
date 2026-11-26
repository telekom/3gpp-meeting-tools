"""Style helpers and layout utilities for PacketStackWidget."""

from __future__ import annotations

import html
from typing import Callable

from PyQt5.QtWidgets import QAction, QMenu, QWidget

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.nas.core.packet_stack import NGAP_SCENARIO_PRESETS, UP_SCENARIO_PRESETS


def flag_pill_qss(*, release_present: bool = True, is_ext: bool = False) -> str:
    """Return theme-tokenized QSS for checkable packet flag and extension-header pills."""
    fg = C.TEXT_BODY if release_present else C.TEXT_FAINT
    bg = C.BG_CARD if release_present else C.BG_SURFACE
    if is_ext:
        chk_fg, chk_bg, chk_bd, chk_hv = C.SUCCESS_DARK, C.SUCCESS_SOFT_ALT, C.SUCCESS, C.SUCCESS_BORDER
    else:
        chk_fg, chk_bg, chk_bd, chk_hv = C.ACCENT_SKY_DARK, C.SKY_SOFT, C.ACCENT_SKY, C.INFO_SOFT_STRONG
    return (
        f"QPushButton {{ font-size: 10px; font-weight: 600; color: {fg}; "
        f"background-color: {bg}; border: 1px solid {C.BORDER_STRONG}; "
        f"border-radius: 10px; padding: 2px 8px; }} "
        f"QPushButton:hover {{ border-color: {chk_bd}; background-color: {C.BG_HOVER}; }} "
        f"QPushButton:checked {{ color: {chk_fg}; background-color: {chk_bg}; "
        f"border: 1.5px solid {chk_bd}; }} "
        f"QPushButton:checked:hover {{ background-color: {chk_hv}; }}"
    )


def quick_btn_qss() -> str:
    """Return subtle neutral QSS for compact action buttons (`Minimal`, `All Opt`, `✕`)."""
    return (
        f"QPushButton {{ font-size: 10px; font-weight: 600; color: {C.TEXT_SECONDARY}; "
        f"background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; "
        f"border-radius: 4px; padding: 2px 7px; }} "
        f"QPushButton:hover {{ background-color: {C.BG_HOVER}; color: {C.TEXT_PRIMARY}; }}"
    )


def clear_flow_layout(flow: FlowLayout) -> None:
    """Remove and schedule deletion of all widgets in a FlowLayout."""
    while flow.count():
        item = flow.takeAt(0)
        if item and item.widget():
            item.widget().deleteLater()


def build_scenarios_guide_html(stack_family: str = "up") -> str:
    """Build rich HTML summary of all scenarios for the active stack family."""
    presets = NGAP_SCENARIO_PRESETS if stack_family == "ngap" else UP_SCENARIO_PRESETS
    title = (
        "N2 NGAP Scenarios (TS 38.412, TS 38.413 &amp; TS 24.501):"
        if stack_family == "ngap"
        else "5G User Plane Scenarios (TS 29.281 &amp; TS 38.415):"
    )
    items = "".join(
        f"<div style='margin: 2px 0;'><b>{html.escape(p.label)}</b> "
        f"<span style='color: {C.TEXT_MUTED};'>[{html.escape(p.interface)}]</span>: "
        f"{html.escape(p.description)}</div>"
        for p in presets
    )
    return f"<div><b>{title}</b>{items}</div>"


def build_packet_copy_menu(
    parent: QWidget,
    on_copy: Callable[[str], None],
    on_open_ide: Callable[[], None],
) -> QMenu:
    """Construct the Copy Packet export dropdown menu."""
    menu = QMenu(parent)
    menu.setStyleSheet(
        f"QMenu {{ font-size: 11px; background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; padding: 4px; }} "
        f"QMenu::item {{ padding: 5px 18px 5px 10px; border-radius: 3px; }} "
        f"QMenu::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; }}"
    )
    for text, mode in (
        ("📋 Copy for Word / Outlook (Rich Diagram)", "word"),
        ("📐 Copy as PlantUML (packetdiag)", "packetdiag"),
        ("📝 Copy as Markdown Table", "markdown"),
        ("📄 Copy Plain Text", "text"),
    ):
        act = QAction(text, parent)
        act.triggered.connect(lambda _c=False, m=mode: on_copy(m))
        menu.addAction(act)
    menu.addSeparator()
    act_ide = QAction("✏️ Open in PlantUML IDE", parent)
    act_ide.triggered.connect(on_open_ide)
    menu.addAction(act_ide)
    return menu
