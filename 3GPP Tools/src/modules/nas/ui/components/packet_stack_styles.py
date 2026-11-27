"""Style helpers and layout utilities for PacketStackWidget."""

from __future__ import annotations

import html
from typing import Callable

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QAction, QMenu, QWidget

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.nas.core.packet_stack import (
    NAS_SCENARIO_PRESETS,
    NGAP_SCENARIO_PRESETS,
    SAMPLE_5GSM_MESSAGES,
    SAMPLE_NAS_MESSAGES,
    UE_POLICY_SPECS,
    UP_SCENARIO_PRESETS,
    get_presets_for_family,
)


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
    if stack_family == "ngap":
        presets = NGAP_SCENARIO_PRESETS
        title = "N2 NGAP Scenarios (TS 38.412, TS 38.413 &amp; TS 24.501):"
    elif stack_family == "nas":
        presets = NAS_SCENARIO_PRESETS
        title = "5GS NAS &amp; Encapsulated Payload Scenarios (TS 24.501 &amp; UE Policy Specs):"
    else:
        presets = UP_SCENARIO_PRESETS
        title = "5G User Plane Scenarios (TS 29.281 &amp; TS 38.415):"
    items = "".join(
        f"<div style='margin: 2px 0;'><b>{html.escape(p.label)}</b> "
        f"<span style='color: {C.TEXT_MUTED};'>[{html.escape(p.interface)}]</span>: "
        f"{html.escape(p.description)}</div>"
        for p in presets
    )
    return f"<div><b>{title}</b>{items}</div>"


def populate_family_dropdowns(widget: QWidget, stack_family: str) -> None:
    """Populate scenario presets, carrier/transport options, and family-specific combos."""
    widget.scenario_combo.clear()
    for idx, p in enumerate(get_presets_for_family(stack_family)):
        widget.scenario_combo.addItem(p.label, p.preset_id)
        widget.scenario_combo.setItemData(idx, p.description, Qt.ToolTipRole)

    widget.transport_combo.clear()
    widget.nas_msg_combo.clear()
    if stack_family == "ngap":
        for label, key in (
            ("Outer: Collapsed", "collapsed"),
            ("Outer: IPv4 + SCTP (48B)", "ipv4_sctp"),
            ("Outer: IPv6 + SCTP (68B)", "ipv6_sctp"),
            ("Outer: SCTP Only (28B)", "sctp_only"),
            ("Outer: NGAP APER Only", "none"),
        ):
            widget.transport_combo.addItem(label, key)
        for nm in SAMPLE_NAS_MESSAGES:
            widget.nas_msg_combo.addItem(f"NAS: {nm}", nm)
        widget.lbl_ext_row.setText("SCTP & NAS:")
        widget.lbl_frame_row.setText("Optional IEs:")
        widget.tpdu_btn.setVisible(False)
    elif stack_family == "nas":
        for label, key in (
            ("Carrier: N1 Standalone (NAS Only)", "n1_standalone"),
            ("Carrier: Over N2 (NGAP + SCTP)", "over_n2_ngap"),
            ("Carrier: Over Uu (RRC + PDCP + RLC + MAC)", "over_uu_rrc"),
        ):
            widget.transport_combo.addItem(label, key)
        for nm in SAMPLE_5GSM_MESSAGES:
            widget.nas_msg_combo.addItem(f"5GSM: {nm}", nm)
        widget.policy_part_combo.clear()
        for pkey, pinfo in UE_POLICY_SPECS.items():
            widget.policy_part_combo.addItem(f"Policy: {pinfo['label']}", pkey)
        widget.lbl_ext_row.setText("NAS Envelope:")
        widget.lbl_frame_row.setText("Payload / IEs:")
        widget.tpdu_btn.setVisible(False)
    else:
        for label, key in (
            ("Outer: Collapsed", "collapsed"),
            ("Outer: IPv4 + UDP (28B)", "ipv4_udp"),
            ("Outer: IPv6 + UDP (48B)", "ipv6_udp"),
            ("Outer: Hidden (3GPP Only)", "none"),
        ):
            widget.transport_combo.addItem(label, key)
        widget.lbl_ext_row.setText("Ext & GTP-U:")
        widget.lbl_frame_row.setText("Frame Flags:")
        widget.tpdu_btn.setVisible(True)
    widget.scenarios_guide_lbl.setText(build_scenarios_guide_html(stack_family))


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
