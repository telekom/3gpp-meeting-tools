"""Interactive 5G User Plane Full Packet Stack widget inside NASInspectorWidget."""

from __future__ import annotations

import html
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QSizePolicy,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, COMBOBOX_STYLE_TOOLBAR
from modules.nas.core.packet_stack import (
    UP_PRESET_MAP,
    UP_SCENARIO_PRESETS,
    ComposedPacket,
    FlagOption,
    PacketConfig,
    compose_up_packet,
    copy_packet_to_clipboard,
    render_packet_html,
    render_packet_plantuml_packetdiag,
    resolve_up_packet_target,
)


def _flag_pill_qss(*, release_present: bool = True, is_ext: bool = False) -> str:
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


def _clear_flow_layout(flow: FlowLayout) -> None:
    while flow.count():
        item = flow.takeAt(0)
        if item and item.widget():
            item.widget().deleteLater()


class PacketStackWidget(QWidget):
    """Interactive octet-by-octet Full Packet Stack view for User Plane protocols."""

    open_spec_requested = pyqtSignal(str, str, str, str)
    inspect_field_requested = pyqtSignal(str, str, str)  # spec_number, clause, field_name
    send_plantuml_to_editor_requested = pyqtSignal(str)

    def __init__(self, db: Any = None, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.db = db
        self._base_target_frame: str = "dl_pdu_session"
        self._config = PacketConfig()
        self._composed: Optional[ComposedPacket] = None
        self._flag_buttons: Dict[str, QPushButton] = {}
        self._updating_ui = False
        self._settings = QSettings("3GPP_Tools", "ProtocolsWorkspaceLayout")
        self._setup_ui()

    def set_database(self, db: Any) -> None:
        self.db = db

    def current_packet(self) -> Optional[ComposedPacket]:
        return self._composed

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        controls = QFrame(self)
        controls.setObjectName("packetStackControls")
        controls.setStyleSheet(
            f"#packetStackControls {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        ctrl_lay = QVBoxLayout(controls)
        ctrl_lay.setContentsMargins(8, 6, 8, 6)
        ctrl_lay.setSpacing(5)

        # Row 1: Scenario preset, Outer transport, T-PDU toggle, Overhead pill, Copy menu
        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        ctrl_h = scale_px(25)

        lbl_sc = QLabel("Scenario:", controls)
        lbl_sc.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {C.TEXT_SECONDARY};")
        top_row.addWidget(lbl_sc)

        self.scenario_combo = QComboBox(controls)
        self.scenario_combo.setMinimumHeight(ctrl_h)
        self.scenario_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.scenario_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        for idx, p in enumerate(UP_SCENARIO_PRESETS):
            self.scenario_combo.addItem(p.label, p.preset_id)
            self.scenario_combo.setItemData(idx, p.description, Qt.ToolTipRole)
        self.scenario_combo.currentIndexChanged.connect(self._on_scenario_changed)
        top_row.addWidget(self.scenario_combo)

        self.btn_scenarios_help = QPushButton("? Scenarios", controls)
        self.btn_scenarios_help.setCheckable(True)
        self.btn_scenarios_help.setMinimumHeight(ctrl_h)
        self.btn_scenarios_help.setCursor(Qt.PointingHandCursor)
        self.btn_scenarios_help.setToolTip("Show or hide explanations for all 6 User Plane scenarios")
        self.btn_scenarios_help.setStyleSheet(_flag_pill_qss())
        self.btn_scenarios_help.toggled.connect(self._on_toggle_scenarios_guide)
        top_row.addWidget(self.btn_scenarios_help)

        self.transport_combo = QComboBox(controls)
        self.transport_combo.setMinimumHeight(ctrl_h)
        self.transport_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.transport_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.transport_combo.addItem("Outer: Collapsed", "collapsed")
        self.transport_combo.addItem("Outer: IPv4 + UDP (28B)", "ipv4_udp")
        self.transport_combo.addItem("Outer: IPv6 + UDP (48B)", "ipv6_udp")
        self.transport_combo.addItem("Outer: Hidden (3GPP Only)", "none")
        self.transport_combo.currentIndexChanged.connect(self._on_transport_changed)
        top_row.addWidget(self.transport_combo)

        self.tpdu_btn = QPushButton("📦 T-PDU", controls)
        self.tpdu_btn.setCheckable(True)
        self.tpdu_btn.setChecked(True)
        self.tpdu_btn.setMinimumHeight(ctrl_h)
        self.tpdu_btn.setCursor(Qt.PointingHandCursor)
        self.tpdu_btn.setToolTip("Toggle encapsulated T-PDU user payload (disable for dummy/control G-PDU)")
        self.tpdu_btn.setStyleSheet(_flag_pill_qss())
        self.tpdu_btn.toggled.connect(self._on_tpdu_toggled)
        top_row.addWidget(self.tpdu_btn)

        top_row.addStretch()

        self.overhead_badge = QLabel("3GPP UP: 12 B", controls)
        self.overhead_badge.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
            f"border: 1px solid {C.INFO_BORDER}; border-radius: 4px; "
            f"padding: 2px 7px; font-size: 10.5px; font-weight: bold;"
        )
        top_row.addWidget(self.overhead_badge)

        self.copy_packet_btn = QPushButton("📋 Copy Packet ▾", controls)
        self.copy_packet_btn.setMinimumHeight(ctrl_h)
        self.copy_packet_btn.setCursor(Qt.PointingHandCursor)
        self.copy_packet_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_packet_btn.setToolTip("Copy assembled packet diagram (Word, PlantUML packetdiag, Markdown, Text)")

        copy_menu = QMenu(self)
        copy_menu.setStyleSheet(
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
            act = QAction(text, self)
            act.triggered.connect(lambda _c=False, m=mode: self._copy_packet(m))
            copy_menu.addAction(act)
        copy_menu.addSeparator()
        act_ide = QAction("✏️ Open in PlantUML IDE", self)
        act_ide.triggered.connect(self._send_to_plantuml_ide)
        copy_menu.addAction(act_ide)

        self.copy_packet_btn.setMenu(copy_menu)
        top_row.addWidget(self.copy_packet_btn)
        ctrl_lay.addLayout(top_row)

        # Inline active scenario explanation strip
        self.scenario_info_lbl = QLabel("", controls)
        self.scenario_info_lbl.setObjectName("packetScenarioInfo")
        self.scenario_info_lbl.setWordWrap(True)
        self.scenario_info_lbl.setStyleSheet(
            f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; "
            f"padding: 3px 7px; font-size: 10.5px;"
        )
        ctrl_lay.addWidget(self.scenario_info_lbl)

        # Collapsible all-scenarios reference guide
        self.scenarios_guide_panel = QFrame(controls)
        self.scenarios_guide_panel.setObjectName("packetScenariosGuide")
        self.scenarios_guide_panel.setStyleSheet(
            f"#packetScenariosGuide {{ background-color: {C.SKY_SURFACE}; "
            f"border: 1px solid {C.SKY_BORDER}; border-radius: 5px; }}"
        )
        sg_lay = QVBoxLayout(self.scenarios_guide_panel)
        sg_lay.setContentsMargins(8, 6, 8, 6)
        sg_lay.setSpacing(3)
        self.scenarios_guide_lbl = QLabel(self._build_scenarios_guide_html(), self.scenarios_guide_panel)
        self.scenarios_guide_lbl.setWordWrap(True)
        self.scenarios_guide_lbl.setStyleSheet(f"font-size: 10.5px; color: {C.TEXT_BODY};")
        sg_lay.addWidget(self.scenarios_guide_lbl)
        self.scenarios_guide_panel.setVisible(False)
        ctrl_lay.addWidget(self.scenarios_guide_panel)

        # Row 2: GTP-U & Extension Header Chain toggles
        ext_row = QHBoxLayout()
        ext_row.setSpacing(6)
        lbl_ext = QLabel("Ext & GTP-U:", controls)
        lbl_ext.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.SUCCESS_DARK};")
        ext_row.addWidget(lbl_ext, 0, Qt.AlignTop)

        self.ext_flags_container = QWidget(controls)
        self.ext_flags_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.ext_flags_flow = FlowLayout(self.ext_flags_container, margin=0, h_spacing=4, v_spacing=4)
        ext_row.addWidget(self.ext_flags_container, 1)
        ctrl_lay.addLayout(ext_row)

        # Row 3: Inner TS 38.415 Frame Flags + Minimal / All Opt
        self.frame_flags_row_widget = QWidget(controls)
        flags_row = QHBoxLayout(self.frame_flags_row_widget)
        flags_row.setContentsMargins(0, 0, 0, 0)
        flags_row.setSpacing(6)

        lbl_frm = QLabel("Frame Flags:", self.frame_flags_row_widget)
        lbl_frm.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.ACCENT_SKY_DARK};")
        flags_row.addWidget(lbl_frm, 0, Qt.AlignTop)

        self.frame_flags_container = QWidget(self.frame_flags_row_widget)
        self.frame_flags_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.frame_flags_flow = FlowLayout(self.frame_flags_container, margin=0, h_spacing=4, v_spacing=4)
        flags_row.addWidget(self.frame_flags_container, 1)

        quick_btn_qss = (
            f"QPushButton {{ font-size: 10px; font-weight: 600; color: {C.TEXT_SECONDARY}; "
            f"background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 7px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; color: {C.TEXT_PRIMARY}; }}"
        )
        self.btn_flags_minimal = QPushButton("Minimal", self.frame_flags_row_widget)
        self.btn_flags_minimal.setCursor(Qt.PointingHandCursor)
        self.btn_flags_minimal.setStyleSheet(quick_btn_qss)
        self.btn_flags_minimal.setToolTip("Disable optional frame flags to inspect minimal overhead")
        self.btn_flags_minimal.clicked.connect(lambda: self._apply_quick_flags(False))
        flags_row.addWidget(self.btn_flags_minimal, 0, Qt.AlignTop)

        self.btn_flags_all = QPushButton("All Opt", self.frame_flags_row_widget)
        self.btn_flags_all.setCursor(Qt.PointingHandCursor)
        self.btn_flags_all.setStyleSheet(quick_btn_qss)
        self.btn_flags_all.setToolTip("Enable all release-supported optional flags in active frame(s)")
        self.btn_flags_all.clicked.connect(lambda: self._apply_quick_flags(True))
        flags_row.addWidget(self.btn_flags_all, 0, Qt.AlignTop)

        ctrl_lay.addWidget(self.frame_flags_row_widget)
        root.addWidget(controls)

        # Packet diagram HTML browser
        self.browser = QTextBrowser(self)
        self.browser.setReadOnly(True)
        self.browser.setOpenLinks(False)
        self.browser.setOpenExternalLinks(False)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setStyleSheet(
            f"QTextBrowser {{ border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; "
            f"padding: 6px; background-color: {C.BG_CARD}; }}"
        )
        bind_reader_zoom_settings(self.browser, self._settings, key="packet_stack/zoom_steps")
        root.addWidget(self.browser, 1)

        # Slim Field Quick-Detail Bar (shown when clicking a field inside the packet)
        self.detail_bar = QFrame(self)
        self.detail_bar.setObjectName("packetFieldDetailBar")
        self.detail_bar.setStyleSheet(
            f"#packetFieldDetailBar {{ background-color: {C.SKY_SURFACE}; "
            f"border: 1px solid {C.SKY_BORDER}; border-radius: 5px; }}"
        )
        det_lay = QHBoxLayout(self.detail_bar)
        det_lay.setContentsMargins(8, 4, 8, 4)
        det_lay.setSpacing(8)
        self.detail_lbl = QLabel("", self.detail_bar)
        self.detail_lbl.setWordWrap(True)
        self.detail_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_STRONG};")
        det_lay.addWidget(self.detail_lbl, 1)

        self.btn_open_field_def = QPushButton("📖 Open Definition", self.detail_bar)
        self.btn_open_field_def.setCursor(Qt.PointingHandCursor)
        self.btn_open_field_def.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_field_def.clicked.connect(self._on_open_selected_field_def)
        det_lay.addWidget(self.btn_open_field_def)

        self.btn_close_detail = QPushButton("✕", self.detail_bar)
        self.btn_close_detail.setCursor(Qt.PointingHandCursor)
        self.btn_close_detail.setStyleSheet(quick_btn_qss)
        self.btn_close_detail.clicked.connect(lambda: self.detail_bar.setVisible(False))
        det_lay.addWidget(self.btn_close_detail)

        self.detail_bar.setVisible(False)
        self._selected_field_ref: Optional[tuple[str, str, str]] = None
        root.addWidget(self.detail_bar)

    @staticmethod
    def _build_scenarios_guide_html() -> str:
        items = "".join(
            f"<div style='margin: 2px 0;'><b>{html.escape(p.label)}</b> "
            f"<span style='color: {C.TEXT_MUTED};'>[{html.escape(p.interface)}]</span>: "
            f"{html.escape(p.description)}</div>"
            for p in UP_SCENARIO_PRESETS
        )
        return f"<div><b>5G User Plane Scenarios (TS 29.281 &amp; TS 38.415):</b>{items}</div>"

    def _on_toggle_scenarios_guide(self, checked: bool) -> None:
        self.scenarios_guide_panel.setVisible(bool(checked))

    def set_context(
        self,
        *,
        db: Any = None,
        spec_number: Optional[str],
        clause: str = "",
        ie_name: str = "",
        message_name: str = "",
        version_hint: str = "",
    ) -> bool:
        """Resolve selection into a 5G UP packet configuration and render it."""
        if db is not None:
            self.db = db
        target = resolve_up_packet_target(
            spec_number=spec_number,
            clause=clause,
            ie_name=ie_name,
            message_name=message_name,
        )
        if not target:
            return False

        preset_id = target["preset_id"]
        preset = UP_PRESET_MAP.get(preset_id, UP_SCENARIO_PRESETS[0])
        flags = dict(preset.default_flags)
        flags.update(target.get("auto_enable_flags") or {})
        self._base_target_frame = target["target_frame"]

        ver_38415 = ""
        ver_29281 = ""
        clean_spec = str(spec_number or "").replace("TS", "").strip()
        if clean_spec == "38.415" and version_hint:
            ver_38415 = version_hint.strip()
        elif clean_spec == "29.281" and version_hint:
            ver_29281 = version_hint.strip()

        self._config = PacketConfig(
            target_frame=self._base_target_frame,
            preset_id=preset_id,
            transport_mode=self._config.transport_mode or preset.transport_mode,
            flags=flags,
            include_tpdu=preset.include_tpdu,
            spec_version_38415=ver_38415,
            spec_version_29281=ver_29281,
            highlight_field=target.get("highlight_field", ""),
        )
        self.detail_bar.setVisible(False)
        self.refresh_packet()
        return True

    def refresh_packet(self) -> None:
        """Recompute `ComposedPacket`, sync controls, and render HTML."""
        self._composed = compose_up_packet(self._config, db=self.db)
        self._config = self._composed.config
        preset = UP_PRESET_MAP.get(self._config.preset_id, UP_SCENARIO_PRESETS[0])

        self._updating_ui = True
        try:
            idx_sc = self.scenario_combo.findData(self._config.preset_id)
            if idx_sc >= 0:
                self.scenario_combo.setCurrentIndex(idx_sc)
            idx_tr = self.transport_combo.findData(self._config.transport_mode)
            if idx_tr >= 0:
                self.transport_combo.setCurrentIndex(idx_tr)
            self.tpdu_btn.setChecked(self._config.include_tpdu)
            self.tpdu_btn.setEnabled(self._config.target_frame != "end_marker")
            self.scenario_info_lbl.setText(f"ℹ️ {preset.description}")
            self._rebuild_flag_pills()
        finally:
            self._updating_ui = False

        ov_txt = f"3GPP UP: {self._composed.overhead_3gpp_octets} B"
        if self._composed.overhead_transport_octets > 0:
            ov_txt += f" · Total: {self._composed.overhead_total_octets} B"
        self.overhead_badge.setText(ov_txt)
        self.browser.setHtml(render_packet_html(self._composed, interactive_links=True))

    def _add_pill(self, opt: FlagOption, container: QWidget, flow: FlowLayout, *, is_ext: bool) -> None:
        btn = QPushButton(opt.label, container)
        btn.setCheckable(True)
        btn.setChecked(bool(self._config.flags.get(opt.key, opt.default_enabled)))
        btn.setCursor(Qt.PointingHandCursor)
        tip = opt.tooltip
        if not opt.is_release_present:
            tip += " (Not in selected release DB; click to preview anyway)"
        btn.setToolTip(tip)
        btn.setStyleSheet(_flag_pill_qss(release_present=opt.is_release_present, is_ext=is_ext))
        btn.toggled.connect(lambda checked, k=opt.key: self._on_flag_toggled(k, checked))
        flow.addWidget(btn)
        self._flag_buttons[opt.key] = btn

    def _rebuild_flag_pills(self) -> None:
        _clear_flow_layout(self.ext_flags_flow)
        _clear_flow_layout(self.frame_flags_flow)
        self._flag_buttons.clear()
        if not self._composed:
            return
        frame_opts: List[FlagOption] = []
        for opt in self._composed.available_flags:
            if opt.group in ("gtpu", "ext"):
                self._add_pill(opt, self.ext_flags_container, self.ext_flags_flow, is_ext=True)
            else:
                frame_opts.append(opt)
                self._add_pill(opt, self.frame_flags_container, self.frame_flags_flow, is_ext=False)
        self.frame_flags_row_widget.setVisible(bool(frame_opts))

    def _on_scenario_changed(self, index: int) -> None:
        if self._updating_ui or index < 0:
            return
        preset_id = str(self.scenario_combo.itemData(index) or "n3_n9")
        preset = UP_PRESET_MAP.get(preset_id)
        if not preset:
            return
        self._config.preset_id = preset_id
        if preset_id == "end_marker":
            self._config.target_frame = "end_marker"
        elif preset_id == "n3_n9_pdu_set":
            self._config.target_frame = "ul_pdu_session" if self._base_target_frame == "ul_pdu_session" else "dl_pdu_session"
        elif self._base_target_frame == "end_marker":
            self._config.target_frame = "dl_pdu_session"
        else:
            self._config.target_frame = self._base_target_frame
        self._config.flags = dict(preset.default_flags)
        if self._base_target_frame == "dl_pdu_set" and preset_id in ("n3_n9_pdu_set", "f1u_cu_du"):
            self._config.flags["ext_pdu_set"] = True
        self._config.include_tpdu = preset.include_tpdu
        self.refresh_packet()

    def _on_transport_changed(self, index: int) -> None:
        if self._updating_ui or index < 0:
            return
        self._config.transport_mode = str(self.transport_combo.itemData(index) or "collapsed")
        self.refresh_packet()

    def _on_tpdu_toggled(self, checked: bool) -> None:
        if self._updating_ui:
            return
        self._config.include_tpdu = bool(checked)
        self.refresh_packet()

    def _on_flag_toggled(self, key: str, checked: bool) -> None:
        if self._updating_ui:
            return
        self._config.flags[key] = bool(checked)
        self.refresh_packet()

    def _apply_quick_flags(self, enable_all: bool) -> None:
        if not self._composed:
            return
        for opt in self._composed.available_flags:
            if opt.group.startswith("frame"):
                self._config.flags[opt.key] = bool(enable_all and opt.is_release_present)
        self.refresh_packet()

    def _copy_packet(self, mode: str) -> None:
        if not self._composed:
            return
        copy_packet_to_clipboard(self._composed, mode=mode)
        self.copy_packet_btn.setText("✓ Copied!")
        QTimer.singleShot(1500, lambda: self.copy_packet_btn.setText("📋 Copy Packet ▾"))

    def _send_to_plantuml_ide(self) -> None:
        if not self._composed:
            return
        self.send_plantuml_to_editor_requested.emit(render_packet_plantuml_packetdiag(self._composed))

    def _on_anchor_clicked(self, url: QUrl) -> None:
        url_str = url.toString()
        if url_str.startswith("action:open_spec:"):
            parts = url_str.replace("action:open_spec:", "").split(":")
            self.open_spec_requested.emit(
                parts[0],
                parts[1] if len(parts) > 1 else "",
                parts[2] if len(parts) > 2 else "",
                "word",
            )
            return

        if url_str.startswith("action:packet_field:"):
            parts = url_str.replace("action:packet_field:", "").split(":", 2)
            spec_num = parts[0] if len(parts) > 0 else ""
            clause = parts[1] if len(parts) > 1 else ""
            field_name = parts[2] if len(parts) > 2 else ""
            fld = self._composed.find_field(field_name) if self._composed else None
            desc = fld.description if fld and fld.description else (f"Condition: {fld.condition}" if fld and fld.condition else "")
            size_txt = f" · {fld.size_label}" if fld else ""
            spec_txt = f"TS {spec_num} Cl. {clause}" if spec_num else ""
            self.detail_lbl.setText(f"<b>{field_name}</b> ({spec_txt}{size_txt})" + (f" — {desc}" if desc else ""))
            self._selected_field_ref = (spec_num, clause, field_name)
            self.btn_open_field_def.setVisible(spec_num in ("29.281", "38.415"))
            self.detail_bar.setVisible(True)

    def _on_open_selected_field_def(self) -> None:
        if self._selected_field_ref:
            spec_num, clause, field_name = self._selected_field_ref
            self.inspect_field_requested.emit(spec_num, clause, field_name)
