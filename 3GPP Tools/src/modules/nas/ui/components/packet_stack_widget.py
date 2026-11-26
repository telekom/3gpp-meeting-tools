"""Interactive 5G User Plane Full Packet Stack widget inside NASInspectorWidget."""

from __future__ import annotations

from typing import Any, Dict, List, Optional
from urllib.parse import unquote

from PyQt5.QtCore import QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, COMBOBOX_STYLE_TOOLBAR
from modules.nas.core.packet_stack import (
    SAMPLE_NAS_MESSAGES,
    ComposedPacket,
    FlagOption,
    PacketConfig,
    compose_packet,
    copy_packet_to_clipboard,
    default_nas_message_for_ngap,
    get_preset_map_for_family,
    get_presets_for_family,
    render_packet_html,
    render_packet_plantuml_packetdiag,
    resolve_packet_target,
)
from .packet_ie_resolver import resolve_packet_field
from .packet_ie_side_panel import PacketIESidePanel
from .packet_stack_styles import (
    build_packet_copy_menu, build_scenarios_guide_html,
    clear_flow_layout, flag_pill_qss, quick_btn_qss,
)


class PacketStackWidget(QWidget):
    """Interactive octet-by-octet Full Packet Stack view with inline IE definition Side Panel."""

    open_spec_requested = pyqtSignal(str, str, str, str)
    inspect_field_requested = pyqtSignal(str, str, str)  # spec_number, clause, field_name
    send_plantuml_to_editor_requested = pyqtSignal(str)

    def __init__(self, db: Any = None, parent: Optional[QWidget] = None, *, settings: Optional[QSettings] = None) -> None:
        super().__init__(parent)
        self.db = db
        self._base_target_frame: str = "dl_pdu_session"
        self._active_family: str = ""
        self._config = PacketConfig()
        self._composed: Optional[ComposedPacket] = None
        self._flag_buttons: Dict[str, QPushButton] = {}
        self._updating_ui = False
        self._settings = settings if settings is not None else QSettings("3GPP_Tools", "ProtocolsWorkspaceLayout")
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
        self.scenario_combo.currentIndexChanged.connect(self._on_scenario_changed)
        top_row.addWidget(self.scenario_combo)

        self.btn_scenarios_help = QPushButton("? Scenarios", controls)
        self.btn_scenarios_help.setCheckable(True)
        self.btn_scenarios_help.setMinimumHeight(ctrl_h)
        self.btn_scenarios_help.setCursor(Qt.PointingHandCursor)
        self.btn_scenarios_help.setToolTip("Show or hide explanations for all scenarios in this protocol family")
        self.btn_scenarios_help.setStyleSheet(flag_pill_qss())
        self.btn_scenarios_help.toggled.connect(self._on_toggle_scenarios_guide)
        top_row.addWidget(self.btn_scenarios_help)

        self.transport_combo = QComboBox(controls)
        self.transport_combo.setMinimumHeight(ctrl_h)
        self.transport_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.transport_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.transport_combo.currentIndexChanged.connect(self._on_transport_changed)
        top_row.addWidget(self.transport_combo)

        self.nas_msg_combo = QComboBox(controls)
        self.nas_msg_combo.setMinimumHeight(ctrl_h)
        self.nas_msg_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.nas_msg_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.nas_msg_combo.setToolTip("Select 5GS NAS message (TS 24.501) carried inside NGAP NAS-PDU")
        for nm in SAMPLE_NAS_MESSAGES:
            self.nas_msg_combo.addItem(f"NAS: {nm}", nm)
        self.nas_msg_combo.currentIndexChanged.connect(self._on_nas_msg_changed)
        self.nas_msg_combo.setVisible(False)
        top_row.addWidget(self.nas_msg_combo)

        self.tpdu_btn = QPushButton("📦 T-PDU", controls)
        self.tpdu_btn.setCheckable(True)
        self.tpdu_btn.setChecked(True)
        self.tpdu_btn.setMinimumHeight(ctrl_h)
        self.tpdu_btn.setCursor(Qt.PointingHandCursor)
        self.tpdu_btn.setToolTip("Toggle encapsulated T-PDU user payload (disable for dummy/control G-PDU)")
        self.tpdu_btn.setStyleSheet(flag_pill_qss())
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

        self.copy_packet_btn.setMenu(
            build_packet_copy_menu(self, self._copy_packet, self._send_to_plantuml_ide)
        )
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
        self.scenarios_guide_lbl = QLabel(build_scenarios_guide_html("up"), self.scenarios_guide_panel)
        self.scenarios_guide_lbl.setWordWrap(True)
        self.scenarios_guide_lbl.setStyleSheet(f"font-size: 10.5px; color: {C.TEXT_BODY};")
        sg_lay.addWidget(self.scenarios_guide_lbl)
        self.scenarios_guide_panel.setVisible(False)
        ctrl_lay.addWidget(self.scenarios_guide_panel)

        # Row 2: GTP-U / SCTP & NAS toggles
        ext_row = QHBoxLayout()
        ext_row.setSpacing(6)
        self.lbl_ext_row = QLabel("Ext & GTP-U:", controls)
        self.lbl_ext_row.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.SUCCESS_DARK};")
        ext_row.addWidget(self.lbl_ext_row, 0, Qt.AlignTop)

        self.ext_flags_container = QWidget(controls)
        self.ext_flags_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.ext_flags_flow = FlowLayout(self.ext_flags_container, margin=0, h_spacing=4, v_spacing=4)
        ext_row.addWidget(self.ext_flags_container, 1)
        ctrl_lay.addLayout(ext_row)

        # Row 3: Inner TS 38.415 Frame Flags or NGAP Optional IEs + Minimal / All Opt
        self.frame_flags_row_widget = QWidget(controls)
        flags_row = QHBoxLayout(self.frame_flags_row_widget)
        flags_row.setContentsMargins(0, 0, 0, 0)
        flags_row.setSpacing(6)

        self.lbl_frame_row = QLabel("Frame Flags:", self.frame_flags_row_widget)
        self.lbl_frame_row.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.ACCENT_SKY_DARK};")
        flags_row.addWidget(self.lbl_frame_row, 0, Qt.AlignTop)

        self.frame_flags_container = QWidget(self.frame_flags_row_widget)
        self.frame_flags_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.frame_flags_flow = FlowLayout(self.frame_flags_container, margin=0, h_spacing=4, v_spacing=4)
        flags_row.addWidget(self.frame_flags_container, 1)

        q_btn_qss = quick_btn_qss()
        self.btn_flags_minimal = QPushButton("Minimal", self.frame_flags_row_widget)
        self.btn_flags_minimal.setCursor(Qt.PointingHandCursor)
        self.btn_flags_minimal.setStyleSheet(q_btn_qss)
        self.btn_flags_minimal.setToolTip("Disable optional frame flags to inspect minimal overhead")
        self.btn_flags_minimal.clicked.connect(lambda: self._apply_quick_flags(False))
        flags_row.addWidget(self.btn_flags_minimal, 0, Qt.AlignTop)

        self.btn_flags_all = QPushButton("All Opt", self.frame_flags_row_widget)
        self.btn_flags_all.setCursor(Qt.PointingHandCursor)
        self.btn_flags_all.setStyleSheet(q_btn_qss)
        self.btn_flags_all.setToolTip("Enable all release-supported optional flags in active frame(s)")
        self.btn_flags_all.clicked.connect(lambda: self._apply_quick_flags(True))
        flags_row.addWidget(self.btn_flags_all, 0, Qt.AlignTop)

        ctrl_lay.addWidget(self.frame_flags_row_widget)
        root.addWidget(controls)

        # Packet diagram HTML browser + collapsible IE Definition Side Panel
        orient_val = str(self._settings.value("packet_stack/side_panel_orientation", "h") or "h").lower()
        initial_orient = Qt.Vertical if orient_val == "v" else Qt.Horizontal
        self.body_splitter = QSplitter(initial_orient, self)
        self.body_splitter.setChildrenCollapsible(False)
        self.body_splitter.splitterMoved.connect(self._on_body_splitter_moved)

        self.browser = QTextBrowser(self.body_splitter)
        self.browser.setReadOnly(True)
        self.browser.setOpenLinks(False)
        self.browser.setOpenExternalLinks(False)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setStyleSheet(
            f"QTextBrowser {{ border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; "
            f"padding: 6px; background-color: {C.BG_CARD}; }}"
        )
        bind_reader_zoom_settings(self.browser, self._settings, key="packet_stack/zoom_steps")
        self.body_splitter.addWidget(self.browser)

        self.side_panel = PacketIESidePanel(self.body_splitter, settings=self._settings)
        self.side_panel.set_docked_bottom(initial_orient == Qt.Vertical)
        self.side_panel.setVisible(False)
        self.side_panel.open_spec_requested.connect(self.open_spec_requested.emit)
        self.side_panel.expand_requested.connect(self.inspect_field_requested.emit)
        self.side_panel.navigate_field_requested.connect(
            lambda sp, cl, fn: self.show_field_in_side_panel(sp, cl, fn, push_history=True)
        )
        self.side_panel.close_requested.connect(self.hide_side_panel)
        self.side_panel.dock_toggle_requested.connect(self.toggle_side_panel_dock)
        self.body_splitter.addWidget(self.side_panel)

        root.addWidget(self.body_splitter, 1)
        self._populate_family_controls("up")

    def _populate_family_controls(self, stack_family: str) -> None:
        if self._active_family == stack_family:
            return
        self._active_family = stack_family
        self._updating_ui = True
        try:
            self.scenario_combo.clear()
            for idx, p in enumerate(get_presets_for_family(stack_family)):
                self.scenario_combo.addItem(p.label, p.preset_id)
                self.scenario_combo.setItemData(idx, p.description, Qt.ToolTipRole)

            self.transport_combo.clear()
            if stack_family == "ngap":
                self.transport_combo.addItem("Outer: Collapsed", "collapsed")
                self.transport_combo.addItem("Outer: IPv4 + SCTP (48B)", "ipv4_sctp")
                self.transport_combo.addItem("Outer: IPv6 + SCTP (68B)", "ipv6_sctp")
                self.transport_combo.addItem("Outer: SCTP Only (28B)", "sctp_only")
                self.transport_combo.addItem("Outer: NGAP APER Only", "none")
                self.lbl_ext_row.setText("SCTP & NAS:")
                self.lbl_frame_row.setText("Optional IEs:")
                self.tpdu_btn.setVisible(False)
            else:
                self.transport_combo.addItem("Outer: Collapsed", "collapsed")
                self.transport_combo.addItem("Outer: IPv4 + UDP (28B)", "ipv4_udp")
                self.transport_combo.addItem("Outer: IPv6 + UDP (48B)", "ipv6_udp")
                self.transport_combo.addItem("Outer: Hidden (3GPP Only)", "none")
                self.lbl_ext_row.setText("Ext & GTP-U:")
                self.lbl_frame_row.setText("Frame Flags:")
                self.tpdu_btn.setVisible(True)
            self.scenarios_guide_lbl.setText(build_scenarios_guide_html(stack_family))
        finally:
            self._updating_ui = False

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
        """Resolve selection into a 5G UP or N2 NGAP packet configuration and render it."""
        if db is not None:
            self.db = db
        target = resolve_packet_target(
            spec_number=spec_number,
            clause=clause,
            ie_name=ie_name,
            message_name=message_name,
        )
        if not target:
            return False

        family = target.get("stack_family", "up")
        self._populate_family_controls(family)
        preset_id = target["preset_id"]
        preset_map = get_preset_map_for_family(family)
        presets = get_presets_for_family(family)
        preset = preset_map.get(preset_id, presets[0])
        flags = dict(preset.default_flags)
        flags.update(target.get("auto_enable_flags") or {})
        self._base_target_frame = target["target_frame"]

        ver_38415 = ver_29281 = ver_38413 = ""
        clean_spec = str(spec_number or "").replace("TS", "").strip()
        if clean_spec == "38.415" and version_hint:
            ver_38415 = version_hint.strip()
        elif clean_spec == "29.281" and version_hint:
            ver_29281 = version_hint.strip()
        elif clean_spec == "38.413" and version_hint:
            ver_38413 = version_hint.strip()

        msg_name = target.get("message_name", "")
        nas_name = default_nas_message_for_ngap(msg_name) if family == "ngap" else ""
        t_mode = self._config.transport_mode if self._config.stack_family == family else preset.transport_mode
        self._config = PacketConfig(
            stack_family=family,
            target_frame=self._base_target_frame,
            preset_id=preset_id,
            transport_mode=t_mode or preset.transport_mode,
            flags=flags,
            include_tpdu=preset.include_tpdu,
            message_name=msg_name,
            nas_message_name=nas_name,
            spec_version_38415=ver_38415,
            spec_version_29281=ver_29281,
            spec_version_38413=ver_38413,
            highlight_field=target.get("highlight_field", ""),
        )
        self.side_panel.setVisible(False)
        self.side_panel.clear()
        self.refresh_packet()
        return True

    def refresh_packet(self) -> None:
        """Recompute `ComposedPacket`, sync controls, and render HTML."""
        self._composed = compose_packet(self._config, db=self.db)
        self._config = self._composed.config
        family = self._config.stack_family
        preset_map = get_preset_map_for_family(family)
        presets = get_presets_for_family(family)
        preset = preset_map.get(self._config.preset_id, presets[0])

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
            show_nas_combo = family == "ngap" and bool(self._config.flags.get("expand_nas_pdu"))
            self.nas_msg_combo.setVisible(show_nas_combo)
            if show_nas_combo:
                idx_nas = self.nas_msg_combo.findData(self._config.nas_message_name)
                if idx_nas >= 0:
                    self.nas_msg_combo.setCurrentIndex(idx_nas)
            self.scenario_info_lbl.setText(f"ℹ️ {preset.description}")
            self._rebuild_flag_pills()
        finally:
            self._updating_ui = False

        lbl_prefix = "N2 / NGAP" if family == "ngap" else "3GPP UP"
        ov_txt = f"{lbl_prefix}: {self._composed.overhead_3gpp_octets} B"
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
        btn.setStyleSheet(flag_pill_qss(release_present=opt.is_release_present, is_ext=is_ext))
        btn.toggled.connect(lambda checked, k=opt.key: self._on_flag_toggled(k, checked))
        flow.addWidget(btn)
        self._flag_buttons[opt.key] = btn

    def _rebuild_flag_pills(self) -> None:
        clear_flow_layout(self.ext_flags_flow)
        clear_flow_layout(self.frame_flags_flow)
        self._flag_buttons.clear()
        if not self._composed:
            return
        frame_opts: List[FlagOption] = []
        for opt in self._composed.available_flags:
            if opt.group in ("gtpu", "ext", "sctp"):
                self._add_pill(opt, self.ext_flags_container, self.ext_flags_flow, is_ext=True)
            else:
                frame_opts.append(opt)
                self._add_pill(opt, self.frame_flags_container, self.frame_flags_flow, is_ext=False)
        self.frame_flags_row_widget.setVisible(bool(frame_opts))

    def _on_scenario_changed(self, index: int) -> None:
        if self._updating_ui or index < 0:
            return
        family = self._config.stack_family
        preset_id = str(self.scenario_combo.itemData(index) or "")
        preset = get_preset_map_for_family(family).get(preset_id)
        if not preset:
            return
        self._config.preset_id = preset_id
        if family == "ngap":
            self._config.flags = dict(preset.default_flags)
            self.refresh_packet()
            return
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

    def _on_nas_msg_changed(self, index: int) -> None:
        if self._updating_ui or index < 0:
            return
        self._config.nas_message_name = str(self.nas_msg_combo.itemData(index) or "")
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
            if opt.group.startswith("frame") or opt.group == "ngap_ie":
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

    def _version_hint_for_spec(self, spec_number: str) -> str:
        clean = (spec_number or "").replace("TS", "").strip()
        if clean == "38.415":
            return self._config.spec_version_38415
        if clean == "29.281":
            return self._config.spec_version_29281
        if clean == "38.413":
            return self._config.spec_version_38413
        return ""

    def _apply_initial_splitter_sizes(self) -> None:
        raw = self._settings.value("packet_stack/side_panel_sizes")
        if isinstance(raw, list) and len(raw) == 2:
            try:
                sizes = [int(raw[0]), int(raw[1])]
                if sizes[0] > 80 and sizes[1] > 80:
                    self.body_splitter.setSizes(sizes)
                    return
            except (TypeError, ValueError):
                pass
        span = self.body_splitter.width() if self.body_splitter.orientation() == Qt.Horizontal else self.body_splitter.height()
        total = max(600, span)
        self.body_splitter.setSizes([int(total * 0.6), int(total * 0.4)])

    def _on_body_splitter_moved(self, _pos: int, _index: int) -> None:
        if not self.side_panel.isVisible():
            return
        sizes = self.body_splitter.sizes()
        if len(sizes) == 2 and sizes[0] > 0 and sizes[1] > 0:
            self._settings.setValue("packet_stack/side_panel_sizes", sizes)

    def show_field_in_side_panel(
        self,
        spec_number: str,
        clause: str,
        field_name: str,
        *,
        push_history: bool = False,
    ) -> None:
        """Resolve and display a packet field/IE inline in the collapsible Side Panel."""
        res = resolve_packet_field(
            self.db,
            self._composed,
            spec_number,
            clause,
            field_name,
            version_hint=self._version_hint_for_spec(spec_number),
        )
        self.side_panel.show_result(res, push_history=push_history)
        if self.side_panel.isHidden():
            self.side_panel.setVisible(True)
            self._apply_initial_splitter_sizes()

    def hide_side_panel(self) -> None:
        self.side_panel.setVisible(False)

    def toggle_side_panel_dock(self) -> None:
        is_horiz = self.body_splitter.orientation() == Qt.Horizontal
        new_orient = Qt.Vertical if is_horiz else Qt.Horizontal
        self.body_splitter.setOrientation(new_orient)
        self.side_panel.set_docked_bottom(new_orient == Qt.Vertical)
        self._settings.setValue("packet_stack/side_panel_orientation", "v" if new_orient == Qt.Vertical else "h")
        if self.side_panel.isVisible():
            span = self.body_splitter.width() if new_orient == Qt.Horizontal else self.body_splitter.height()
            total = max(600, span)
            self.body_splitter.setSizes([int(total * 0.58), int(total * 0.42)])

    def _on_anchor_clicked(self, url: QUrl) -> None:
        url_str = url.toString()
        if url_str.startswith("action:open_spec:"):
            parts = [unquote(p) for p in url_str.replace("action:open_spec:", "").split(":")]
            spec_num = parts[0] if len(parts) > 0 else ""
            clause = parts[2] if len(parts) > 2 else ""
            self.show_field_in_side_panel(spec_num, clause, "")
            return

        if url_str.startswith("action:packet_field:"):
            parts = [unquote(p) for p in url_str.replace("action:packet_field:", "").split(":", 2)]
            spec_num = parts[0] if len(parts) > 0 else ""
            clause = parts[1] if len(parts) > 1 else ""
            field_name = parts[2] if len(parts) > 2 else ""
            self.show_field_in_side_panel(spec_num, clause, field_name)
