"""Dialog widget construction, preference loading, and venue card rendering mixin."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C, R, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SPEC_FILTER_COMBO_STYLE,
)
from core.ui.world_clock.venue_card import _VenueClockCard
from core.ui.world_clock_service import VenueClockState, WorldClockSnapshot
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset
from modules.meetings.ui.tdocs_components import CheckableComboBox


class WorldClockUiSetupMixin:
    """Builds reference clocks, venue list, WG preference controls, and live updates."""

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # --- 1. Reference Clocks Card (Laptop + UTC + Physical IP Detector) ---
        ref_card = QFrame()
        ref_card.setStyleSheet(CARD_FRAME_STYLE)
        ref_layout = QVBoxLayout(ref_card)
        ref_layout.setContentsMargins(14, 12, 14, 12)
        ref_layout.setSpacing(8)

        clocks_row = QHBoxLayout()
        clocks_row.setSpacing(18)

        # Laptop clock column
        lap_box = QVBoxLayout()
        lap_box.setSpacing(2)
        lap_hdr = QLabel("💻 Laptop Local Time (IT Policy)")
        lap_hdr.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; font-weight: 600; color: {C.TEXT_MUTED};")
        self.laptop_time_lbl = QLabel("--:--:--")
        self.laptop_time_lbl.setStyleSheet(
            f"font-family: {T.FONT_FAMILY_MONO}; font-size: 22px; font-weight: 700; color: {C.TEXT_PRIMARY};"
        )
        self.laptop_sub_lbl = QLabel("")
        self.laptop_sub_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        lap_box.addWidget(lap_hdr)
        lap_box.addWidget(self.laptop_time_lbl)
        lap_box.addWidget(self.laptop_sub_lbl)
        clocks_row.addLayout(lap_box, 1)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        clocks_row.addWidget(sep1)

        # UTC clock column
        utc_box = QVBoxLayout()
        utc_box.setSpacing(2)
        utc_hdr = QLabel("🌐 Official 3GPP Time (UTC)")
        utc_hdr.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; font-weight: 600; color: {C.TEXT_MUTED};")
        self.utc_time_lbl = QLabel("--:--:-- UTC")
        self.utc_time_lbl.setStyleSheet(
            f"font-family: {T.FONT_FAMILY_MONO}; font-size: 22px; font-weight: 700; color: {C.ACCENT_PRIMARY};"
        )
        self.utc_sub_lbl = QLabel("")
        self.utc_sub_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        utc_box.addWidget(utc_hdr)
        utc_box.addWidget(self.utc_time_lbl)
        utc_box.addWidget(self.utc_sub_lbl)
        clocks_row.addLayout(utc_box, 1)

        ref_layout.addLayout(clocks_row)

        # Physical IP detection bar
        ip_row = QHBoxLayout()
        ip_row.setSpacing(8)
        self.btn_detect_ip = QPushButton("📡 Detect My Physical Timezone (via Public IP)")
        self.btn_detect_ip.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_detect_ip.setCursor(Qt.PointingHandCursor)
        self.btn_detect_ip.setToolTip(
            "Queries a public Geo-IP timezone server to detect your current physical timezone "
            "(useful when not on a corporate full-tunnel VPN)."
        )
        self.btn_detect_ip.clicked.connect(self._on_detect_ip_clicked)
        ip_row.addWidget(self.btn_detect_ip)

        self.ip_status_lbl = QLabel("")
        self.ip_status_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        ip_row.addWidget(self.ip_status_lbl, 1)

        self.btn_apply_ip = QPushButton("Apply to Venue")
        self.btn_apply_ip.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_apply_ip.setCursor(Qt.PointingHandCursor)
        self.btn_apply_ip.setVisible(False)
        self.btn_apply_ip.clicked.connect(self._on_apply_ip_to_primary)
        ip_row.addWidget(self.btn_apply_ip)

        ref_layout.addLayout(ip_row)
        root.addWidget(ref_card)

        # --- 2. Scrollable Active Venues List ---
        venues_hdr_row = QHBoxLayout()
        venues_hdr = QLabel("<b>📍 3GPP Meeting Venues</b>")
        venues_hdr.setStyleSheet(f"font-size: {T.SIZE_SECTION_TITLE}px; color: {C.TEXT_PRIMARY};")
        venues_hdr_row.addWidget(venues_hdr)
        venues_hdr_row.addStretch(1)

        self.btn_sync_all = QPushButton("🔄 Sync All Venues from Web")
        self.btn_sync_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_sync_all.setCursor(Qt.PointingHandCursor)
        self.btn_sync_all.clicked.connect(self._on_sync_all_venues)
        venues_hdr_row.addWidget(self.btn_sync_all)
        root.addLayout(venues_hdr_row)

        self.venues_scroll = QScrollArea()
        self.venues_scroll.setWidgetResizable(True)
        self.venues_scroll.setFrameShape(QFrame.NoFrame)
        self.venues_scroll.setStyleSheet("QScrollArea { background: transparent; }")

        self.venues_container = QWidget()
        self.venues_container.setStyleSheet("background: transparent;")
        self.venues_layout = QVBoxLayout(self.venues_container)
        self.venues_layout.setContentsMargins(0, 0, 4, 0)
        self.venues_layout.setSpacing(8)
        self.venues_scroll.setWidget(self.venues_container)
        root.addWidget(self.venues_scroll, 1)

        # --- 3. Working Group Filter & Bar Display Preferences ---
        pref_card = QFrame()
        pref_card.setStyleSheet(CARD_FRAME_STYLE)
        pref_layout = QVBoxLayout(pref_card)
        pref_layout.setContentsMargins(14, 10, 14, 10)
        pref_layout.setSpacing(6)

        pref_title = QLabel("<b>🎯 Status Bar Display & Working Group Filter</b>")
        pref_title.setStyleSheet(f"font-size: {T.SIZE_BODY}px; color: {C.TEXT_PRIMARY};")
        pref_layout.addWidget(pref_title)

        wg_row = QHBoxLayout()
        wg_row.setSpacing(10)
        self.wg_mode_group = QButtonGroup(self)
        self.rb_all_wgs = QRadioButton("All Active Working Groups")
        self.rb_selected_wgs = QRadioButton("Only Selected WGs:")
        self.wg_mode_group.addButton(self.rb_all_wgs, 0)
        self.wg_mode_group.addButton(self.rb_selected_wgs, 1)

        self.wg_combo = CheckableComboBox("WGs", self)
        self.wg_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_combo.setMinimumWidth(180)
        self.wg_combo.addItems(self.service.get_available_working_groups())

        wg_row.addWidget(self.rb_all_wgs)
        wg_row.addWidget(self.rb_selected_wgs)
        wg_row.addWidget(self.wg_combo)
        wg_row.addStretch(1)
        pref_layout.addLayout(wg_row)

        opts_row = QHBoxLayout()
        opts_row.setSpacing(16)
        self.chk_show_utc = QCheckBox("Show UTC clock in status bar")
        self.chk_show_upcoming = QCheckBox("Show next upcoming venue during off-weeks")
        opts_row.addWidget(self.chk_show_utc)
        opts_row.addWidget(self.chk_show_upcoming)
        opts_row.addStretch(1)
        pref_layout.addLayout(opts_row)

        root.addWidget(pref_card)

        # --- 4. Footer ---
        footer = QHBoxLayout()
        self.footer_hint = QLabel("Overrides save to both the meeting folder (meeting_timezone.json) and app config.")
        self.footer_hint.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED};")
        footer.addWidget(self.footer_hint, 1)

        self.btn_close = QPushButton("Close")
        self.btn_close.setObjectName("primaryBtn")
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.clicked.connect(self.accept)
        footer.addWidget(self.btn_close)
        root.addLayout(footer)

        # Wire preference change signals after widgets exist
        self.rb_all_wgs.toggled.connect(self._on_preferences_edited)
        self.rb_selected_wgs.toggled.connect(self._on_preferences_edited)
        self.wg_combo.selectionChanged.connect(self._on_wg_combo_changed)
        self.chk_show_utc.toggled.connect(self._on_preferences_edited)
        self.chk_show_upcoming.toggled.connect(self._on_preferences_edited)

    def _load_preferences_into_ui(self) -> None:
        cfg = self.service.settings.get_clock_config()
        wg_filter = [str(w).strip() for w in (cfg.get("wg_filter") or []) if str(w).strip()]

        for w in (self.rb_all_wgs, self.rb_selected_wgs, self.wg_combo, self.chk_show_utc, self.chk_show_upcoming):
            w.blockSignals(True)

        self.chk_show_utc.setChecked(bool(cfg.get("show_utc_in_bar", True)))
        self.chk_show_upcoming.setChecked(bool(cfg.get("show_upcoming_when_idle", True)))

        if wg_filter:
            self.rb_selected_wgs.setChecked(True)
            self.wg_combo.setEnabled(True)
            selected_upper = {w.upper() for w in wg_filter}
            model = self.wg_combo.model()
            all_checked = True
            for i in range(1, model.rowCount()):
                item = model.item(i)
                if item:
                    is_sel = str(item.data(Qt.UserRole) or item.text()).strip().upper() in selected_upper
                    item.setCheckState(Qt.Checked if is_sel else Qt.Unchecked)
                    if not is_sel:
                        all_checked = False
            root_item = model.item(0)
            if root_item:
                root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
            self.wg_combo._refresh_text_and_tooltip()
            self.wg_combo.update()
        else:
            self.rb_all_wgs.setChecked(True)
            self.wg_combo.setEnabled(False)

        for w in (self.rb_all_wgs, self.rb_selected_wgs, self.wg_combo, self.chk_show_utc, self.chk_show_upcoming):
            w.blockSignals(False)

    def _on_wg_combo_changed(self, _checked: list) -> None:
        if not self.rb_selected_wgs.isChecked():
            self.rb_selected_wgs.blockSignals(True)
            self.rb_selected_wgs.setChecked(True)
            self.rb_selected_wgs.blockSignals(False)
        self._on_preferences_edited()

    def _on_preferences_edited(self) -> None:
        use_selected = self.rb_selected_wgs.isChecked()
        self.wg_combo.setEnabled(use_selected)

        cfg = self.service.settings.get_clock_config()
        if use_selected:
            checked = self.wg_combo.getCheckedItems()
            total_wgs = max(0, self.wg_combo.model().rowCount() - 1)
            if len(checked) == total_wgs and total_wgs > 0:
                cfg["wg_filter"] = []
            else:
                cfg["wg_filter"] = list(checked)
        else:
            cfg["wg_filter"] = []

        cfg["show_utc_in_bar"] = self.chk_show_utc.isChecked()
        cfg["show_upcoming_when_idle"] = self.chk_show_upcoming.isChecked()
        self.service.settings.save_clock_config(cfg)
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=True)
        self.settings_changed.emit()

    def _update_reference_clocks(self, snapshot: WorldClockSnapshot) -> None:
        self.laptop_time_lbl.setText(snapshot.laptop_dt.strftime("%H:%M:%S"))
        self.laptop_sub_lbl.setText(
            f"{snapshot.laptop_dt:%a %d %b %Y}  ·  {format_utc_offset(snapshot.laptop_offset_minutes)}"
        )
        self.utc_time_lbl.setText(f"{snapshot.utc_dt:%H:%M:%S} UTC")
        self.utc_sub_lbl.setText(f"{snapshot.utc_dt:%a %d %b %Y}")

    def _clear_venues_layout(self) -> None:
        self._venue_cards.clear()
        while self.venues_layout.count():
            item = self.venues_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _make_card(
        self,
        venue: VenueClockState,
        explicit_dir: Optional[Path] = None,
        is_focused: bool = False,
    ) -> _VenueClockCard:
        card = _VenueClockCard(
            venue=venue,
            explicit_meeting_dir=explicit_dir,
            is_focused_meeting=is_focused,
            parent=self.venues_container,
        )
        card.offset_override_requested.connect(self._on_offset_override)
        card.nudge_requested.connect(self._on_nudge_offset)
        card.web_sync_requested.connect(self._on_web_sync_venue)
        card.reset_requested.connect(self._on_reset_venue)
        self._venue_cards.append((venue.location.lower(), is_focused, card))
        return card

    def _rebuild_venue_cards(
        self,
        snapshot: WorldClockSnapshot,
        focused_state: Optional[VenueClockState],
    ) -> None:
        self._clear_venues_layout()

        if focused_state is not None:
            card = self._make_card(focused_state, explicit_dir=self.focus_meeting_dir, is_focused=True)
            self.venues_layout.addWidget(card)

        if snapshot.primary_venues:
            for v in snapshot.primary_venues:
                if (
                    focused_state is not None
                    and v.location.lower() == focused_state.location.lower()
                    and len(snapshot.primary_venues) == 1
                ):
                    continue
                self.venues_layout.addWidget(self._make_card(v))
        elif focused_state is None:
            empty_lbl = QLabel(
                "No in-person 3GPP meetings are active right now for the selected Working Group filter."
            )
            empty_lbl.setWordWrap(True)
            empty_lbl.setAlignment(Qt.AlignCenter)
            empty_lbl.setStyleSheet(
                f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_BODY}px; padding: 24px; "
                f"background: {C.BG_SURFACE}; border: 1px dashed {C.BORDER_STRONG}; border-radius: {R.LG}px;"
            )
            self.venues_layout.addWidget(empty_lbl)

        if snapshot.other_venues:
            other_hdr = QLabel("<b>Other Active 3GPP Venues (Hidden from Status Bar by WG Filter)</b>")
            other_hdr.setStyleSheet(
                f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED}; margin-top: 8px;"
            )
            self.venues_layout.addWidget(other_hdr)
            for v in snapshot.other_venues:
                self.venues_layout.addWidget(self._make_card(v))

        self.venues_layout.addStretch(1)

    def _update_existing_cards(
        self,
        snapshot: WorldClockSnapshot,
        focused_state: Optional[VenueClockState],
    ) -> None:
        by_loc: Dict[str, VenueClockState] = {}
        for v in snapshot.primary_venues + snapshot.other_venues:
            by_loc[v.location.lower()] = v

        for loc_key, is_focused, card in self._venue_cards:
            if is_focused and focused_state is not None:
                card.update_state(focused_state)
            elif loc_key in by_loc:
                card.update_state(by_loc[loc_key])
