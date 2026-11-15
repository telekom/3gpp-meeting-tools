# --- File: src/core/ui/display_settings_dialog.py ---
"""
Display & Text Size settings dialog for configuring global UI typography scaling.
"""

from __future__ import annotations

from typing import Dict, Optional

from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QSlider,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import (
    DEFAULT_FONT_SCALE,
    FONT_SCALE_PRESETS,
    MAX_FONT_SCALE,
    MIN_FONT_SCALE,
    apply_font_scale,
    clamp_font_scale,
    get_font_scale,
    preset_for_scale,
    scale_px,
)
from core.ui.page_ribbon import apply_lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY


class DisplaySettingsDialog(QDialog):
    """Fluent configuration dialog for per-user text size presets and live preview."""

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        settings: Optional[QSettings] = None,
    ) -> None:
        super().__init__(parent)
        self._settings = settings
        self._initial_scale = get_font_scale()
        self._selected_scale = self._initial_scale
        self._syncing_controls = False
        self._preset_radios: Dict[str, QRadioButton] = {}
        self._preset_cards: Dict[str, QFrame] = {}

        self.setWindowTitle("Display & Text Size")
        self.setMinimumWidth(scale_px(560))
        self.setModal(True)

        self._build_ui()
        self._apply_styles()
        self._sync_ui_from_scale(self._selected_scale)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
        root.setSpacing(S.MD)

        # Header
        header_box = QVBoxLayout()
        header_box.setSpacing(S.XS)
        self.title_label = QLabel("Display & Text Size")
        self.title_label.setObjectName("displayTitle")
        header_box.addWidget(self.title_label)

        self.subtitle_label = QLabel(
            "Adjust text size across tables, toolbars, menus, and panels. "
            "Saved per user on this machine."
        )
        self.subtitle_label.setWordWrap(True)
        self.subtitle_label.setObjectName("displaySubtitle")
        header_box.addWidget(self.subtitle_label)
        root.addLayout(header_box)

        # Preset cards
        presets_frame = QFrame(self)
        presets_frame.setObjectName("displaySectionCard")
        presets_layout = QVBoxLayout(presets_frame)
        presets_layout.setContentsMargins(S.MD, S.MD, S.MD, S.MD)
        presets_layout.setSpacing(S.SM)

        sec_lbl = QLabel("Text Size Presets")
        sec_lbl.setObjectName("displaySectionLabel")
        presets_layout.addWidget(sec_lbl)

        self.preset_group = QButtonGroup(self)
        for preset in FONT_SCALE_PRESETS:
            row_card = QFrame(presets_frame)
            row_card.setObjectName("presetRow")
            row_card.setProperty("selected", False)
            row_card.setCursor(Qt.PointingHandCursor)

            row_layout = QHBoxLayout(row_card)
            row_layout.setContentsMargins(S.MD, S.SM, S.MD, S.SM)
            row_layout.setSpacing(S.SM)

            radio = QRadioButton(preset.label, row_card)
            radio.setObjectName("presetRadio")
            self.preset_group.addButton(radio)
            radio.toggled.connect(
                lambda checked, sc=preset.scale: self._on_preset_toggled(checked, sc)
            )
            row_layout.addWidget(radio)

            desc_lbl = QLabel(preset.description, row_card)
            desc_lbl.setObjectName("presetDesc")
            row_layout.addWidget(desc_lbl, 1)

            row_card.mousePressEvent = lambda ev, r=radio: r.setChecked(True)  # type: ignore[assignment]
            presets_layout.addWidget(row_card)

            self._preset_radios[preset.key] = radio
            self._preset_cards[preset.key] = row_card

        # Custom percentage slider
        slider_row = QHBoxLayout()
        slider_row.setContentsMargins(0, S.XS, 0, 0)
        slider_row.setSpacing(S.SM)

        slider_title = QLabel("Fine-tune:")
        slider_title.setObjectName("displaySectionLabel")
        slider_row.addWidget(slider_title)

        self.scale_slider = QSlider(Qt.Horizontal, presets_frame)
        self.scale_slider.setMinimum(int(round(MIN_FONT_SCALE * 100)))
        self.scale_slider.setMaximum(int(round(MAX_FONT_SCALE * 100)))
        self.scale_slider.setSingleStep(5)
        self.scale_slider.setPageStep(5)
        self.scale_slider.setTickPosition(QSlider.TicksBelow)
        self.scale_slider.setTickInterval(10)
        self.scale_slider.valueChanged.connect(self._on_slider_changed)
        slider_row.addWidget(self.scale_slider, 1)

        self.pct_badge = QLabel("100%")
        self.pct_badge.setObjectName("scalePctBadge")
        self.pct_badge.setAlignment(Qt.AlignCenter)
        self.pct_badge.setMinimumWidth(58)
        slider_row.addWidget(self.pct_badge)
        presets_layout.addLayout(slider_row)

        root.addWidget(presets_frame)

        # Live Preview Card
        preview_card = QFrame(self)
        preview_card.setObjectName("displayPreviewCard")
        pv_layout = QVBoxLayout(preview_card)
        pv_layout.setContentsMargins(S.MD, S.MD, S.MD, S.MD)
        pv_layout.setSpacing(S.SM)

        pv_header = QHBoxLayout()
        pv_caption = QLabel("Live Text Preview")
        pv_caption.setObjectName("displaySectionLabel")
        pv_header.addWidget(pv_caption)
        pv_header.addStretch()
        self.pv_metrics_label = QLabel("")
        self.pv_metrics_label.setObjectName("presetDesc")
        pv_header.addWidget(self.pv_metrics_label)
        pv_layout.addLayout(pv_header)

        sample_row = QFrame(preview_card)
        sample_row.setObjectName("previewSampleRow")
        sr_layout = QHBoxLayout(sample_row)
        sr_layout.setContentsMargins(S.MD, S.SM, S.MD, S.SM)
        sr_layout.setSpacing(S.SM)

        self.pv_badge = QLabel("SA2")
        self.pv_badge.setObjectName("previewWgBadge")
        self.pv_badge.setAlignment(Qt.AlignCenter)
        sr_layout.addWidget(self.pv_badge)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        self.pv_title = QLabel("TS 23.501 · System architecture for the 5G System (5GS)")
        self.pv_title.setObjectName("previewTitleText")
        self.pv_body = QLabel("Clause 5.2.3.2: AMF registration management & UDM subscription retrieval.")
        self.pv_body.setObjectName("previewBodyText")
        self.pv_body.setWordWrap(True)
        text_col.addWidget(self.pv_title)
        text_col.addWidget(self.pv_body)
        sr_layout.addLayout(text_col, 1)

        self.pv_btn = QPushButton("Open Word")
        self.pv_btn.setObjectName("previewActionBtn")
        self.pv_btn.setEnabled(False)
        sr_layout.addWidget(self.pv_btn)
        pv_layout.addWidget(sample_row)

        tip_lbl = QLabel(
            "Tip: In reading panes & editors (PlantUML, Spec Search, Assistant, Emails), "
            "you can also press Ctrl++ / Ctrl+- / Ctrl+0 to zoom text immediately."
        )
        tip_lbl.setWordWrap(True)
        tip_lbl.setObjectName("displayTipLabel")
        pv_layout.addWidget(tip_lbl)

        root.addWidget(preview_card)

        # Footer Buttons
        footer = QHBoxLayout()
        footer.setSpacing(S.SM)

        self.reset_btn = QPushButton("Reset to Default (100%)")
        self.reset_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.reset_btn, "rotate-ccw", size=14)
        self.reset_btn.clicked.connect(self._on_reset_clicked)
        footer.addWidget(self.reset_btn)

        footer.addStretch()

        self.apply_btn = QPushButton("Apply")
        self.apply_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.apply_btn.clicked.connect(self._on_apply_clicked)
        footer.addWidget(self.apply_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.clicked.connect(self.reject)
        footer.addWidget(self.cancel_btn)

        self.ok_btn = QPushButton("OK")
        self.ok_btn.setObjectName("primaryBtn")
        self.ok_btn.setDefault(True)
        self.ok_btn.clicked.connect(self._on_ok_clicked)
        footer.addWidget(self.ok_btn)

        root.addLayout(footer)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {C.BG_CANVAS};
            }}
            QPushButton#primaryBtn {{
                background-color: {C.ACCENT_PRIMARY};
                color: {C.TEXT_ON_ACCENT};
                border: none;
                border-radius: {R.MD}px;
                padding: 6px 16px;
                font-size: {T.SIZE_TABLE}px;
                font-weight: 700;
            }}
            QPushButton#primaryBtn:hover {{
                background-color: {C.ACCENT_HOVER};
            }}
            QLabel#displayTitle {{
                font-size: {T.SIZE_PAGE_TITLE}px;
                font-weight: 700;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#displaySubtitle, QLabel#presetDesc, QLabel#displayTipLabel {{
                font-size: {T.SIZE_CAPTION}px;
                color: {C.TEXT_MUTED};
            }}
            QLabel#displaySectionLabel {{
                font-size: {T.SIZE_TABLE}px;
                font-weight: 700;
                color: {C.TEXT_STRONG};
            }}
            QFrame#displaySectionCard, QFrame#displayPreviewCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QFrame#presetRow {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
            }}
            QFrame#presetRow[selected="true"] {{
                background-color: {C.ACCENT_SOFT};
                border: 1px solid {C.ACCENT_PRIMARY};
            }}
            QRadioButton#presetRadio {{
                font-size: {T.SIZE_BODY}px;
                font-weight: 600;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#scalePctBadge {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
                border: 1px solid {C.ACCENT_RING};
                border-radius: {R.SM}px;
                padding: 3px 8px;
                font-size: {T.SIZE_TABLE}px;
                font-weight: 700;
            }}
            QFrame#previewSampleRow {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
            }}
            QLabel#previewWgBadge {{
                background-color: {C.WG_SA_BG};
                color: {C.WG_SA_TEXT};
                border: 1px solid {C.WG_SA_BORDER};
                border-radius: {R.SM}px;
                padding: 3px 8px;
                font-weight: 700;
            }}
            QLabel#previewTitleText {{
                color: {C.TEXT_PRIMARY};
                font-weight: 700;
            }}
            QLabel#previewBodyText {{
                color: {C.TEXT_BODY};
            }}
            QPushButton#previewActionBtn {{
                background-color: {C.SUCCESS_SOFT};
                color: {C.SUCCESS};
                border: 1px solid {C.SUCCESS_BORDER};
                border-radius: {R.MD}px;
                padding: 5px 12px;
                font-weight: 700;
            }}
        """)

    def _sync_ui_from_scale(self, scale: float) -> None:
        self._syncing_controls = True
        try:
            clamped = clamp_font_scale(scale)
            self._selected_scale = clamped
            pct = int(round(clamped * 100))
            self.scale_slider.setValue(pct)
            self.pct_badge.setText(f"{pct}%")

            matched = preset_for_scale(clamped)
            self.preset_group.setExclusive(False)
            for key, radio in self._preset_radios.items():
                is_match = matched is not None and matched.key == key
                radio.setChecked(is_match)
                card = self._preset_cards[key]
                card.setProperty("selected", is_match)
                card.style().unpolish(card)
                card.style().polish(card)
            self.preset_group.setExclusive(True)

            self._update_preview(clamped)
        finally:
            self._syncing_controls = False

    def _update_preview(self, scale: float) -> None:
        body_px = scale_px(T.SIZE_BODY, scale=scale)
        table_px = scale_px(T.SIZE_TABLE, scale=scale)
        caption_px = scale_px(T.SIZE_CAPTION, scale=scale)
        self.pv_metrics_label.setText(
            f"Body {body_px}px · Table {table_px}px · Caption {caption_px}px"
        )

        f_title = QFont(self.pv_title.font())
        f_title.setPixelSize(body_px)
        f_title.setBold(True)
        self.pv_title.setFont(f_title)

        f_body = QFont(self.pv_body.font())
        f_body.setPixelSize(table_px)
        self.pv_body.setFont(f_body)

        f_badge = QFont(self.pv_badge.font())
        f_badge.setPixelSize(caption_px)
        f_badge.setBold(True)
        self.pv_badge.setFont(f_badge)

        f_btn = QFont(self.pv_btn.font())
        f_btn.setPixelSize(table_px)
        f_btn.setBold(True)
        self.pv_btn.setFont(f_btn)

    def _on_preset_toggled(self, checked: bool, scale: float) -> None:
        if self._syncing_controls or not checked:
            return
        self._sync_ui_from_scale(scale)

    def _on_slider_changed(self, pct: int) -> None:
        if self._syncing_controls:
            return
        self._sync_ui_from_scale(pct / 100.0)

    def _on_reset_clicked(self) -> None:
        self._sync_ui_from_scale(DEFAULT_FONT_SCALE)

    def _apply_current_selection(self) -> float:
        applied = apply_font_scale(
            self._selected_scale,
            app=QApplication.instance(),
            save=True,
            settings=self._settings,
        )
        self._apply_styles()
        self._sync_ui_from_scale(applied)
        return applied

    def _on_apply_clicked(self) -> None:
        self._apply_current_selection()

    def _on_ok_clicked(self) -> None:
        self._apply_current_selection()
        self.accept()

    @property
    def selected_scale(self) -> float:
        return self._selected_scale
