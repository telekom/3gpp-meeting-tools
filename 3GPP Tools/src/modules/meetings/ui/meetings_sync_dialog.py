"""Custom and Deep Sync configuration dialog for the Meetings workspace."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C, R, S, T
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, SPEC_FILTER_COMBO_STYLE
from modules.meetings.core.crawler.sources import MEETING_SOURCES


class MeetingsSyncDialog(QDialog):
    """Dialog allowing users to scope working groups and sync passes explicitly."""

    def __init__(
        self,
        initial_wgs: Optional[list] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Custom & Deep Meetings Sync")
        self.setMinimumWidth(scale_px(540))

        known_wgs = list(MEETING_SOURCES.keys())
        cleaned_initial = [w for w in (initial_wgs or []) if w in MEETING_SOURCES]
        self._start_scoped = 0 < len(cleaned_initial) < len(known_wgs)
        self._initial_wgs = set(cleaned_initial if self._start_scoped else known_wgs)

        self.wg_checkboxes: dict[str, QCheckBox] = {}
        self._setup_ui(known_wgs)

    def _setup_ui(self, known_wgs: list) -> None:
        self.setStyleSheet(
            f"""
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.MD}px;
                margin-top: 10px;
                padding-top: 12px;
                background-color: {C.BG_CARD};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
            """
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
        layout.setSpacing(S.MD)

        intro = QLabel(
            "Choose which Working Groups and synchronization steps to run. "
            "Folders and official DynaReport dates always run before Docs/ scanning."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: {scale_px(T.SIZE_CAPTION)}px;")
        layout.addWidget(intro)

        # 1. Working Group Scope
        scope_group = QGroupBox("1. Working Group Scope", self)
        scope_layout = QVBoxLayout(scope_group)
        scope_layout.setSpacing(S.SM)

        self.scope_btn_group = QButtonGroup(self)
        self.rb_all_wgs = QRadioButton(f"All Working Groups ({len(known_wgs)})", scope_group)
        self.rb_selected_wgs = QRadioButton("Selected Working Groups Only", scope_group)
        self.scope_btn_group.addButton(self.rb_all_wgs)
        self.scope_btn_group.addButton(self.rb_selected_wgs)

        scope_radios = QHBoxLayout()
        scope_radios.addWidget(self.rb_all_wgs)
        scope_radios.addWidget(self.rb_selected_wgs)
        scope_radios.addStretch()
        scope_layout.addLayout(scope_radios)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(S.XS)
        self._quick_scope_btns: list[QPushButton] = []
        for label, pred in (
            ("All", lambda _w: True),
            ("None", lambda _w: False),
            ("SA", lambda w: w.startswith("SA")),
            ("RAN", lambda w: w.startswith("RAN")),
            ("CT", lambda w: w.startswith("CT")),
        ):
            btn = QPushButton(label, scope_group)
            btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            btn.setMinimumHeight(scale_px(24))
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, p=pred: self._apply_wg_preset(p))
            quick_row.addWidget(btn)
            self._quick_scope_btns.append(btn)
        quick_row.addStretch()
        scope_layout.addLayout(quick_row)

        grid = QGridLayout()
        grid.setHorizontalSpacing(S.MD)
        grid.setVerticalSpacing(S.XS)
        for idx, wg in enumerate(known_wgs):
            cb = QCheckBox(wg, scope_group)
            cb.setChecked(wg in self._initial_wgs)
            cb.toggled.connect(self._update_state)
            self.wg_checkboxes[wg] = cb
            grid.addWidget(cb, idx // 4, idx % 4)
        scope_layout.addLayout(grid)
        layout.addWidget(scope_group)

        # 2. Sync Steps
        steps_group = QGroupBox("2. Sync Steps (Execution Order)", self)
        steps_layout = QVBoxLayout(steps_group)
        steps_layout.setSpacing(S.SM)

        self.chk_wg = QCheckBox("Step 1: Discover new FTP meeting folders (~5–8s)", steps_group)
        self.chk_wg.setChecked(True)
        self.chk_wg.setToolTip(
            "Scans 3GPP FTP WG directories for newly created meeting folders."
        )
        self.chk_wg.toggled.connect(self._update_state)
        steps_layout.addWidget(self.chk_wg)

        self.chk_dyna = QCheckBox(
            "Step 2: Fetch official dates, locations && Portal IDs via DynaReport (~3s)",
            steps_group,
        )
        self.chk_dyna.setChecked(True)
        self.chk_dyna.setToolTip(
            "Fetches official meeting dates, locations, Portal MtgIds, and future scheduled meetings."
        )
        self.chk_dyna.toggled.connect(self._update_state)
        steps_layout.addWidget(self.chk_dyna)

        docs_row = QHBoxLayout()
        docs_row.setSpacing(S.SM)
        docs_lbl = QLabel("Step 3: Docs/ TDoc range scan:", steps_group)
        docs_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-weight: 600;")
        docs_row.addWidget(docs_lbl)

        self.docs_mode_combo = QComboBox(steps_group)
        self.docs_mode_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.docs_mode_combo.setMinimumHeight(scale_px(28))
        self.docs_mode_combo.addItem(
            "Smart — Recent (12 months), upcoming & missing (~15–30s)", "smart"
        )
        self.docs_mode_combo.addItem("Skip Docs/ scan (Fastest, ~10s)", "none")
        self.docs_mode_combo.addItem(
            "Full Historical — Re-scan all ~2,600 meetings (~20–30 min)", "all"
        )
        self.docs_mode_combo.currentIndexChanged.connect(self._update_state)
        docs_row.addWidget(self.docs_mode_combo, 1)
        steps_layout.addLayout(docs_row)

        self.full_docs_warning = QFrame(steps_group)
        self.full_docs_warning.setStyleSheet(
            f"""
            QFrame {{
                background-color: {C.WARNING_SOFT};
                border: 1px solid {C.WARNING_BORDER};
                border-radius: {R.SM}px;
            }}
            """
        )
        warn_layout = QHBoxLayout(self.full_docs_warning)
        warn_layout.setContentsMargins(S.SM, S.XS, S.SM, S.XS)
        warn_lbl = QLabel(
            "Full Historical mode scans Docs/ for every archived meeting since 1999 (~20–30 min across all WGs). "
            "You can stop it anytime using the Stop button in the Meetings banner.",
            self.full_docs_warning,
        )
        warn_lbl.setWordWrap(True)
        warn_lbl.setStyleSheet(
            f"color: {C.WARNING_DARK}; font-size: {scale_px(T.SIZE_CAPTION)}px; background: transparent; border: none;"
        )
        warn_layout.addWidget(warn_lbl)
        steps_layout.addWidget(self.full_docs_warning)
        layout.addWidget(steps_group)

        # Actions
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_start = QPushButton("Start Sync", self)
        self.btn_start.setObjectName("primaryBtn")
        self.btn_start.setMinimumHeight(scale_px(30))
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.clicked.connect(self.accept)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cancel.setMinimumHeight(scale_px(30))
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)

        btn_row.addWidget(self.btn_start)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)

        if self._start_scoped:
            self.rb_selected_wgs.setChecked(True)
        else:
            self.rb_all_wgs.setChecked(True)

        self.rb_all_wgs.toggled.connect(self._update_state)
        self.rb_selected_wgs.toggled.connect(self._update_state)
        self._update_state()

    def _apply_wg_preset(self, predicate) -> None:
        self.rb_selected_wgs.setChecked(True)
        for wg, cb in self.wg_checkboxes.items():
            cb.setChecked(bool(predicate(wg)))
        self._update_state()

    def _update_state(self) -> None:
        scoped = self.rb_selected_wgs.isChecked()
        for cb in self.wg_checkboxes.values():
            cb.setEnabled(scoped)
        for btn in self._quick_scope_btns:
            btn.setEnabled(scoped)

        selected_count = (
            len(self.wg_checkboxes)
            if not scoped
            else sum(1 for cb in self.wg_checkboxes.values() if cb.isChecked())
        )
        docs_mode = self.docs_mode_combo.currentData()
        has_step = self.chk_wg.isChecked() or self.chk_dyna.isChecked() or docs_mode != "none"

        self.full_docs_warning.setVisible(docs_mode == "all")
        self.btn_start.setEnabled(selected_count > 0 and has_step)

    def get_sync_options(self) -> dict:
        """Return crawler parameters and human-readable label for the selected sync configuration."""
        target_wgs = (
            []
            if self.rb_all_wgs.isChecked()
            else [wg for wg, cb in self.wg_checkboxes.items() if cb.isChecked()]
        )
        docs_mode = self.docs_mode_combo.currentData() or "smart"
        if target_wgs:
            wg_tag = ",".join(target_wgs[:3]) + (f"+{len(target_wgs) - 3}" if len(target_wgs) > 3 else "")
            label = f"3GPP_Meetings_{wg_tag}"
        else:
            label = "3GPP_Meetings"
        return {
            "sync_wg": self.chk_wg.isChecked(),
            "sync_dyna": self.chk_dyna.isChecked(),
            "docs_mode": docs_mode,
            "target_wgs": target_wgs,
            "label": label,
        }
