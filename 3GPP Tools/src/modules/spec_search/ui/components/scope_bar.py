"""Minimal Mode top active specification scope badge bar for SpecSearchTab."""

from __future__ import annotations

from typing import Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C, R, T
from modules.spec_search.core.spec_search_db import extract_major_release


class ActiveSpecsScopeBar(QFrame):
    """
    Compact Minimal-Mode top strip displaying currently included specs as clickable solo badges,
    with 1-click Full Mode expansion, scope presets, and Import actions.
    """

    MAX_VISIBLE_BADGES = 12

    expand_full_mode_requested = pyqtSignal()
    solo_spec_toggled = pyqtSignal(str)
    preset_requested = pyqtSignal(str)
    import_smart_requested = pyqtSignal()
    import_browse_requested = pyqtSignal()
    import_local_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("activeSpecsScopeBar")
        self._spec_buttons: Dict[str, QPushButton] = {}
        self.btn_clear_solo: Optional[QPushButton] = None
        self.btn_overflow: Optional[QPushButton] = None
        self._setup_ui()

    def _setup_ui(self) -> None:
        self.setStyleSheet(
            f"""
            QFrame#activeSpecsScopeBar {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.MD}px;
            }}
            """
        )

        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 5, 8, 5)
        outer.setSpacing(4)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        self.btn_expand_specs = QPushButton("📚 Specs (0/0) ▾", self)
        self.btn_expand_specs.setCursor(Qt.PointingHandCursor)
        self.btn_expand_specs.setMinimumHeight(scale_px(24, minimum=22))
        self.btn_expand_specs.setToolTip("Expand Full Mode specification selector (Ctrl+B)")
        self.btn_expand_specs.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                padding: 2px 8px;
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.ACCENT_SKY};
                color: {C.ACCENT_SKY_DARK};
            }}
            """
        )
        self.btn_expand_specs.clicked.connect(self.expand_full_mode_requested.emit)
        row.addWidget(self.btn_expand_specs)

        self.badges_host = QWidget(self)
        self.badges_host.setStyleSheet(".QWidget { background: transparent; border: none; }")
        self._badges_layout = FlowLayout(self.badges_host, margin=0, h_spacing=4, v_spacing=3)
        row.addWidget(self.badges_host, stretch=1)

        self.btn_presets = QPushButton("⚡ Scope ▾", self)
        self.btn_presets.setCursor(Qt.PointingHandCursor)
        self.btn_presets.setMinimumHeight(scale_px(24, minimum=22))
        self.btn_presets.setToolTip("Quickly adjust included specification versions")
        self.btn_presets.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {C.SKY_SURFACE};
                color: {C.ACCENT_SKY_DARK};
                border: 1px solid {C.SKY_BORDER};
                border-radius: {R.SM}px;
                padding: 2px 8px;
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_SOFT};
                border-color: {C.ACCENT_SKY};
            }}
            QPushButton::menu-indicator {{
                image: none;
                width: 0px;
            }}
            """
        )
        presets_menu = QMenu(self)
        presets_menu.addAction("Select All Releases").triggered.connect(lambda: self.preset_requested.emit("all"))
        presets_menu.addAction("Deselect All").triggered.connect(lambda: self.preset_requested.emit("none"))
        presets_menu.addSeparator()
        presets_menu.addAction("⭐ Latest per Release").triggered.connect(lambda: self.preset_requested.emit("latest"))
        presets_menu.addAction("Rel-18+ Only").triggered.connect(lambda: self.preset_requested.emit("rel18"))
        self.btn_presets.setMenu(presets_menu)
        row.addWidget(self.btn_presets)

        self.btn_import = QPushButton("📥 Import ▾", self)
        self.btn_import.setCursor(Qt.PointingHandCursor)
        self.btn_import.setMinimumHeight(scale_px(24, minimum=22))
        self.btn_import.setToolTip("Import specifications from DB or local Word .docx")
        self.btn_import.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                padding: 2px 8px;
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton::menu-indicator {{
                image: none;
                width: 0px;
            }}
            """
        )
        import_menu = QMenu(self)
        import_menu.addAction("⭐ Smart Import Specification (Latest per Release)...").triggered.connect(
            self.import_smart_requested.emit
        )
        import_menu.addSeparator()
        import_menu.addAction("📥 Browse All Specs && Releases...").triggered.connect(
            self.import_browse_requested.emit
        )
        import_menu.addAction("📁 Import Local Word Document (.docx)...").triggered.connect(
            self.import_local_requested.emit
        )
        self.btn_import.setMenu(import_menu)
        row.addWidget(self.btn_import)

        outer.addLayout(row)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(
            f"""
            QProgressBar {{
                border: none;
                background-color: {C.BORDER_SUBTLE};
                border-radius: 2px;
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_SKY};
                border-radius: 2px;
            }}
            """
        )
        self.progress_bar.setVisible(False)
        outer.addWidget(self.progress_bar)

    def _badge_stylesheet(self, dimmed: bool = False) -> str:
        if dimmed:
            base_bg, base_fg, base_border = C.BG_SURFACE, C.TEXT_FAINT, C.BORDER_SUBTLE
            hover_bg, hover_fg, hover_border = C.BG_CANVAS, C.TEXT_BODY, C.BORDER_STRONG
        else:
            base_bg, base_fg, base_border = C.SKY_SURFACE, C.ACCENT_SKY_DARK, C.SKY_BORDER
            hover_bg, hover_fg, hover_border = C.SKY_SOFT, C.ACCENT_SKY_DARK, C.ACCENT_SKY

        return f"""
            QPushButton {{
                background-color: {base_bg};
                color: {base_fg};
                border: 1px solid {base_border};
                border-radius: 10px;
                padding: 1px 8px;
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {hover_bg};
                color: {hover_fg};
                border-color: {hover_border};
            }}
            QPushButton:checked {{
                background-color: {C.ACCENT_SKY};
                color: {C.BG_CARD};
                border-color: {C.ACCENT_SKY};
            }}
            QPushButton:checked:hover {{
                background-color: {C.ACCENT_SKY_DARK};
                color: {C.BG_CARD};
                border-color: {C.ACCENT_SKY_DARK};
            }}
        """

    def update_scope(
        self,
        checked_by_spec: Dict[str, List[str]],
        total_indexed_specs: int,
        spec_titles: Optional[Dict[str, str]] = None,
        solo_spec: Optional[str] = None,
    ) -> None:
        """Rebuilds active specification badge pills and updates selection counts."""
        self._badges_layout.clear()
        self._spec_buttons.clear()
        self.btn_clear_solo = None
        self.btn_overflow = None

        active_specs = sorted(checked_by_spec.keys())
        self.btn_expand_specs.setText(f"📚 Specs ({len(active_specs)}/{total_indexed_specs}) ▾")

        if not active_specs:
            empty_lbl = QLabel("No specs selected — click Specs ▾ or Scope ▾", self.badges_host)
            empty_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px;")
            self._badges_layout.addWidget(empty_lbl)
            return

        visible_specs = active_specs[: self.MAX_VISIBLE_BADGES]
        for spec_num in visible_specs:
            vers = checked_by_spec.get(spec_num, [])
            if len(vers) == 1:
                rel = extract_major_release(vers[0])
                short_rel = rel.replace("Rel-", "R") if rel.startswith("Rel-") else rel
                label = f"TS {spec_num} · {short_rel}"
            else:
                label = f"TS {spec_num} ({len(vers)})"

            btn = QPushButton(label, self.badges_host)
            btn.setCheckable(True)
            is_solo = solo_spec == spec_num
            btn.setChecked(is_solo)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setMinimumHeight(scale_px(21, minimum=19))
            btn.setStyleSheet(self._badge_stylesheet(dimmed=bool(solo_spec and not is_solo)))

            title = (spec_titles or {}).get(spec_num, "")
            ver_preview = ", ".join(f"v{str(v).lstrip('vV')}" for v in vers)
            tip_parts = [f"TS {spec_num}: {title}" if title else f"TS {spec_num}"]
            if ver_preview:
                tip_parts.append(f"Selected: {ver_preview}")
            tip_parts.append("Click to solo this specification (click again to show all)")
            btn.setToolTip("\n".join(tip_parts))

            btn.clicked.connect(lambda checked, s=spec_num: self._on_badge_clicked(s, checked))
            self._spec_buttons[spec_num] = btn
            self._badges_layout.addWidget(btn)

        overflow_count = len(active_specs) - len(visible_specs)
        if overflow_count > 0:
            overflow_btn = QPushButton(f"+{overflow_count} more", self.badges_host)
            overflow_btn.setCursor(Qt.PointingHandCursor)
            overflow_btn.setMinimumHeight(scale_px(21, minimum=19))
            overflow_btn.setToolTip("Expand Full Mode to view and edit all selected specifications")
            overflow_btn.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {C.BG_CANVAS};
                    color: {C.TEXT_SECONDARY};
                    border: 1px dashed {C.BORDER_STRONG};
                    border-radius: 10px;
                    padding: 1px 8px;
                    font-size: {T.SIZE_CAPTION}px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    background-color: {C.BORDER_SUBTLE};
                    color: {C.TEXT_PRIMARY};
                }}
                """
            )
            overflow_btn.clicked.connect(self.expand_full_mode_requested.emit)
            self.btn_overflow = overflow_btn
            self._badges_layout.addWidget(overflow_btn)

        if solo_spec:
            clear_btn = QPushButton("✕ Show All", self.badges_host)
            clear_btn.setCursor(Qt.PointingHandCursor)
            clear_btn.setMinimumHeight(scale_px(21, minimum=19))
            clear_btn.setToolTip("Clear solo filter and search across all checked specifications")
            clear_btn.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {C.WARNING_SOFT_ALT};
                    color: {C.WARNING_DARK};
                    border: 1px solid {C.WARNING_BORDER};
                    border-radius: 10px;
                    padding: 1px 8px;
                    font-size: {T.SIZE_CAPTION}px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    background-color: {C.WARNING_BORDER};
                }}
                """
            )
            clear_btn.clicked.connect(lambda: self.solo_spec_toggled.emit(""))
            self.btn_clear_solo = clear_btn
            self._badges_layout.addWidget(clear_btn)

    def _on_badge_clicked(self, spec_num: str, checked: bool) -> None:
        self.solo_spec_toggled.emit(spec_num if checked else "")

    def set_progress(self, value: int, visible: bool, indeterminate: bool = False) -> None:
        """Mirrors batch ingestion or maintenance progress inside the Minimal Mode bar."""
        self.progress_bar.setVisible(visible)
        if indeterminate:
            self.progress_bar.setRange(0, 0)
        else:
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(value)
