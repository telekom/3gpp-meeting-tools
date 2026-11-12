# --- File: src/core/ui/nav_rail.py ---
"""
Collapsible grouped navigation rail for the 3GPP Delegate Tools shell (Phase 1.1).

Replaces the top flat QTabBar visually while keeping the underlying `QTabWidget`
intact for full API compatibility across modules and tests.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.theme import B, C, R, S, T
from core.ui.ui_components import escape_qt_mnemonic


@dataclass(frozen=True)
class NavItemSpec:
    """Declarative entry in the application navigation rail."""

    key: str
    group: str
    label: str
    icon_name: str
    tab_index: int
    shortcut: str = ""
    tooltip: str = ""


DEFAULT_NAV_ITEMS: Sequence[NavItemSpec] = (
    # Authoring & Office
    NavItemSpec("plantuml", "Authoring & Office", "PlantUML", "code", 0, "Ctrl+1", "PlantUML Studio & Sequence Editor"),
    NavItemSpec("visio", "Authoring & Office", "Visio & Office", "layers", 1, "Ctrl+2", "Batch Visio, PowerPoint & Stencils Studio"),
    NavItemSpec("word", "Authoring & Office", "Word Tools", "file-text", 2, "Ctrl+3", "3GPP Word Authoring, CR/pCR, Compare & Markdown"),
    # 3GPP Standards
    NavItemSpec("specs", "3GPP Standards", "Specifications", "book-open", 3, "Ctrl+4", "3GPP Specification Archive & Versions"),
    NavItemSpec("spec_search", "3GPP Standards", "Spec Search", "search", 4, "Ctrl+5", "Full-Text Specification Search & Clause Diff"),
    NavItemSpec("protocols", "3GPP Standards", "Protocols", "network", 5, "Ctrl+6", "NAS, RRC, NGAP & PFCP Protocol Evolution"),
    # Delegate & AI
    NavItemSpec("meetings", "Delegate & AI", "Meetings", "calendar", 6, "Ctrl+7", "3GPP Meetings, TDocs, Agreements & Session Plan"),
    NavItemSpec("work_items", "Delegate & AI", "Work Items", "clipboard-list", 7, "Ctrl+8", "3GPP Work Items & Feature Tracker"),
    NavItemSpec("assistant", "Delegate & AI", "Assistant", "sparkles", 8, "Ctrl+9", "Evidence-Grounded Local 3GPP AI Assistant"),
)


class NavRail(QFrame):
    """
    Grouped, collapsible left navigation rail synchronized with a `QTabWidget`.

    Emits `tab_selected(index)` when a destination is clicked and
    `collapsed_changed(bool)` whenever the rail transitions between expanded
    and icon-only modes.
    """

    tab_selected = pyqtSignal(int)
    collapsed_changed = pyqtSignal(bool)

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        items: Optional[Sequence[NavItemSpec]] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("appNavRail")
        self._items: List[NavItemSpec] = list(items or DEFAULT_NAV_ITEMS)
        self._collapsed: bool = False
        self._manual_collapsed: Optional[bool] = None
        self._responsive_compact: bool = False
        self._active_index: int = 0

        self._buttons_by_index: Dict[int, QPushButton] = {}
        self._group_headers: List[QLabel] = []
        self._group_dividers: List[QFrame] = []
        self._footer_buttons: List[QPushButton] = []

        self._build_ui()
        self._apply_styles()
        self._apply_collapsed_state(False, emit=False)

    def _build_ui(self) -> None:
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        root = QVBoxLayout(self)
        root.setContentsMargins(S.SM, S.MD, S.SM, S.SM)
        root.setSpacing(S.SM)

        # --- Top Brand Header & Collapse Toggle ---
        header_row = QHBoxLayout()
        header_row.setContentsMargins(S.XS, 0, S.XS, 0)
        header_row.setSpacing(S.SM)

        self.brand_badge = QLabel("3GPP")
        self.brand_badge.setObjectName("navRailBrandBadge")
        self.brand_badge.setAlignment(Qt.AlignCenter)
        header_row.addWidget(self.brand_badge)

        self.brand_title = QLabel("Delegate Tools")
        self.brand_title.setObjectName("navRailBrandTitle")
        header_row.addWidget(self.brand_title, 1)

        self.toggle_btn = QPushButton()
        self.toggle_btn.setObjectName("navRailToggleBtn")
        self.toggle_btn.setCursor(Qt.PointingHandCursor)
        self.toggle_btn.setIcon(lucide_icon("panel-left-close", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=16))
        self.toggle_btn.setIconSize(QSize(16, 16))
        self.toggle_btn.setToolTip("Collapse sidebar")
        self.toggle_btn.setAccessibleName("Toggle Navigation Rail")
        self.toggle_btn.clicked.connect(self.toggle_collapsed)
        header_row.addWidget(self.toggle_btn)

        root.addLayout(header_row)

        top_sep = QFrame()
        top_sep.setObjectName("navRailSeparator")
        top_sep.setFrameShape(QFrame.HLine)
        root.addWidget(top_sep)

        # --- Scrollable Grouped Destinations ---
        scroll = QScrollArea()
        scroll.setObjectName("navRailScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        nav_container = QWidget()
        nav_container.setObjectName("navRailContainer")
        nav_layout = QVBoxLayout(nav_container)
        nav_layout.setContentsMargins(0, S.XS, 0, S.XS)
        nav_layout.setSpacing(2)

        self._btn_group = QButtonGroup(self)
        self._btn_group.setExclusive(True)

        current_group: Optional[str] = None
        for spec in self._items:
            if spec.group != current_group:
                if current_group is not None:
                    div = QFrame()
                    div.setObjectName("navRailGroupDivider")
                    div.setFrameShape(QFrame.HLine)
                    div.setVisible(False)
                    nav_layout.addWidget(div)
                    self._group_dividers.append(div)

                hdr = QLabel(spec.group.upper())
                hdr.setObjectName("navRailGroupLabel")
                if current_group is not None:
                    hdr.setContentsMargins(S.SM, S.MD, S.SM, S.XS)
                else:
                    hdr.setContentsMargins(S.SM, S.XS, S.SM, S.XS)
                nav_layout.addWidget(hdr)
                self._group_headers.append(hdr)
                current_group = spec.group

            btn = QPushButton(escape_qt_mnemonic(spec.label))
            btn.setObjectName("navRailItemBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setProperty("fullLabel", escape_qt_mnemonic(spec.label))
            btn.setProperty("tabIndex", spec.tab_index)
            btn.setProperty("navKey", spec.key)
            btn.setIcon(
                lucide_icon(
                    spec.icon_name,
                    color=C.TEXT_SECONDARY,
                    active_color=C.ACCENT_PRIMARY,
                    size=16,
                )
            )
            btn.setIconSize(QSize(16, 16))
            tip = spec.tooltip or spec.label
            if spec.shortcut:
                tip = f"{tip} ({spec.shortcut})"
            btn.setToolTip(tip)

            self._btn_group.addButton(btn, spec.tab_index)
            nav_layout.addWidget(btn)
            self._buttons_by_index[spec.tab_index] = btn

        nav_layout.addStretch(1)
        scroll.setWidget(nav_container)
        root.addWidget(scroll, 1)

        # --- Footer Slot (Populated in Phase 1.2 for Help & Settings) ---
        footer_sep = QFrame()
        footer_sep.setObjectName("navRailSeparator")
        footer_sep.setFrameShape(QFrame.HLine)
        root.addWidget(footer_sep)

        self.footer_widget = QWidget()
        self.footer_widget.setObjectName("navRailFooter")
        self.footer_layout = QVBoxLayout(self.footer_widget)
        self.footer_layout.setContentsMargins(0, 0, 0, 0)
        self.footer_layout.setSpacing(2)
        root.addWidget(self.footer_widget)

        self._btn_group.idClicked.connect(self._on_button_clicked)
        self.set_active_index(0)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QFrame#appNavRail {{
                background-color: {C.BG_SURFACE};
                border-right: 1px solid {C.BORDER_SUBTLE};
            }}
            QScrollArea#navRailScroll, QWidget#navRailContainer, QWidget#navRailFooter {{
                background-color: transparent;
            }}
            QLabel#navRailBrandBadge {{
                background-color: {C.ACCENT_PRIMARY};
                color: {C.TEXT_ON_ACCENT};
                border-radius: {R.SM}px;
                padding: 3px 6px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 700;
            }}
            QLabel#navRailBrandTitle {{
                color: {C.TEXT_PRIMARY};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_SECTION_TITLE}px;
                font-weight: 700;
            }}
            QPushButton#navRailToggleBtn {{
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: {R.MD}px;
                padding: 4px;
                min-width: 26px;
                min-height: 26px;
            }}
            QPushButton#navRailToggleBtn:hover {{
                background-color: {C.BG_MUTED};
                border-color: {C.BORDER_STRONG};
            }}
            QFrame#navRailSeparator, QFrame#navRailGroupDivider {{
                color: {C.BORDER_SUBTLE};
                background-color: {C.BORDER_SUBTLE};
                max-height: 1px;
                margin: 2px 4px;
            }}
            QLabel#navRailGroupLabel {{
                color: {C.TEXT_FAINT};
                font-family: {T.FONT_FAMILY_UI};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 0.6px;
            }}
            QPushButton#navRailItemBtn, QPushButton#navRailFooterBtn {{
                text-align: left;
                background-color: transparent;
                color: {C.TEXT_BODY};
                border: 1px solid transparent;
                border-left: 3px solid transparent;
                border-radius: {R.MD}px;
                padding: 7px 10px;
                min-height: 18px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 500;
            }}
            QPushButton#navRailItemBtn:hover, QPushButton#navRailFooterBtn:hover {{
                background-color: {C.BG_HOVER};
                color: {C.TEXT_PRIMARY};
            }}
            QPushButton#navRailItemBtn:checked {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
                border-left: 3px solid {C.ACCENT_PRIMARY};
                font-weight: 600;
            }}
            QPushButton#navRailItemBtn:checked:hover {{
                background-color: {C.ACCENT_SOFT_HOVER};
            }}
            QPushButton#navRailFooterBtn::menu-indicator {{
                image: none;
                width: 0px;
            }}
            QPushButton#navRailItemBtn[collapsed="true"], QPushButton#navRailFooterBtn[collapsed="true"] {{
                text-align: center;
                padding: 7px 0px;
                border-left: 1px solid transparent;
            }}
            QPushButton#navRailItemBtn[collapsed="true"]:checked {{
                border: 1px solid {C.ACCENT_RING};
                background-color: {C.ACCENT_SOFT};
            }}
        """)

    def _on_button_clicked(self, tab_index: int) -> None:
        self._active_index = tab_index
        self.tab_selected.emit(tab_index)

    # --- Public API ---

    @property
    def items(self) -> List[NavItemSpec]:
        return list(self._items)

    @property
    def is_collapsed(self) -> bool:
        return self._collapsed

    @property
    def manual_collapsed(self) -> Optional[bool]:
        return self._manual_collapsed

    def active_index(self) -> int:
        return self._active_index

    def set_active_index(self, index: int) -> None:
        """Synchronize highlighted rail button with `QTabWidget.currentIndex()`."""
        btn = self._buttons_by_index.get(index)
        if btn is not None:
            self._active_index = index
            btn.setChecked(True)

    def register_footer_button(self, btn: QPushButton, full_label: str) -> None:
        """Register a footer action button so it collapses/expands cleanly with the rail."""
        btn.setObjectName("navRailFooterBtn")
        btn.setProperty("fullLabel", full_label)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setText("" if self._collapsed else full_label)
        btn.setProperty("collapsed", "true" if self._collapsed else "false")
        self.footer_layout.addWidget(btn)
        self._footer_buttons.append(btn)
        self.style().unpolish(btn)
        self.style().polish(btn)

    def set_collapsed(self, collapsed: bool, manual: bool = False) -> None:
        """Explicitly set collapsed state (optionally recording manual user pin)."""
        if manual:
            self._manual_collapsed = collapsed
        self._apply_collapsed_state(collapsed, emit=True)

    def toggle_collapsed(self) -> None:
        """Toggle between expanded (224 px) and icon-rail (56 px) states."""
        self.set_collapsed(not self._collapsed, manual=True)

    def restore_manual_preference(self, manual_collapsed: Optional[bool]) -> None:
        """Restore persisted manual preference (Phase 1.6 `QSettings`)."""
        self._manual_collapsed = manual_collapsed
        if not self._responsive_compact and manual_collapsed is not None:
            self._apply_collapsed_state(manual_collapsed, emit=False)

    def apply_responsive_width(self, window_width: int) -> None:
        """
        Apply responsive breakpoint hysteresis (`<940 px` collapse, `>=970 px` expand).
        Preserves manual collapse preference on wider viewports.
        """
        if window_width < B.COMPACT_COLLAPSE_PX:
            if not self._responsive_compact:
                self._responsive_compact = True
                self._apply_collapsed_state(True, emit=True)
        elif window_width >= B.COMPACT_EXPAND_PX:
            if self._responsive_compact:
                self._responsive_compact = False
                target = bool(self._manual_collapsed) if self._manual_collapsed is not None else False
                self._apply_collapsed_state(target, emit=True)

    def _apply_collapsed_state(self, collapsed: bool, emit: bool) -> None:
        changed = collapsed != self._collapsed
        self._collapsed = collapsed
        target_width = S.RAIL_COLLAPSED_WIDTH if collapsed else S.RAIL_EXPANDED_WIDTH
        self.setFixedWidth(target_width)

        self.brand_badge.setVisible(not collapsed)
        self.brand_title.setVisible(not collapsed)

        toggle_icon = "panel-left-open" if collapsed else "panel-left-close"
        self.toggle_btn.setIcon(
            lucide_icon(toggle_icon, color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=16)
        )
        self.toggle_btn.setToolTip("Expand sidebar" if collapsed else "Collapse sidebar")

        for hdr in self._group_headers:
            hdr.setVisible(not collapsed)
        for div in self._group_dividers:
            div.setVisible(collapsed)

        for btn in list(self._buttons_by_index.values()) + list(self._footer_buttons):
            full_label = btn.property("fullLabel") or ""
            btn.setText("" if collapsed else full_label)
            btn.setProperty("collapsed", "true" if collapsed else "false")
            self.style().unpolish(btn)
            self.style().polish(btn)

        if emit and changed:
            self.collapsed_changed.emit(collapsed)
