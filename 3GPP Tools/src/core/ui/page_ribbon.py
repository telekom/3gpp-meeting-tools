# --- File: src/core/ui/page_ribbon.py ---
"""
Reusable 2-tier carded page header & filter ribbon for Phase 2 tab consistency.

Tier 1 holds the page title, optional context pill, primary action bar, and
right-aligned utility/settings actions. Tier 2 holds responsive auto-wrapping
search inputs, segmented filter pills, combo boxes, and reset buttons via
`FlowLayout` so narrow laptop panes never clip horizontally.
"""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QAbstractButton, QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from core.ui.flow_layout import FlowLayout
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T


def apply_lucide_icon(
    button: QAbstractButton,
    icon_name: str,
    color: Optional[str] = None,
    active_color: Optional[str] = None,
    size: int = 14,
) -> None:
    """Attach a crisp theme-aware Lucide vector icon to a button or tool button."""
    icon = lucide_icon(
        icon_name,
        color=color or C.TEXT_SECONDARY,
        active_color=active_color or C.ACCENT_PRIMARY,
        size=size,
    )
    button.setIcon(icon)
    button.setIconSize(QSize(size, size))


def make_filter_label(text: str, emphasized: bool = False) -> QLabel:
    """Create a consistent compact uppercase/muted filter field label."""
    lbl = QLabel(text)
    if emphasized:
        lbl.setStyleSheet(
            f"color: {C.TEXT_SECONDARY}; font-size: {T.SIZE_CAPTION}px; font-weight: bold;"
        )
    else:
        lbl.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; font-weight: 600; margin-left: 2px;"
        )
    return lbl


class PageRibbonCard(QFrame):
    """Consistent 2-tier carded ribbon used at the top of Phase 2 workspace tabs."""

    def __init__(
        self,
        title: str = "",
        subtitle: str = "",
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("pageRibbonCard")
        self.setStyleSheet(
            f"""
            QFrame#pageRibbonCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            """
        )

        self._root_layout = QVBoxLayout(self)
        self._root_layout.setContentsMargins(S.MD, S.SM + 2, S.MD, S.SM + 2)
        self._root_layout.setSpacing(S.SM)

        # Tier 1: Title + primary action flow + right utility actions
        self.tier1_widget = QWidget(self)
        self.tier1_widget.setStyleSheet("background: transparent; border: none;")
        self.tier1_outer = QHBoxLayout(self.tier1_widget)
        self.tier1_outer.setContentsMargins(0, 0, 0, 0)
        self.tier1_outer.setSpacing(S.SM)

        if title or subtitle:
            title_box = QVBoxLayout()
            title_box.setContentsMargins(0, 0, S.SM, 0)
            title_box.setSpacing(1)

            self.title_label = QLabel(title)
            self.title_label.setStyleSheet(
                f"color: {C.TEXT_PRIMARY}; font-size: 15px; font-weight: 600;"
            )
            title_box.addWidget(self.title_label)

            self.subtitle_label: Optional[QLabel] = None
            if subtitle:
                self.subtitle_label = QLabel(subtitle)
                self.subtitle_label.setStyleSheet(
                    f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px;"
                )
                title_box.addWidget(self.subtitle_label)

            self.tier1_outer.addLayout(title_box)
        else:
            self.title_label = None
            self.subtitle_label = None

        self.primary_flow = FlowLayout(margin=0, h_spacing=S.XS + 2, v_spacing=S.XS)
        self.primary_container = QWidget(self.tier1_widget)
        self.primary_container.setLayout(self.primary_flow)
        self.primary_container.setStyleSheet("background: transparent; border: none;")
        self.tier1_outer.addWidget(self.primary_container, 1)

        self.utility_layout = QHBoxLayout()
        self.utility_layout.setContentsMargins(0, 0, 0, 0)
        self.utility_layout.setSpacing(S.XS + 2)
        self.tier1_outer.addLayout(self.utility_layout)

        self._root_layout.addWidget(self.tier1_widget)

        # Divider (hidden until Tier 2 widgets are attached)
        self.divider = QFrame(self)
        self.divider.setFixedHeight(1)
        self.divider.setStyleSheet(f"background-color: {C.BORDER_SUBTLE}; border: none;")
        self.divider.hide()
        self._root_layout.addWidget(self.divider)

        # Tier 2: Responsive filter / search / tag flow bar
        self.tier2_widget = QWidget(self)
        self.tier2_widget.setStyleSheet("background: transparent; border: none;")
        self.filter_flow = FlowLayout(self.tier2_widget, margin=0, h_spacing=S.XS + 2, v_spacing=S.XS)
        self.tier2_widget.hide()
        self._root_layout.addWidget(self.tier2_widget)

    def add_primary_widget(self, widget: QWidget) -> None:
        """Add a widget to the left/main flow area of Tier 1."""
        self.primary_flow.addWidget(widget)

    def add_utility_widget(self, widget: QWidget) -> None:
        """Add a widget to the right-aligned utility strip of Tier 1."""
        self.utility_layout.addWidget(widget)

    def add_filter_widget(self, widget: QWidget) -> None:
        """Add a filter control into the wrapping Tier 2 ribbon row."""
        if self.tier2_widget.isHidden():
            self.divider.show()
            self.tier2_widget.show()
        self.filter_flow.addWidget(widget)
