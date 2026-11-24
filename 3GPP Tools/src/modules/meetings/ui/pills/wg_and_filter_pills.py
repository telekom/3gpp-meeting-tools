"""Working Group color palette helper, WgTogglePill, and FilterPillButton widgets."""

from __future__ import annotations

from typing import Dict, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QPushButton, QSizePolicy

from core.ui.theme import C


def get_wg_colors(wg: str) -> Dict[str, str]:
    """Returns archetype color schemes based on Working Group family."""
    wg_u = str(wg or "").strip().upper()
    if wg_u.startswith("SA"):
        return {
            "bg_active": C.INFO_SOFT_STRONG,
            "fg_active": C.INFO_DARK,
            "border_active": C.INFO_BORDER_MEDIUM,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.INFO_SOFT,
        }
    elif wg_u.startswith("RAN"):
        return {
            "bg_active": C.VIOLET_SOFT,
            "fg_active": C.VIOLET_DARK,
            "border_active": C.VIOLET_BORDER_STRONG,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.VIOLET_SURFACE,
        }
    else:  # CT or others
        return {
            "bg_active": C.TEAL_SOFT,
            "fg_active": C.TEAL_DARK,
            "border_active": C.TEAL_BORDER,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.TEAL_SURFACE,
        }


class WgTogglePill(QPushButton):
    """
    Checkable pill button for selecting Target Working Groups.
    Displays '✓ WG' when checked and '+ WG' when unchecked with family-specific colors.
    """

    def __init__(self, wg: str, checked: bool = False, parent=None):
        super().__init__(parent)
        self.wg = str(wg).strip().upper()
        self.colors = get_wg_colors(self.wg)
        self.setCheckable(True)
        self.setChecked(checked)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        self._update_appearance()
        self.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        self._update_appearance()

    def setChecked(self, checked: bool):
        super().setChecked(checked)
        self._update_appearance()

    def _update_appearance(self):
        c = self.colors
        if self.isChecked():
            self.setText(f"✓ {self.wg}")
            self.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    border-color: {c['fg_active']};
                }}
            """
            )
        else:
            self.setText(f"+ {self.wg}")
            self.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    color: {C.TEXT_STRONG};
                    border-color: {C.TEXT_FAINT};
                }}
            """
            )


class FilterPillButton(QPushButton):
    """
    Clickable, checkable filter pill with title, count badge, and custom color scheme.
    Used for 1-click filtering by Working Group and Decision Category.
    """

    def __init__(
        self,
        key: str,
        label: str,
        count: Optional[int] = None,
        colors: Optional[Dict[str, str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.key = key
        self.base_label = label
        self._count = count
        self.colors = colors or {
            "bg_active": C.ACCENT_ALT,
            "fg_active": C.BG_CARD,
            "border_active": C.INFO,
            "bg_inactive": C.BG_CARD,
            "fg_inactive": C.TEXT_SECONDARY,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.BG_SURFACE,
        }
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._update_text_and_style()
        self.toggled.connect(self._on_toggled)

    def set_count(self, count: Optional[int]):
        self._count = count
        self._update_text_and_style()

    def _on_toggled(self, checked: bool):
        self._update_text_and_style()

    def _update_text_and_style(self):
        c = self.colors
        cnt_str = f" ({self._count})" if self._count is not None and self._count >= 0 else ""
        self.setText(f"{self.base_label}{cnt_str}")
        if self.isChecked():
            self.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
            """
            )
        else:
            self.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    border-color: {C.TEXT_FAINT};
                    color: {C.TEXT_PRIMARY};
                }}
            """
            )
