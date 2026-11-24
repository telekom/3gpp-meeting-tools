# --- File: src/core/ui/status_bar/styles.py ---
"""Status bar QSS helpers and shared constants."""

from core.ui.theme import C, R, T

_MINIMAL_HYSTERESIS_PX = 24


def _status_pill_qss(
    bg: str = "transparent",
    fg: str = C.TEXT_MUTED,
    bd: str = C.BORDER_STRONG,
    hover_bg: str = C.BG_CANVAS,
    hover_bd: str = C.BORDER_HOVER,
    compact: bool = False,
) -> str:
    """Build consistent DPI-friendly status bar pill QSS."""
    pad = "2px 5px" if compact else "2px 8px"
    return (
        f"QPushButton {{"
        f" background-color: {bg}; border: 1px solid {bd}; border-radius: {R.SM}px;"
        f" padding: {pad}; font-family: {T.FONT_FAMILY_UI}; font-size: {T.SIZE_CAPTION}px;"
        f" font-weight: 600; color: {fg};"
        f"}}"
        f"QPushButton:hover {{ background-color: {hover_bg}; border-color: {hover_bd}; }}"
        f"QPushButton::menu-indicator {{ image: none; width: 0px; }}"
    )
