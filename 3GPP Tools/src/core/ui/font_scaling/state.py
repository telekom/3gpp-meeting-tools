"""Shared typography scale state, presets, math helpers, and QSettings persistence."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from PyQt5.QtCore import QObject, QSettings, pyqtSignal
from PyQt5.QtGui import QFont

from core.ui.shell_state import create_shell_settings

logger = logging.getLogger(__name__)

SETTINGS_KEY_FONT_SCALE = "ui/font_scale"
DEFAULT_FONT_SCALE = 1.0
MIN_FONT_SCALE = 0.85
MAX_FONT_SCALE = 1.60


@dataclass(frozen=True)
class FontScalePreset:
    """Named accessibility/readability preset for the UI typography scale."""

    key: str
    label: str
    description: str
    scale: float


FONT_SCALE_PRESETS: Tuple[FontScalePreset, ...] = (
    FontScalePreset(
        key="compact",
        label="Compact (90%)",
        description="~12 px body · Fits more rows & columns on small displays",
        scale=0.90,
    ),
    FontScalePreset(
        key="standard",
        label="Standard (100% · Default)",
        description="13 px body · Balanced default desktop typography",
        scale=1.00,
    ),
    FontScalePreset(
        key="comfortable",
        label="Comfortable (115%)",
        description="~15 px body · Gently enlarged labels, menus & tables",
        scale=1.15,
    ),
    FontScalePreset(
        key="large",
        label="Large (125%)",
        description="~16 px body · High legibility for laptops & reading",
        scale=1.25,
    ),
    FontScalePreset(
        key="xlarge",
        label="Extra Large (140%)",
        description="~18 px body · Maximum readability across all tools",
        scale=1.40,
    ),
)

_FONT_SIZE_RE = re.compile(r"(font-size\s*:\s*)(\d+(?:\.\d+)?)\s*(px|pt)\b", re.IGNORECASE)
_HEIGHT_RE = re.compile(r"((?:min-height|max-height)\s*:\s*)(\d+(?:\.\d+)?)\s*px\b", re.IGNORECASE)

# Canonical mutable scaling state (always mutated via module attribute references)
_CURRENT_SCALE: float = DEFAULT_FONT_SCALE
_SCALED_TO_BASE: Dict[str, str] = {}
_APP_BASE_QSS: Optional[str] = None
_ORIG_WIDGET_SET_STYLESHEET = None
_ORIG_APP_SET_STYLESHEET = None
_ORIG_LABEL_INIT = None
_ORIG_LABEL_SET_TEXT = None
_ORIG_TEXTEDIT_SET_HTML = None
_ORIG_TEXTEDIT_SET_PLAIN = None
_ORIG_TEXTEDIT_CLEAR = None
_ORIG_TEXTBROWSER_SET_HTML = None
_ORIG_TEXTDOC_SET_HTML = None
_NOTIFIER: Optional["_FontScaleNotifier"] = None


class _FontScaleNotifier(QObject):
    """Emits `scale_changed(float)` whenever the active UI font scale updates."""

    scale_changed = pyqtSignal(float)


def get_font_scale_notifier() -> _FontScaleNotifier:
    """Return the singleton QObject notifier for live font-scale updates."""
    global _NOTIFIER
    if _NOTIFIER is None:
        _NOTIFIER = _FontScaleNotifier()
    return _NOTIFIER


def clamp_font_scale(value: object) -> float:
    """Parse and clamp a font scale factor safely into [0.85, 1.60]."""
    try:
        parsed = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return DEFAULT_FONT_SCALE
    if parsed <= 0.0:
        return DEFAULT_FONT_SCALE
    return round(max(MIN_FONT_SCALE, min(MAX_FONT_SCALE, parsed)), 2)


def get_font_scale() -> float:
    """Return the currently active global UI font scale factor."""
    return _CURRENT_SCALE


def preset_for_scale(scale: Optional[float] = None) -> Optional[FontScalePreset]:
    """Return the matching `FontScalePreset` for `scale` if within 1%, else None."""
    target = clamp_font_scale(get_font_scale() if scale is None else scale)
    for preset in FONT_SCALE_PRESETS:
        if abs(preset.scale - target) < 0.015:
            return preset
    return None


def scale_px(base_px: int, minimum: Optional[int] = None, scale: Optional[float] = None) -> int:
    """Scale an integer pixel dimension using the active (or specified) font scale."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    scaled = int(round(base_px * s))
    floor_val = 1 if minimum is None else int(minimum)
    return max(floor_val, scaled)


def scale_pt(base_pt: int, minimum: int = 7, scale: Optional[float] = None) -> int:
    """Scale an integer point font size using the active (or specified) font scale."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    return max(int(minimum), int(round(base_pt * s)))


def scaled_font(
    base_font: Optional[QFont] = None,
    base_px: int = 12,
    *,
    bold: Optional[bool] = None,
    family: Optional[str] = None,
    minimum: int = 7,
    scale: Optional[float] = None,
) -> QFont:
    """Return a copy of `base_font` sized strictly in scaled logical pixels."""
    f = QFont(base_font) if base_font is not None else QFont()
    if family:
        f.setFamily(family)
    if bold is not None:
        f.setBold(bool(bold))
    f.setPixelSize(scale_px(base_px, minimum=minimum, scale=scale))
    return f


def scale_row_height(base_px: int, minimum: Optional[int] = None, scale: Optional[float] = None) -> int:
    """Scale a table/tree row height cleanly so enlarged text never clips vertically."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    floor_val = max(22, int(round(base_px * 0.90))) if minimum is None else int(minimum)
    if s <= 1.0:
        return max(floor_val, int(round(base_px * max(0.92, s))))
    return max(base_px, int(round(base_px * s)))


def load_font_scale(settings: Optional[QSettings] = None) -> float:
    """Load persisted UI font scale from per-user QSettings."""
    cfg = settings or create_shell_settings()
    raw = cfg.value(SETTINGS_KEY_FONT_SCALE, DEFAULT_FONT_SCALE)
    return clamp_font_scale(raw)


def save_font_scale(scale: float, settings: Optional[QSettings] = None) -> float:
    """Persist UI font scale into per-user QSettings and return the clamped value."""
    clamped = clamp_font_scale(scale)
    cfg = settings or create_shell_settings()
    cfg.setValue(SETTINGS_KEY_FONT_SCALE, f"{clamped:.2f}")
    cfg.sync()
    return clamped
