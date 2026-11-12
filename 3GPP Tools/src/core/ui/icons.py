# --- File: src/core/ui/icons.py ---
"""
Crisp vector SVG icon renderer for 3GPP Delegate Tools.

Provides Lucide-style MIT 24x24 stroke SVG icons rendered at runtime via QtSvg
(`QSvgRenderer`) with theme color injection (`core.ui.theme.C`) and HiDPI
scaling support. No external icon assets or runtime network fetches are needed.
"""

from __future__ import annotations

from typing import Dict, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPainter, QPixmap
from PyQt5.QtSvg import QSvgRenderer

from core.ui.theme import C

# Embedded Lucide-style 24x24 stroke primitives (inner SVG body only).
_LUCIDE_PATHS: Dict[str, str] = {
    "code": (
        '<polyline points="16 18 22 12 16 6"/>'
        '<polyline points="8 6 2 12 8 18"/>'
    ),
    "layers": (
        '<polygon points="12 2 2 7 12 12 22 7 12 2"/>'
        '<polyline points="2 17 12 22 22 17"/>'
        '<polyline points="2 12 12 17 22 12"/>'
    ),
    "file-text": (
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
        '<line x1="16" y1="13" x2="8" y2="13"/>'
        '<line x1="16" y1="17" x2="8" y2="17"/>'
        '<polyline points="10 9 9 9 8 9"/>'
    ),
    "book-open": (
        '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/>'
        '<path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>'
    ),
    "search": (
        '<circle cx="11" cy="11" r="8"/>'
        '<line x1="21" y1="21" x2="16.65" y2="16.65"/>'
    ),
    "network": (
        '<rect x="16" y="16" width="6" height="6" rx="1"/>'
        '<rect x="2" y="16" width="6" height="6" rx="1"/>'
        '<rect x="9" y="2" width="6" height="6" rx="1"/>'
        '<path d="M5 16v-3a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v3"/>'
        '<path d="M12 12V8"/>'
    ),
    "calendar": (
        '<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>'
        '<line x1="16" y1="2" x2="16" y2="6"/>'
        '<line x1="8" y1="2" x2="8" y2="6"/>'
        '<line x1="3" y1="10" x2="21" y2="10"/>'
    ),
    "clipboard-list": (
        '<rect x="8" y="2" width="8" height="4" rx="1" ry="1"/>'
        '<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>'
        '<path d="M12 11h4"/>'
        '<path d="M12 16h4"/>'
        '<path d="M8 11h.01"/>'
        '<path d="M8 16h.01"/>'
    ),
    "sparkles": (
        '<path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"/>'
        '<path d="M5 3v4"/>'
        '<path d="M19 17v4"/>'
        '<path d="M3 5h4"/>'
        '<path d="M17 19h4"/>'
    ),
    "panel-left-close": (
        '<rect x="3" y="3" width="18" height="18" rx="2"/>'
        '<path d="M9 3v18"/>'
        '<path d="m16 15-3-3 3-3"/>'
    ),
    "panel-left-open": (
        '<rect x="3" y="3" width="18" height="18" rx="2"/>'
        '<path d="M9 3v18"/>'
        '<path d="m14 9 3 3-3 3"/>'
    ),
    "help-circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>'
        '<line x1="12" y1="17" x2="12.01" y2="17"/>'
    ),
    "settings": (
        '<circle cx="12" cy="12" r="3"/>'
        '<path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>'
    ),
    "terminal": (
        '<polyline points="4 17 10 11 4 5"/>'
        '<line x1="12" y1="19" x2="20" y2="19"/>'
    ),
    "chevron-up": '<polyline points="18 15 12 9 6 15"/>',
    "chevron-down": '<polyline points="6 9 12 15 18 9"/>',
    "chevron-right": '<polyline points="9 18 15 12 9 6"/>',
    "check-circle": (
        '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>'
        '<polyline points="22 4 12 14.01 9 11.01"/>'
    ),
    "alert-triangle": (
        '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/>'
        '<line x1="12" y1="9" x2="12" y2="13"/>'
        '<line x1="12" y1="17" x2="12.01" y2="17"/>'
    ),
    "alert-circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="12" y1="8" x2="12" y2="12"/>'
        '<line x1="12" y1="16" x2="12.01" y2="16"/>'
    ),
    "info": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="12" y1="16" x2="12" y2="12"/>'
        '<line x1="12" y1="8" x2="12.01" y2="8"/>'
    ),
    "x": (
        '<line x1="18" y1="6" x2="6" y2="18"/>'
        '<line x1="6" y1="6" x2="18" y2="18"/>'
    ),
    "folder-open": (
        '<path d="m6 14 1.45-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.55 6a2 2 0 0 1-1.94 1.5H4a2 2 0 0 1-2-2V5c0-1.1.9-2 2-2h3.93a2 2 0 0 1 1.66.9l.82 1.2a2 2 0 0 0 1.66.9H18a2 2 0 0 1 2 2v2"/>'
    ),
    "external-link": (
        '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>'
        '<polyline points="15 3 21 3 21 9"/>'
        '<line x1="10" y1="14" x2="21" y2="3"/>'
    ),
    "wrench": (
        '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>'
    ),
}

_PIXMAP_CACHE: Dict[Tuple[str, str, int, int], QPixmap] = {}


def get_available_icon_names() -> Tuple[str, ...]:
    """Return sorted tuple of all built-in Lucide icon names."""
    return tuple(sorted(_LUCIDE_PATHS.keys()))


def build_svg_markup(
    name: str,
    color: Optional[str] = None,
    stroke_width: float = 1.9,
) -> str:
    """Return full SVG XML markup for a named Lucide icon."""
    body = _LUCIDE_PATHS.get(name, _LUCIDE_PATHS["circle"] if "circle" in _LUCIDE_PATHS else _LUCIDE_PATHS["info"])
    stroke = color or C.TEXT_BODY
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" '
        f'viewBox="0 0 24 24" fill="none" stroke="{stroke}" '
        f'stroke-width="{stroke_width}" stroke-linecap="round" stroke-linejoin="round">'
        f"{body}</svg>"
    )


def render_svg_pixmap(
    name: str,
    color: Optional[str] = None,
    size: int = 16,
    scale: int = 2,
) -> QPixmap:
    """Render a crisp HiDPI `QPixmap` for a built-in Lucide icon."""
    resolved_color = color or C.TEXT_BODY
    cache_key = (name, resolved_color, size, scale)
    cached = _PIXMAP_CACHE.get(cache_key)
    if cached is not None and not cached.isNull():
        return cached

    px_size = max(1, int(size * scale))
    pixmap = QPixmap(px_size, px_size)
    pixmap.fill(Qt.transparent)

    svg_bytes = build_svg_markup(name, color=resolved_color).encode("utf-8")
    renderer = QSvgRenderer(svg_bytes)
    if renderer.isValid():
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)
        renderer.render(painter)
        painter.end()

    pixmap.setDevicePixelRatio(float(scale))
    _PIXMAP_CACHE[cache_key] = pixmap
    return pixmap


def lucide_icon(
    name: str,
    color: Optional[str] = None,
    active_color: Optional[str] = None,
    size: int = 16,
) -> QIcon:
    """Create a multi-state `QIcon` (normal + checked/selected/active)."""
    normal_col = color or C.TEXT_SECONDARY
    selected_col = active_color or C.ACCENT_PRIMARY

    icon = QIcon()
    normal_pm = render_svg_pixmap(name, color=normal_col, size=size)
    active_pm = render_svg_pixmap(name, color=selected_col, size=size)

    icon.addPixmap(normal_pm, QIcon.Normal, QIcon.Off)
    icon.addPixmap(active_pm, QIcon.Normal, QIcon.On)
    icon.addPixmap(active_pm, QIcon.Selected, QIcon.Off)
    icon.addPixmap(active_pm, QIcon.Selected, QIcon.On)
    icon.addPixmap(active_pm, QIcon.Active, QIcon.On)
    return icon
