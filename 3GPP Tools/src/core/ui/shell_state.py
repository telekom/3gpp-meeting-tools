# --- File: src/core/ui/shell_state.py ---
"""
QSettings persistence for main-shell layout and chrome state (Phase 1.6).

Saves and restores window geometry (safely clamped to active monitors),
navigation rail collapse preference, bottom Console & Queue drawer state
and splitter proportions, and last active workspace tab.
"""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtCore import QByteArray, QRect, QSettings
from PyQt5.QtGui import QGuiApplication
from PyQt5.QtWidgets import QWidget

SETTINGS_ORG = "3GPPDelegateTools"
SETTINGS_APP = "MainShell"


def create_shell_settings() -> QSettings:
    """Return default persistent `QSettings` store for the application shell."""
    return QSettings(SETTINGS_ORG, SETTINGS_APP)


def _to_bool(value: object, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in ("1", "true", "yes", "on"):
        return True
    if text in ("0", "false", "no", "off"):
        return False
    return default


def _to_int(value: object, default: int) -> int:
    try:
        return int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default


def _parse_int_list(value: object) -> List[int]:
    if isinstance(value, (list, tuple)):
        out: List[int] = []
        for item in value:
            try:
                out.append(int(item))
            except (TypeError, ValueError):
                pass
        return out
    if isinstance(value, str) and value.strip():
        out = []
        for part in value.split(","):
            try:
                out.append(int(part.strip()))
            except ValueError:
                pass
        return out
    return []


def clamp_window_to_available_screens(window: QWidget) -> None:
    """
    Ensure `window` rests on an attached screen (preventing off-screen
    windows when an external monitor is disconnected between sessions).
    """
    screens = QGuiApplication.screens()
    if not screens:
        return

    geom = window.geometry()
    # Check whether the window's top-center titlebar area intersects any active screen
    title_probe = QRect(geom.x() + max(0, geom.width() // 4), geom.y(), max(40, geom.width() // 2), 40)
    visible_on_any = any(s.availableGeometry().intersects(title_probe) for s in screens)

    target_screen = QGuiApplication.screenAt(geom.center()) or QGuiApplication.primaryScreen()
    if target_screen is None:
        target_screen = screens[0]
    avail = target_screen.availableGeometry()

    if not visible_on_any:
        new_w = min(geom.width(), avail.width())
        new_h = min(geom.height(), avail.height())
        new_x = avail.x() + max(0, (avail.width() - new_w) // 2)
        new_y = avail.y() + max(0, (avail.height() - new_h) // 2)
        window.setGeometry(new_x, new_y, new_w, new_h)
        return

    # Clamp size if screen resolution shrank even though top-left is visible
    clamped_w = min(geom.width(), avail.width())
    clamped_h = min(geom.height(), avail.height())
    if clamped_w != geom.width() or clamped_h != geom.height():
        window.resize(clamped_w, clamped_h)


def save_shell_state(window: object, settings: Optional[QSettings] = None) -> None:
    """Persist window geometry, rail/drawer states, splitter sizes, and active tab."""
    cfg = settings or create_shell_settings()

    if hasattr(window, "saveGeometry"):
        cfg.setValue("window/geometry", window.saveGeometry())

    nav_rail = getattr(window, "nav_rail", None)
    if nav_rail is not None:
        pref = getattr(nav_rail, "manual_collapsed", None)
        cfg.setValue("nav_rail/manual_collapsed", "auto" if pref is None else ("true" if pref else "false"))

    drawer = getattr(window, "bottom_drawer", None)
    if drawer is not None:
        cfg.setValue("drawer/open", bool(drawer.is_open))
        current_h = drawer.height() if drawer.is_open and drawer.height() >= 120 else drawer.preferred_height
        cfg.setValue("drawer/preferred_height", int(current_h))
        inner = drawer.inner_sizes()
        if inner and sum(inner) > 0:
            cfg.setValue("drawer/inner_splitter", ",".join(str(x) for x in inner))

    tabs = getattr(window, "tabs", None)
    if tabs is not None:
        cfg.setValue("shell/active_tab_index", int(tabs.currentIndex()))

    cfg.sync()


def restore_shell_state(window: object, settings: Optional[QSettings] = None) -> None:
    """Restore persisted layout state cleanly onto `window`."""
    cfg = settings or create_shell_settings()

    geom_val = cfg.value("window/geometry")
    if isinstance(geom_val, QByteArray) and not geom_val.isEmpty() and hasattr(window, "restoreGeometry"):
        window.restoreGeometry(geom_val)

    if isinstance(window, QWidget):
        clamp_window_to_available_screens(window)

    nav_rail = getattr(window, "nav_rail", None)
    if nav_rail is not None:
        raw_pref = cfg.value("nav_rail/manual_collapsed", "auto")
        pref_str = str(raw_pref).strip().lower()
        manual_pref: Optional[bool]
        if pref_str in ("true", "1"):
            manual_pref = True
        elif pref_str in ("false", "0"):
            manual_pref = False
        else:
            manual_pref = None
        nav_rail.restore_manual_preference(manual_pref)
        if hasattr(window, "width"):
            nav_rail.apply_responsive_width(window.width())

    drawer = getattr(window, "bottom_drawer", None)
    splitter = getattr(window, "splitter", None)
    if drawer is not None:
        pref_h = _to_int(cfg.value("drawer/preferred_height"), drawer.preferred_height)
        drawer.preferred_height = pref_h

        inner_sizes = _parse_int_list(cfg.value("drawer/inner_splitter"))
        if inner_sizes:
            drawer.set_inner_sizes(inner_sizes)

        is_open = _to_bool(cfg.value("drawer/open"), default=False)
        drawer.set_open(is_open, emit=False)
        if splitter is not None:
            drawer.apply_splitter_sizes(splitter)

        status_bar = getattr(window, "status_bar", None)
        if status_bar is not None and hasattr(status_bar, "update_drawer_badge"):
            status_bar.update_drawer_badge(drawer.is_open, drawer.queue_count, drawer.unread_errors)

    tabs = getattr(window, "tabs", None)
    if tabs is not None and tabs.count() > 0:
        saved_idx = _to_int(cfg.value("shell/active_tab_index"), 0)
        if 0 <= saved_idx < tabs.count():
            tabs.setCurrentIndex(saved_idx)
            if nav_rail is not None:
                nav_rail.set_active_index(saved_idx)
