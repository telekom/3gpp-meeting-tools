"""Facade lookup helper and `refresh_btn` busy/unlock state helpers for `tdocs_agenda`."""

import sys

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMessageBox


def _facade_attr(name: str, default):
    """Resolve `QMessageBox` / `QTimer` through legacy controller shim to preserve test patches."""
    facade = sys.modules.get("modules.meetings.ui.tdocs_agenda_controller")
    if facade is not None and hasattr(facade, name):
        return getattr(facade, name)
    return default


def _qmessagebox():
    return _facade_attr("QMessageBox", QMessageBox)


def _qtimer():
    return _facade_attr("QTimer", QTimer)


def _next_refresh_seq(window) -> int:
    """Increment and return a monotonic sequence token on `window` for `refresh_btn` state."""
    cur = getattr(window, "_refresh_btn_seq", 0)
    seq = (cur if isinstance(cur, int) else 0) + 1
    window._refresh_btn_seq = seq
    return seq


def _set_refresh_btn_busy(window, label: str, max_len: int = 36) -> int:
    """Mark `window.refresh_btn` busy with `label`, preserving its default tooltip."""
    seq = _next_refresh_seq(window)
    btn = getattr(window, "refresh_btn", None)
    if btn is None:
        return seq
    if not hasattr(window, "_default_refresh_tooltip"):
        try:
            window._default_refresh_tooltip = btn.toolTip()
        except Exception:
            window._default_refresh_tooltip = ""
    btn.setEnabled(False)
    text = f"⏳ {label}" if not str(label).startswith("⏳") else str(label)
    btn.setText(text[:max_len])
    return seq


def _update_refresh_btn_progress(window, msg: str, max_len: int = 36):
    """Update `window.refresh_btn` text and tooltip during a background download/probe."""
    btn = getattr(window, "refresh_btn", None)
    if btn is None:
        return
    clean = str(msg or "").strip()
    display = f"⏳ {clean}"
    if len(display) > max_len:
        display = display[: max_len - 1].rstrip() + "…"
    btn.setText(display)
    if clean:
        btn.setToolTip(clean)


def _restore_refresh_btn(window, label: str = "🔄 Refresh", reset_after_ms: int = 0) -> int:
    """Re-enable `window.refresh_btn`, restore its tooltip, and optionally reset after delay."""
    seq = _next_refresh_seq(window)
    btn = getattr(window, "refresh_btn", None)
    if btn is None:
        return seq
    btn.setEnabled(True)
    btn.setText(label)
    default_tip = getattr(window, "_default_refresh_tooltip", None)
    if isinstance(default_tip, str) and default_tip:
        btn.setToolTip(default_tip)

    if reset_after_ms > 0:
        def _reset_if_current():
            try:
                if getattr(window, "_refresh_btn_seq", None) == seq and btn.isEnabled():
                    btn.setText("🔄 Refresh")
            except RuntimeError:
                pass

        _qtimer().singleShot(reset_after_ms, _reset_if_current)
    return seq


def _ensure_refresh_btn_unlocked(window, expected_seq: int):
    """Safety net on `QThread.finished` so `refresh_btn` never remains stuck disabled."""
    try:
        if getattr(window, "_refresh_btn_seq", None) == expected_seq:
            _restore_refresh_btn(window)
    except RuntimeError:
        pass
