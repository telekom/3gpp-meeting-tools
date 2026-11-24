"""Window restoration preferences, geometry encoding, and serialized session state."""

from __future__ import annotations

import json
import logging
import sys
from typing import Any, Dict, Iterable, List, Optional

from PyQt5.QtCore import QByteArray, QSettings, Qt
from PyQt5.QtWidgets import QApplication, QWidget

from core.ui.shell_state import clamp_window_to_available_screens, create_shell_settings

logger = logging.getLogger(__name__)

RESTORE_MODE_ALWAYS = "always"
RESTORE_MODE_RESTART_ONLY = "restart_only"
RESTORE_MODE_OFF = "off"
VALID_RESTORE_MODES = (
    RESTORE_MODE_ALWAYS,
    RESTORE_MODE_RESTART_ONLY,
    RESTORE_MODE_OFF,
)

_KEY_RESTORE_MODE = "session/restore_mode"
_KEY_PENDING_RESTART = "session/pending_restart_restore"
_KEY_OPEN_WINDOWS = "session/open_windows"


def _to_bool(value: object, default: bool = False) -> bool:
    """Parse flexible QSettings boolean values with fallback."""
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


def get_window_restore_mode(settings: Optional[QSettings] = None) -> str:
    """Return the configured workspace-window restoration mode."""
    cfg = settings or create_shell_settings()
    raw = str(cfg.value(_KEY_RESTORE_MODE, RESTORE_MODE_ALWAYS) or RESTORE_MODE_ALWAYS).strip().lower()
    if raw in VALID_RESTORE_MODES:
        return raw
    return RESTORE_MODE_ALWAYS


def set_window_restore_mode(mode: str, settings: Optional[QSettings] = None) -> str:
    """Persist the workspace-window restoration mode and return the normalized value."""
    normalized = str(mode or RESTORE_MODE_ALWAYS).strip().lower()
    if normalized not in VALID_RESTORE_MODES:
        normalized = RESTORE_MODE_ALWAYS
    cfg = settings or create_shell_settings()
    cfg.setValue(_KEY_RESTORE_MODE, normalized)
    cfg.sync()
    return normalized


def mark_pending_restart_restore(settings: Optional[QSettings] = None) -> None:
    """Mark that the next application launch originates from an in-app restart or update."""
    cfg = settings or create_shell_settings()
    cfg.setValue(_KEY_PENDING_RESTART, True)
    cfg.sync()


def consume_pending_restart_restore(settings: Optional[QSettings] = None) -> bool:
    """Read and clear the one-shot restart-originated launch flag."""
    cfg = settings or create_shell_settings()
    pending = _to_bool(cfg.value(_KEY_PENDING_RESTART), default=False)
    if pending or cfg.contains(_KEY_PENDING_RESTART):
        cfg.remove(_KEY_PENDING_RESTART)
        cfg.sync()
    return pending


def should_restore_windows(
    settings: Optional[QSettings] = None,
    argv: Optional[Iterable[str]] = None,
    modifiers: Optional[Qt.KeyboardModifiers] = None,
) -> bool:
    """
    Determine whether secondary workspace windows should be restored for this launch.

    Always consumes `pending_restart_restore` so stale restart flags never leak into
    later cold launches. Supports `--no-restore` and holding `Shift` as escape hatches.
    """
    cfg = settings or create_shell_settings()
    was_restart = consume_pending_restart_restore(cfg)

    args = list(sys.argv if argv is None else argv)
    if any(str(a).strip().lower() == "--no-restore" for a in args):
        logger.info("⏭️ [SESSION] Skipping window restoration (--no-restore).")
        return False

    if modifiers is None:
        try:
            modifiers = QApplication.queryKeyboardModifiers()
        except Exception:
            modifiers = Qt.NoModifier
    if bool(modifiers & Qt.ShiftModifier):
        logger.info("⏭️ [SESSION] Skipping window restoration (Shift held on startup).")
        return False

    mode = get_window_restore_mode(cfg)
    if mode == RESTORE_MODE_OFF:
        return False
    if mode == RESTORE_MODE_RESTART_ONLY:
        return was_restart
    return True


def encode_window_geometry(widget: QWidget) -> str:
    """Encode a widget's Qt geometry to an ASCII base64 string."""
    try:
        ba = widget.saveGeometry()
        if isinstance(ba, QByteArray) and not ba.isEmpty():
            return bytes(ba.toBase64()).decode("ascii")
    except Exception:
        pass
    return ""


def apply_window_geometry(widget: QWidget, geom_b64: str) -> None:
    """Restore a widget's base64-encoded Qt geometry and clamp to active screens."""
    if not geom_b64:
        return
    try:
        raw = QByteArray.fromBase64(str(geom_b64).encode("ascii"))
        if not raw.isEmpty() and hasattr(widget, "restoreGeometry"):
            widget.restoreGeometry(raw)
            clamp_window_to_available_screens(widget)
    except Exception as exc:
        logger.debug("Could not restore geometry for %s: %s", type(widget).__name__, exc)


def save_open_windows_state(
    descriptors: List[Dict[str, Any]],
    settings: Optional[QSettings] = None,
) -> None:
    """Persist serialized workspace-window descriptors into QSettings."""
    cfg = settings or create_shell_settings()
    try:
        payload = json.dumps(descriptors or [], ensure_ascii=False)
    except Exception as exc:
        logger.warning("⚠️ [SESSION] Failed to serialize open window descriptors: %s", exc)
        return
    cfg.setValue(_KEY_OPEN_WINDOWS, payload)
    cfg.sync()


def load_open_windows_state(settings: Optional[QSettings] = None) -> List[Dict[str, Any]]:
    """Load persisted workspace-window descriptors from QSettings."""
    cfg = settings or create_shell_settings()
    raw = cfg.value(_KEY_OPEN_WINDOWS, "[]")
    if isinstance(raw, list):
        return [d for d in raw if isinstance(d, dict)]
    if not isinstance(raw, str) or not raw.strip():
        return []
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            return [d for d in parsed if isinstance(d, dict)]
    except Exception as exc:
        logger.warning("⚠️ [SESSION] Failed to parse saved window session: %s", exc)
    return []


def _show_restored_widget(widget: QWidget, descriptor: Dict[str, Any]) -> None:
    """Apply saved geometry/maximized state and present the restored window."""
    geom_b64 = str(descriptor.get("geometry") or "")
    apply_window_geometry(widget, geom_b64)
    if descriptor.get("maximized"):
        widget.showMaximized()
    else:
        widget.show()
