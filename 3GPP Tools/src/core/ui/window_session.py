# --- File: src/core/ui/window_session.py ---
"""
Workspace window session persistence and restoration.

Captures open primary workspace windows (TDocsWindow, SessionPlanWindow,
EmailManagerWindow, MeetingEmailInboxDialog, TopicDossierDialog,
AgreementTrackerDialog, ContributionReportDialog, SpecVersionHistoryDialog,
SpecClauseExcerptDialog, SpecClauseDiffDialog, CrPcrDialog) at shutdown and
reopens them with their geometry and key in-window state on startup or restart.
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Set

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


def _make_descriptor(win_type: str, widget: QWidget, state: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "type": win_type,
        "geometry": encode_window_geometry(widget),
        "maximized": bool(widget.isMaximized()),
        "state": state,
    }


def _collect_from_widget(
    widget: QWidget,
    seen_ids: Set[int],
    descriptors: List[Dict[str, Any]],
    type_map: Dict[str, type],
) -> None:
    if widget is None or id(widget) in seen_ids:
        return
    try:
        if not widget.isVisible():
            return
    except RuntimeError:
        return

    for win_type, cls in type_map.items():
        if isinstance(widget, cls):
            seen_ids.add(id(widget))
            try:
                state = widget.get_session_state() if hasattr(widget, "get_session_state") else {}
                if isinstance(state, dict):
                    descriptors.append(_make_descriptor(win_type, widget, state))
            except Exception as exc:
                logger.warning("⚠️ [SESSION] Could not capture state for %s: %s", win_type, exc)
            return


def _resolve_supported_window_types() -> Dict[str, type]:
    """Lazily import primary workspace window classes to avoid circular startup imports."""
    type_map: Dict[str, type] = {}
    candidates = [
        ("tdocs_window", "modules.meetings.ui.tdocs.window.main_window", "TDocsWindow"),
        ("session_plan", "modules.meetings.ui.session_plan.window.main_window", "SessionPlanWindow"),
        ("email_manager", "modules.emails.ui.email_window", "EmailManagerWindow"),
        ("meeting_email_inbox", "modules.emails.ui.meeting_email_inbox_dialog", "MeetingEmailInboxDialog"),
        ("topic_dossier", "modules.meetings.ui.dialogs.topic_dossier_dialog", "TopicDossierDialog"),
        ("agreement_tracker", "modules.meetings.ui.agreement_tracker_dialog", "AgreementTrackerDialog"),
        ("contribution_report", "modules.meetings.ui.contributions.contribution_dialog", "ContributionReportDialog"),
        ("spec_history", "modules.specifications.ui.spec_history_dialog", "SpecVersionHistoryDialog"),
        ("spec_clause_excerpt", "modules.specifications.ui.spec_clause_dialog", "SpecClauseExcerptDialog"),
        ("spec_clause_diff", "modules.spec_search.ui.spec_clause_diff_dialog", "SpecClauseDiffDialog"),
        ("cr_pcr", "modules.word_tools.ui.cr_pcr_dialog", "CrPcrDialog"),
    ]
    for key, mod_name, cls_name in candidates:
        try:
            mod = __import__(mod_name, fromlist=[cls_name])
            cls = getattr(mod, cls_name, None)
            if cls is not None:
                type_map[key] = cls
        except Exception as exc:
            logger.debug("Skipping window class %s.%s during session scan: %s", mod_name, cls_name, exc)
    return type_map


def collect_open_windows_state(main_window: Any) -> List[Dict[str, Any]]:
    """
    Snapshot all visible primary workspace windows in restoration order.

    Parent meeting windows (`TDocsWindow`) are recorded before their dependent child
    windows (`SessionPlanWindow`, `EmailManagerWindow`, `MeetingEmailInboxDialog`)
    so restoration can reattach owners cleanly.
    """
    type_map = _resolve_supported_window_types()
    descriptors: List[Dict[str, Any]] = []
    seen_ids: Set[int] = set()

    # 1. Walk MeetingsTab TDocsWindows first, along with their owned sub-windows
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is not None:
        for tdoc_win in list(getattr(meetings_tab, "tdoc_windows", {}).values()):
            _collect_from_widget(tdoc_win, seen_ids, descriptors, type_map)
            if tdoc_win is not None:
                for child_attr in (
                    "_session_plan_window",
                    "email_window",
                    "_inbox_dialog",
                    "_agreement_tracker_dlg",
                ):
                    _collect_from_widget(
                        getattr(tdoc_win, child_attr, None),
                        seen_ids,
                        descriptors,
                        type_map,
                    )

        for tab_attr in ("topic_dossier_dialog", "agreement_tracker_dialog", "contribution_dialog"):
            _collect_from_widget(
                getattr(meetings_tab, tab_attr, None),
                seen_ids,
                descriptors,
                type_map,
            )

    # 2. Walk remaining visible top-level windows (Specifications, Spec Search, Word Tools, etc.)
    try:
        top_widgets = list(QApplication.topLevelWidgets())
    except Exception:
        top_widgets = []

    for widget in top_widgets:
        if widget is main_window:
            continue
        _collect_from_widget(widget, seen_ids, descriptors, type_map)

    return descriptors


def _show_restored_widget(widget: QWidget, descriptor: Dict[str, Any]) -> None:
    """Apply saved geometry/maximized state and present the restored window."""
    geom_b64 = str(descriptor.get("geometry") or "")
    apply_window_geometry(widget, geom_b64)
    if descriptor.get("maximized"):
        widget.showMaximized()
    else:
        widget.show()


def _find_open_tdocs_window(meetings_tab: Any, meeting_dir: str, mtg_id: str) -> Optional[Any]:
    tdoc_windows = getattr(meetings_tab, "tdoc_windows", {}) or {}
    if mtg_id and mtg_id in tdoc_windows:
        return tdoc_windows[mtg_id]
    target_dir = str(Path(meeting_dir)).lower() if meeting_dir else ""
    if target_dir:
        for win in tdoc_windows.values():
            try:
                if str(Path(win.meeting_dir)).lower() == target_dir:
                    return win
            except Exception:
                pass
    return None


def _restore_tdocs_window(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    filepath = str(state.get("filepath") or "").strip()
    mtg_info = state.get("mtg_info") or {}
    if not filepath or not Path(filepath).is_file() or not isinstance(mtg_info, dict):
        logger.info("⏭️ [SESSION] Skipping TDocsWindow restore; missing file: %s", filepath)
        return None

    mtg_id = str(mtg_info.get("mtg_id") or "").strip()
    win = _find_open_tdocs_window(meetings_tab, "", mtg_id)
    if win is None:
        meetings_tab._open_tdocs_window(mtg_info, filepath, quiet=True)
        win = _find_open_tdocs_window(meetings_tab, "", mtg_id)
    if win is None:
        return None

    _show_restored_widget(win, descriptor)
    if hasattr(win, "restore_session_state"):
        win.restore_session_state(state)
    return win


def _restore_session_plan(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    meeting_dir = str(state.get("meeting_dir") or "").strip()
    mtg_info = state.get("mtg_info") or {}
    mtg_id = str(mtg_info.get("mtg_id") or "").strip() if isinstance(mtg_info, dict) else ""
    tdoc_win = _find_open_tdocs_window(meetings_tab, meeting_dir, mtg_id)
    if tdoc_win is None or not Path(meeting_dir).exists():
        return None

    tdoc_win._open_session_plan_window()
    sp_win = getattr(tdoc_win, "_session_plan_window", None)
    if sp_win is None:
        return None
    _show_restored_widget(sp_win, descriptor)
    if hasattr(sp_win, "restore_session_state"):
        sp_win.restore_session_state(state)
    return sp_win


def _restore_email_manager(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    meeting_dir = str(state.get("meeting_dir") or "").strip()
    tdoc_win = _find_open_tdocs_window(meetings_tab, meeting_dir, "")
    if tdoc_win is None:
        return None
    tdoc_win._open_email_manager()
    em_win = getattr(tdoc_win, "email_window", None)
    if em_win is None:
        return None
    _show_restored_widget(em_win, descriptor)
    if hasattr(em_win, "restore_session_state"):
        em_win.restore_session_state(state)
    return em_win


def _restore_meeting_email_inbox(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    meeting_dir = str(state.get("meeting_dir") or "").strip()
    mtg_id = str(state.get("meeting_id") or "").strip()
    tdoc_win = _find_open_tdocs_window(meetings_tab, meeting_dir, mtg_id)
    if tdoc_win is None:
        return None
    tdoc_win._open_email_inbox()
    dlg = getattr(tdoc_win, "_inbox_dialog", None)
    if dlg is None:
        return None
    _show_restored_widget(dlg, descriptor)
    return dlg


def _restore_topic_dossier(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    from modules.meetings.ui.meetings_actions import open_topic_dossier_studio

    open_topic_dossier_studio(meetings_tab)
    dlg = getattr(meetings_tab, "topic_dossier_dialog", None)
    if dlg is None:
        return None
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_agreement_tracker(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    from modules.meetings.ui.meetings_actions import open_agreement_tracker

    open_agreement_tracker(meetings_tab)
    dlg = getattr(meetings_tab, "agreement_tracker_dialog", None)
    if dlg is None:
        return None
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_contribution_report(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    meetings_tab = getattr(main_window, "meetings_tab", None)
    if meetings_tab is None:
        return None
    from modules.meetings.ui.meetings_actions import open_contribution_report

    open_contribution_report(meetings_tab)
    dlg = getattr(meetings_tab, "contribution_dialog", None)
    if dlg is None:
        return None
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_spec_history(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    specs_tab = getattr(main_window, "specs_tab", None)
    spec_number = str(state.get("spec_number") or "").strip()
    if not spec_number or specs_tab is None:
        return None

    from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog

    dlg = SpecVersionHistoryDialog(
        spec_number=spec_number,
        initial_version=str(state.get("version") or ""),
        parent=specs_tab,
        history_db=getattr(specs_tab, "history_db", None),
        specs_db=getattr(specs_tab, "db", None),
        initial_search=str(state.get("search_text") or ""),
        initial_clause=str(state.get("clause_text") or ""),
        sort_by_origin_relevance=bool(state.get("origin_ranking", False)),
    )
    specs_tab._history_dialog = dlg
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_spec_clause_excerpt(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    specs_tab = getattr(main_window, "specs_tab", None)
    spec_num = str(state.get("spec_num") or "").strip()
    version = str(state.get("version") or "").strip()
    spec_dl_dir = str(state.get("spec_dl_dir") or "").strip()
    if not spec_num or not version or not spec_dl_dir or not Path(spec_dl_dir).exists():
        return None

    from modules.specifications.ui.spec_clause_dialog import SpecClauseExcerptDialog

    dlg = SpecClauseExcerptDialog(
        spec_num=spec_num,
        spec_type=str(state.get("spec_type") or "TS"),
        version=version,
        fname=str(state.get("fname") or ""),
        url=str(state.get("url") or ""),
        spec_dl_dir=Path(spec_dl_dir),
        parent=specs_tab,
    )
    if specs_tab is not None:
        if not hasattr(specs_tab, "_active_excerpt_dialogs"):
            specs_tab._active_excerpt_dialogs = []
        specs_tab._active_excerpt_dialogs.append(dlg)
        dlg.destroyed.connect(
            lambda: specs_tab._active_excerpt_dialogs.remove(dlg)
            if hasattr(specs_tab, "_active_excerpt_dialogs") and dlg in specs_tab._active_excerpt_dialogs
            else None
        )
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_spec_clause_diff(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    spec_search_tab = getattr(main_window, "spec_search_tab", None)
    if spec_search_tab is None:
        return None
    spec_number = str(state.get("spec_number") or "").strip()
    version_b = str(state.get("version_b") or "").strip()
    clause_number = str(state.get("clause_number") or "").strip()
    if not spec_number or not version_b or not clause_number:
        return None

    from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog

    dlg = SpecClauseDiffDialog(
        db=spec_search_tab.db,
        spec_number=spec_number,
        current_version=version_b,
        clause_number=clause_number,
        clause_title=str(state.get("clause_title") or ""),
        cache_dir=getattr(spec_search_tab, "cache_dir", None),
        specs_db=getattr(spec_search_tab, "specs_db", None),
        base_version=str(state.get("version_a") or "") or None,
        focus_attribution=False,
        parent=spec_search_tab,
    )
    spec_search_tab._diff_dialog = dlg
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


def _restore_cr_pcr(main_window: Any, state: Dict[str, Any], descriptor: Dict[str, Any]) -> Optional[QWidget]:
    specs_tab = getattr(main_window, "specs_tab", None)
    from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

    spec_num = str(state.get("spec_num") or "").strip()
    version = str(state.get("version") or "").strip()
    spec_dl_dir_raw = str(state.get("spec_dl_dir") or "").strip()
    spec_dl_dir = Path(spec_dl_dir_raw) if spec_dl_dir_raw else None

    dlg = CrPcrDialog(
        spec_num=spec_num,
        spec_type=str(state.get("spec_type") or "TS"),
        version=version,
        release=str(state.get("release") or ""),
        fname=str(state.get("fname") or ""),
        url=str(state.get("url") or ""),
        spec_dl_dir=spec_dl_dir,
        meeting_info=state.get("meeting_info") if isinstance(state.get("meeting_info"), dict) else {},
        tdoc_info=state.get("tdoc_info") if isinstance(state.get("tdoc_info"), dict) else {},
        initial_mode=str(state.get("mode") or "pcr"),
        parent=specs_tab or main_window,
        auto_start=bool(spec_num and version),
    )
    owner = specs_tab or main_window
    if owner is not None:
        if not hasattr(owner, "_active_cr_dialogs"):
            owner._active_cr_dialogs = []
        owner._active_cr_dialogs.append(dlg)
        dlg.destroyed.connect(
            lambda: owner._active_cr_dialogs.remove(dlg)
            if hasattr(owner, "_active_cr_dialogs") and dlg in owner._active_cr_dialogs
            else None
        )
    _show_restored_widget(dlg, descriptor)
    if hasattr(dlg, "restore_session_state"):
        dlg.restore_session_state(state)
    return dlg


_RESTORERS: Dict[str, Callable[[Any, Dict[str, Any], Dict[str, Any]], Optional[QWidget]]] = {
    "tdocs_window": _restore_tdocs_window,
    "session_plan": _restore_session_plan,
    "email_manager": _restore_email_manager,
    "meeting_email_inbox": _restore_meeting_email_inbox,
    "topic_dossier": _restore_topic_dossier,
    "agreement_tracker": _restore_agreement_tracker,
    "contribution_report": _restore_contribution_report,
    "spec_history": _restore_spec_history,
    "spec_clause_excerpt": _restore_spec_clause_excerpt,
    "spec_clause_diff": _restore_spec_clause_diff,
    "cr_pcr": _restore_cr_pcr,
}


def restore_open_windows(
    main_window: Any,
    settings: Optional[QSettings] = None,
    argv: Optional[Iterable[str]] = None,
    modifiers: Optional[Qt.KeyboardModifiers] = None,
) -> int:
    """
    Restore previously open workspace windows if enabled for this launch.

    Returns the number of windows successfully reopened.
    """
    cfg = settings or create_shell_settings()
    if not should_restore_windows(settings=cfg, argv=argv, modifiers=modifiers):
        return 0

    descriptors = load_open_windows_state(cfg)
    if not descriptors:
        return 0

    restored_count = 0
    for desc in descriptors:
        win_type = str(desc.get("type") or "").strip()
        state = desc.get("state")
        if not isinstance(state, dict):
            state = {}
        restorer = _RESTORERS.get(win_type)
        if restorer is None:
            continue
        try:
            widget = restorer(main_window, state, desc)
            if widget is not None:
                restored_count += 1
        except Exception as exc:
            logger.warning("⚠️ [SESSION] Could not restore window '%s': %s", win_type, exc)

    if restored_count > 0:
        logger.info("🪟 [SESSION] Restored %d workspace window(s).", restored_count)
    return restored_count
