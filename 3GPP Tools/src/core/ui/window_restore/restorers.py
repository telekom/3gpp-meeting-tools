"""Per-window restoration handlers and session restore orchestrator."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Optional

from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtWidgets import QWidget

from core.ui.shell_state import create_shell_settings

from .preferences import (
    _show_restored_widget,
    load_open_windows_state,
    should_restore_windows,
)

logger = logging.getLogger(__name__)


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
    """Restore previously open workspace windows if enabled for this launch."""
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
