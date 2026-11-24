"""Open workspace window discovery and state collection."""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Set

from PyQt5.QtWidgets import QApplication, QWidget

from .preferences import encode_window_geometry

logger = logging.getLogger(__name__)


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

    try:
        top_widgets = list(QApplication.topLevelWidgets())
    except Exception:
        top_widgets = []

    for widget in top_widgets:
        if widget is main_window:
            continue
        _collect_from_widget(widget, seen_ids, descriptors, type_map)

    return descriptors
