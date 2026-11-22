# --- File: src/modules/emails/ui/meeting_inbox/followed_ais_helper.py ---
"""
Helper for resolving and persisting a meeting's Followed Agenda Items (`👀 Followed AIs`).
Combines inherited AIs (from Followed/Starred Work Items, Session Plan assignments,
`🔵 My TDoc` contributions, and legacy `followed_ais` tables) with per-meeting user overrides.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from modules.emails.core.legacy_email_migrator import read_legacy_followed_ais_and_starred_tdocs


def _load_tracked_wi_tokens() -> Set[str]:
    """Returns uppercase WI tokens for all globally followed or starred Work Items."""
    try:
        from core.utils.paths import get_project_root
        from modules.work_items.core.wi_database import WorkItemsDatabase
        from modules.work_items.core.wi_settings import WorkItemsSettings

        settings = WorkItemsSettings()
        starred_codes, following_codes = settings.get_tracking_sets()
        all_codes = set(starred_codes or set()) | set(following_codes or set())
        if not all_codes:
            return set()
        db_path = get_project_root() / "db" / "3gpp_data.db"
        if db_path.exists():
            wi_db = WorkItemsDatabase(db_path)
            tokens = wi_db.get_tracked_wi_tokens(all_codes)
            return {str(t).strip().upper() for t in (tokens or set()) if str(t).strip()}
        return {str(c).strip().upper() for c in all_codes if str(c).strip()}
    except Exception:
        return set()


def _load_session_plan_followed_ais(agenda_dir: Optional[Path]) -> Set[str]:
    """Reads AIs marked as Assigned or Followed in `<meeting_dir>/Agenda/session_plan.db` if present."""
    out: Set[str] = set()
    if not agenda_dir:
        return out
    sp_db = Path(agenda_dir) / "session_plan.db"
    if not sp_db.is_file():
        return out
    try:
        conn = sqlite3.connect(sp_db, timeout=5.0)
        try:
            tables = {
                r[0]
                for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }
            if "ai_delegation" in tables:
                cols = {r[1] for r in conn.execute("PRAGMA table_info(ai_delegation)").fetchall()}
                if "follow_state" in cols:
                    rows = conn.execute(
                        "SELECT agenda_item FROM ai_delegation WHERE LOWER(COALESCE(follow_state, '')) IN ('assigned', 'followed')"
                    ).fetchall()
                elif "is_followed" in cols:
                    rows = conn.execute(
                        "SELECT agenda_item FROM ai_delegation WHERE COALESCE(is_followed, 1) = 1 AND COALESCE(assigned_delegates, '') != ''"
                    ).fetchall()
                else:
                    rows = []
                for r in rows:
                    ai = str(r[0] or "").strip()
                    if ai:
                        out.add(ai)
        finally:
            conn.close()
    except Exception:
        pass
    return out


def compute_inherited_followed_ais(
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db_path: Optional[Path] = None,
    tracked_wi_tokens: Optional[Set[str]] = None,
) -> Set[str]:
    """
    Derives the default Followed AIs for a meeting from:
    1. AIs containing TDocs linked to the user's Followed / Starred Work Items
    2. AIs containing TDocs with `My Status == '🔵 My TDoc'` (or non-Neutral stance)
    3. AIs assigned in the meeting's Session Plan (`session_plan.db`)
    4. Legacy `followed_ais` table in `Agenda/emails.db` (if present)
    """
    inherited: Set[str] = set()
    meta_map = tdoc_meta_map or {}
    wi_tokens = (
        {str(t).strip().upper() for t in tracked_wi_tokens if str(t).strip()}
        if tracked_wi_tokens is not None
        else _load_tracked_wi_tokens()
    )

    for _tdoc, meta in meta_map.items():
        ai = str(meta.get("ai") or "").strip()
        if not ai:
            continue
        my_status = str(meta.get("my_status") or "").strip()
        if my_status and my_status != "⚪ Neutral":
            inherited.add(ai)
            continue
        if wi_tokens:
            td_wis = {str(w).strip().upper() for w in (meta.get("wis") or []) if str(w).strip()}
            if td_wis & wi_tokens:
                inherited.add(ai)

    if general_email_db_path:
        db_p = Path(general_email_db_path)
        leg_ais, leg_starred = read_legacy_followed_ais_and_starred_tdocs(db_p)
        inherited.update(leg_ais)
        for st_tdoc in leg_starred:
            ai = str((meta_map.get(st_tdoc) or {}).get("ai") or "").strip()
            if ai:
                inherited.add(ai)
        inherited.update(_load_session_plan_followed_ais(db_p.parent))

    return inherited


def load_effective_followed_ais(
    settings: Any,
    settings_prefix: str,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db_path: Optional[Path] = None,
    available_ais: Optional[Iterable[str]] = None,
) -> Tuple[Set[str], Set[str], bool]:
    """
    Returns `(effective_followed_ais, inherited_followed_ais, has_custom_override)`.
    If the user has customized Followed AIs for this meeting in `QSettings`, returns their override;
    otherwise returns `inherited_followed_ais`.
    """
    inherited = compute_inherited_followed_ais(
        tdoc_meta_map=tdoc_meta_map,
        general_email_db_path=general_email_db_path,
    )
    if available_ais is not None:
        avail_set = {str(a).strip() for a in available_ais if str(a).strip()}
        if avail_set:
            inherited = {a for a in inherited if a in avail_set} or inherited

    if settings is not None:
        has_override = bool(settings.value(f"{settings_prefix}/followed_ais_customized", False, type=bool))
        if has_override:
            raw = str(settings.value(f"{settings_prefix}/followed_ais_list", "[]") or "[]")
            try:
                parsed = json.loads(raw)
                if isinstance(parsed, list):
                    custom = {str(x).strip() for x in parsed if str(x).strip()}
                    return custom, inherited, True
            except Exception:
                pass
    return set(inherited), inherited, False


def save_followed_ais_override(
    settings: Any,
    settings_prefix: str,
    selected_ais: Iterable[str],
    *,
    reset_to_inherited: bool = False,
) -> None:
    """Persists (or resets) the per-meeting Followed AIs override in `QSettings`."""
    if settings is None:
        return
    try:
        if reset_to_inherited:
            settings.setValue(f"{settings_prefix}/followed_ais_customized", False)
            settings.remove(f"{settings_prefix}/followed_ais_list")
        else:
            clean = sorted({str(a).strip() for a in selected_ais if str(a).strip()})
            settings.setValue(f"{settings_prefix}/followed_ais_customized", True)
            settings.setValue(f"{settings_prefix}/followed_ais_list", json.dumps(clean))
        settings.sync()
    except Exception:
        pass
