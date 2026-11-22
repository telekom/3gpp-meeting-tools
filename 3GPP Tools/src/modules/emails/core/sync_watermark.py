# --- File: src/modules/emails/core/sync_watermark.py ---
"""Sync watermark persistence, locale-safe MAPI Table.Restrict filters, and Live Auto-Sync state."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
import logging
import re
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)

DEFAULT_OVERLAP_MINUTES = 5
DEFAULT_LIVE_SYNC_INTERVAL_MS = 120_000  # 2 minutes
LAST_SYNC_META_KEY = "last_sync_utc"


def utc_now() -> datetime:
    """Returns current timezone-aware UTC datetime."""
    return datetime.now(timezone.utc)


def format_utc_iso(dt: datetime) -> str:
    """Formats a datetime in canonical UTC ISO-8601 format (`YYYY-MM-DDTHH:MM:SSZ`)."""
    aware = dt.astimezone(timezone.utc) if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    return aware.strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_utc_iso(raw: Any) -> Optional[datetime]:
    """Parses a persisted UTC ISO timestamp into an aware UTC datetime, or None if invalid."""
    text = str(raw or "").strip()
    if not text:
        return None
    try:
        normalized = f"{text[:-1]}+00:00" if text.endswith("Z") else text
        parsed = datetime.fromisoformat(normalized)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except Exception:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:19], fmt).replace(tzinfo=timezone.utc)
        except Exception:
            continue
    return None


def utc_to_local_naive(dt_utc: datetime) -> datetime:
    """Converts a UTC datetime to a local naive datetime matching Outlook ReceivedTime/SentOn."""
    aware = dt_utc.astimezone(timezone.utc) if dt_utc.tzinfo else dt_utc.replace(tzinfo=timezone.utc)
    return aware.astimezone().replace(tzinfo=None)


def local_naive_to_utc(dt_local: datetime) -> datetime:
    """Converts a local naive datetime into an aware UTC datetime."""
    if dt_local.tzinfo is not None:
        return dt_local.astimezone(timezone.utc)
    return dt_local.astimezone(timezone.utc)


def format_dasl_since_filter(since_dt_utc: datetime, is_sent_folder: bool = False) -> str:
    """Builds a locale-independent Outlook DASL Table.Restrict filter string in UTC."""
    aware = since_dt_utc.astimezone(timezone.utc) if since_dt_utc.tzinfo else since_dt_utc.replace(tzinfo=timezone.utc)
    ts = aware.strftime("%Y-%m-%d %H:%M:%S")
    schema_prop = "urn:schemas:httpmail:date" if is_sent_folder else "urn:schemas:httpmail:datereceived"
    return f'@SQL="{schema_prop}" >= \'{ts}\''


def apply_mapi_since_restriction(
    tbl: Any,
    since_dt_utc: Optional[datetime],
    is_sent_folder: bool = False,
) -> Tuple[Any, bool]:
    """Applies a MAPI Table.Restrict() filter and returns `(restricted_table, applied)`."""
    if tbl is None or since_dt_utc is None or not hasattr(tbl, "Restrict"):
        return tbl, False
    dasl = format_dasl_since_filter(since_dt_utc, is_sent_folder=is_sent_folder)
    try:
        restricted = tbl.Restrict(dasl)
        if restricted is not None:
            return restricted, True
    except Exception as exc:
        logger.debug("MAPI Table.Restrict fallback (%s): %s", dasl, exc)
    return tbl, False


def _parse_yyyy_mm_dd(raw: Any) -> Optional[date]:
    text = str(raw or "").strip()
    if len(text) < 10:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _format_gap_description(cur_utc: datetime, last_dt: datetime, since_utc: datetime) -> str:
    delta_sec = max(0, int((cur_utc - last_dt).total_seconds()))
    if delta_sec >= 86400:
        days = delta_sec // 86400
        hours = (delta_sec % 86400) // 3600
        return f"catch-up {days}d {hours}h"
    if delta_sec >= 3600:
        hours = delta_sec // 3600
        mins = (delta_sec % 3600) // 60
        return f"catch-up {hours}h {mins}m"
    local_since = utc_to_local_naive(since_utc)
    return f"since {local_since.strftime('%H:%M')}"


def resolve_sync_window(
    general_db: Any,
    meeting_start: str = "",
    meeting_end: str = "",
    *,
    sync_mode: str = "incremental",
    overlap_minutes: int = DEFAULT_OVERLAP_MINUTES,
    days_buffer: int = 3,
    now_utc: Optional[datetime] = None,
) -> Dict[str, Any]:
    """Resolves the Outlook sync window from the persisted DB watermark or meeting date bounds."""
    mode_clean = str(sync_mode or "incremental").strip().lower()
    if now_utc is None:
        cur_utc = utc_now()
    elif now_utc.tzinfo is None:
        cur_utc = now_utc.replace(tzinfo=timezone.utc)
    else:
        cur_utc = now_utc.astimezone(timezone.utc)

    today_local = cur_utc.astimezone().date()
    buf = max(0, int(days_buffer if days_buffer is not None else 3))

    if mode_clean == "full":
        s_d = _parse_yyyy_mm_dd(meeting_start) or (today_local - timedelta(days=10))
        e_d = _parse_yyyy_mm_dd(meeting_end) or today_local
        if e_d < s_d:
            e_d = s_d
        return {
            "mode": "full",
            "since_dt_utc": None,
            "start_date": s_d.strftime("%Y-%m-%d"),
            "end_date": e_d.strftime("%Y-%m-%d"),
            "days_buffer": buf,
            "is_first_sync": False,
            "gap_description": f"full ({s_d.strftime('%Y-%m-%d')}..{e_d.strftime('%Y-%m-%d')})",
            "last_sync_utc": "",
        }

    last_raw = ""
    if general_db is not None and hasattr(general_db, "get_sync_meta"):
        try:
            last_raw = str(general_db.get_sync_meta(LAST_SYNC_META_KEY, "") or "").strip()
        except Exception:
            last_raw = ""

    last_dt = parse_utc_iso(last_raw)
    if last_dt is not None:
        since_utc = last_dt - timedelta(minutes=max(1, int(overlap_minutes)))
        return {
            "mode": "incremental",
            "since_dt_utc": since_utc,
            "start_date": "",
            "end_date": "",
            "days_buffer": 0,
            "is_first_sync": False,
            "gap_description": _format_gap_description(cur_utc, last_dt, since_utc),
            "last_sync_utc": format_utc_iso(last_dt),
        }

    # First-time sync (no watermark stored yet): bound by meeting dates to avoid scanning years of history
    s_d = _parse_yyyy_mm_dd(meeting_start)
    e_d = _parse_yyyy_mm_dd(meeting_end)
    if s_d and e_d:
        first_start_d = s_d - timedelta(days=buf)
        end_bound_d = max(e_d, today_local)
        since_utc = local_naive_to_utc(datetime(first_start_d.year, first_start_d.month, first_start_d.day, 0, 0, 0))
        return {
            "mode": "incremental",
            "since_dt_utc": since_utc,
            "start_date": s_d.strftime("%Y-%m-%d"),
            "end_date": end_bound_d.strftime("%Y-%m-%d"),
            "days_buffer": buf,
            "is_first_sync": True,
            "gap_description": f"first sync ({first_start_d.strftime('%Y-%m-%d')}..)",
            "last_sync_utc": "",
        }

    fallback_start_d = today_local - timedelta(days=10)
    return {
        "mode": "incremental",
        "since_dt_utc": cur_utc - timedelta(days=10),
        "start_date": fallback_start_d.strftime("%Y-%m-%d"),
        "end_date": today_local.strftime("%Y-%m-%d"),
        "days_buffer": 0,
        "is_first_sync": True,
        "gap_description": "first sync (last 10d)",
        "last_sync_utc": "",
    }


def _safe_meeting_key(meeting_id: str) -> str:
    raw = str(meeting_id or "_global").strip()
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", raw).strip("_") or "_global"


def get_live_sync_settings():
    """Returns shared QSettings store for per-meeting Live Auto-Sync state."""
    from PyQt5.QtCore import QSettings
    return QSettings("3GPP_Tools", "EmailLiveSync")


def load_meeting_live_sync_enabled(meeting_id: str, settings: Any = None) -> bool:
    """Loads whether continuous Live Auto-Sync is enabled for the given meeting."""
    store = settings if settings is not None else get_live_sync_settings()
    key = f"meetings/{_safe_meeting_key(meeting_id)}/live_sync_enabled"
    try:
        return bool(store.value(key, False, type=bool))
    except Exception:
        return False


def save_meeting_live_sync_enabled(meeting_id: str, enabled: bool, settings: Any = None) -> None:
    """Persists continuous Live Auto-Sync toggle state for the given meeting."""
    store = settings if settings is not None else get_live_sync_settings()
    key = f"meetings/{_safe_meeting_key(meeting_id)}/live_sync_enabled"
    try:
        store.setValue(key, bool(enabled))
        store.sync()
    except Exception as exc:
        logger.debug("Could not persist live_sync_enabled: %s", exc)
