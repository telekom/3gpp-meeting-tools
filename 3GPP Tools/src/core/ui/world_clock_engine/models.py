"""Data models, city query normalization, and freshness helpers for World Clock."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Dict, List, Optional

_COUNTRY_CITY_FALLBACKS: Dict[str, str] = {
    "india": "New Delhi",
    "china": "Beijing",
    "south korea": "Seoul",
    "korea": "Seoul",
    "malta": "Valletta",
    "japan": "Tokyo",
    "finland": "Helsinki",
    "texas": "Dallas",
    "virginia": "Reston",
    "new jersey": "Newark",
    "puerto rico": "San Juan",
    "hawaii": "Honolulu",
}


@dataclass
class VenueClockState:
    """Resolved live clock state for a single 3GPP meeting venue location."""

    location: str
    wgs: List[str] = field(default_factory=list)
    matching_wgs: List[str] = field(default_factory=list)
    meetings: List[dict] = field(default_factory=list)
    tz_offset_minutes: Optional[int] = None
    tz_label: str = "Unresolved Timezone"
    iana_tz: str = ""
    source: str = "unresolved"  # "manual_override" | "meeting_folder" | "web_api" | "builtin" | "unresolved"
    venue_dt: Optional[datetime] = None
    day_diff_vs_laptop: int = 0
    offset_diff_vs_laptop_min: int = 0
    is_upcoming: bool = False
    start_date: str = ""
    end_date: str = ""


@dataclass
class WorldClockSnapshot:
    """Complete snapshot of reference clocks and active/upcoming 3GPP venues."""

    utc_dt: datetime
    laptop_dt: datetime
    laptop_offset_minutes: int
    laptop_label: str
    primary_venues: List[VenueClockState] = field(default_factory=list)
    other_venues: List[VenueClockState] = field(default_factory=list)
    wg_filter: List[str] = field(default_factory=list)
    show_utc_in_bar: bool = True
    show_upcoming_when_idle: bool = True


def _clean_city_query(location: str) -> str:
    """Normalizes a 3GPP calendar location string into a clean city query for geocoding."""
    raw = (location or "").strip()
    if not raw:
        return ""
    low = raw.lower()
    if low in _COUNTRY_CITY_FALLBACKS:
        return _COUNTRY_CITY_FALLBACKS[low]

    if "gothenburg" in low or "göteborg" in low or "goteborg" in low:
        return "Gothenburg"

    cleaned = re.sub(r"\(.*?\)", "", raw).strip()
    cleaned = re.sub(r"\bmetropolitan area\b", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = cleaned.replace("-", " ").strip()

    # Strip trailing 2-letter country/state code if separated by comma (e.g. "Incheon, KR" -> "Incheon")
    if "," in cleaned:
        parts = [p.strip() for p in cleaned.split(",") if p.strip()]
        if parts:
            cleaned = parts[0]

    return _COUNTRY_CITY_FALLBACKS.get(cleaned.lower(), cleaned)


def _short_city_name(location: str) -> str:
    """Shortens verbose venue names for compact status-bar display."""
    raw = (location or "").strip()
    if not raw:
        return "Venue"
    low = raw.lower()
    if "gothenburg" in low or "göteborg" in low or "goteborg" in low:
        return "Gothenburg"
    if "," in raw:
        return raw.split(",", 1)[0].strip()
    if len(raw) > 22 and " " in raw:
        return raw.split(" ", 1)[0].strip()
    return raw


def _format_wg_summary(wgs: List[str], max_items: int = 6) -> str:
    """Formats a sorted list of WGs compactly."""
    clean = sorted({str(w).strip() for w in wgs if str(w).strip()})
    if not clean:
        return ""
    if len(clean) <= max_items:
        return ", ".join(clean)
    return ", ".join(clean[:max_items]) + f" +{len(clean) - max_items}"


def _is_web_entry_fresh(updated_at: str, max_age_days: int = 7) -> bool:
    if not updated_at:
        return False
    try:
        ts = updated_at.strip()
        if ts.endswith("Z"):
            ts = ts[:-1] + "+00:00"
        dt = datetime.fromisoformat(ts)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        age = (datetime.now(timezone.utc) - dt).total_seconds()
        return 0 <= age <= max_age_days * 86400
    except Exception:
        return False
