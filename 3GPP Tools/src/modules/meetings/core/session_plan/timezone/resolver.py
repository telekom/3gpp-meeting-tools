"""City lookup, IANA zoneinfo evaluation, and Agenda deadlines timezone auto-detection."""

from __future__ import annotations

import logging
import re
from datetime import datetime, timezone
from typing import Optional, Tuple

from modules.meetings.core.session_plan.timezone.cities import (
    CITY_TO_IANA_TZ,
    KNOWN_3GPP_CITIES,
    _SORTED_CITY_KEYS,
    format_utc_offset,
)

try:
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
except ImportError:  # pragma: no cover - Python < 3.9 fallback
    ZoneInfo = None  # type: ignore[assignment]
    ZoneInfoNotFoundError = Exception  # type: ignore[assignment]

logger = logging.getLogger(__name__)


class TimezoneResolverMixin:
    """Mixin resolving laptop, Agenda-deadline, IANA, and 3GPP city timezones."""

    @classmethod
    def get_laptop_tz_offset_minutes(cls) -> int:
        """
        Calculates the user's laptop local UTC offset in minutes right now.
        For an Austrian corporate laptop, this returns +60 (CET in winter) or +120 (CEST in summer).
        """
        now = datetime.now()
        utc_now = datetime.now(timezone.utc).replace(tzinfo=None)
        diff = now - utc_now
        return int(round(diff.total_seconds() / 60))

    @classmethod
    def get_laptop_tz_label(cls) -> str:
        """User-friendly label for laptop local time."""
        mins = cls.get_laptop_tz_offset_minutes()
        return f"Austrian Laptop Time ({format_utc_offset(mins)})"

    @classmethod
    def extract_tz_from_agenda_text(cls, text: str) -> Optional[Tuple[int, str]]:
        """
        Tier 1: Parses Agenda document deadlines table text for Local time vs UTC.
        Example: '| Start of meeting | Monday 12 October 2026 | 0900 Local time | 0700 UTC |'
        Computes offset = 09:00 - 07:00 = +120 minutes (UTC+02:00).
        """
        if not text:
            return None

        patterns = [
            r"(\d{3,4})\s*(?:hours|hrs)?\s*Local\s*time.*?(\d{3,4})\s*(?:hours|hrs)?\s*UTC",
            r"(\d{1,2}:\d{2})\s*Local\s*time.*?(\d{1,2}:\d{2})\s*UTC",
        ]

        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
            if m:
                local_raw, utc_raw = m.group(1), m.group(2)
                local_min = cls._time_str_to_minutes(local_raw)
                utc_min = cls._time_str_to_minutes(utc_raw)
                if local_min is not None and utc_min is not None:
                    diff_min = local_min - utc_min
                    if diff_min > 720:
                        diff_min -= 1440
                    elif diff_min < -720:
                        diff_min += 1440

                    label = f"{format_utc_offset(diff_min)} (from Agenda Deadlines)"
                    logger.info(f"[SessionPlanTZ] Auto-detected meeting timezone from Agenda: {diff_min}m ({label})")
                    return (diff_min, label)

        return None

    @classmethod
    def _match_city_key(cls, location_str: str) -> Optional[str]:
        if not location_str:
            return None
        loc_lower = location_str.strip().lower()
        if not loc_lower:
            return None

        # 1. Word-boundary match (longest key first)
        for city_key in _SORTED_CITY_KEYS:
            pattern = r"(?<!\w)" + re.escape(city_key) + r"(?!\w)"
            if re.search(pattern, loc_lower):
                return city_key

        # 2. Substring fallback (longest key first)
        for city_key in _SORTED_CITY_KEYS:
            if city_key in loc_lower:
                return city_key

        return None

    @classmethod
    def _parse_reference_utc(cls, meeting_date_str: str = "") -> datetime:
        """Parses a meeting date string into a reference UTC datetime, or returns current UTC."""
        now_utc = datetime.now(timezone.utc)
        if not meeting_date_str:
            return now_utc

        s = meeting_date_str.strip()
        m_iso = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
        if m_iso:
            try:
                y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
                return datetime(y, m, d, 12, 0, tzinfo=timezone.utc)
            except ValueError:
                pass

        month_names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        m_mon = re.search(
            r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b",
            s,
            re.I,
        )
        if m_mon:
            month = month_names.index(m_mon.group(1)[:3].lower()) + 1
            m_year = re.search(r"\b(19\d{2}|20\d{2})\b", s)
            year = int(m_year.group(1)) if m_year else now_utc.year
            return datetime(year, month, 15, 12, 0, tzinfo=timezone.utc)

        m_num = re.search(r"[-/](\d{1,2})[-/]", s)
        if m_num:
            try:
                month = max(1, min(12, int(m_num.group(1))))
                return datetime(now_utc.year, month, 15, 12, 0, tzinfo=timezone.utc)
            except ValueError:
                pass

        return now_utc

    @classmethod
    def resolve_iana_offset_minutes(
        cls, iana_tz: str, ref_utc: Optional[datetime] = None
    ) -> Optional[Tuple[int, str]]:
        """Resolves an IANA timezone string to (offset_minutes, tz_abbr) using Python zoneinfo."""
        if not iana_tz or ZoneInfo is None:
            return None
        try:
            zi = ZoneInfo(iana_tz)
            utc_dt = ref_utc or datetime.now(timezone.utc)
            aware_dt = utc_dt.astimezone(zi)
            utc_offset = aware_dt.utcoffset()
            if utc_offset is None:
                return None
            offset_min = int(round(utc_offset.total_seconds() / 60))
            abbr = aware_dt.tzname() or ""
            return offset_min, abbr
        except ZoneInfoNotFoundError:
            return None
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] ZoneInfo resolution failed for '{iana_tz}': {e}")
            return None

    @classmethod
    def lookup_city_timezone_details(
        cls, location_str: str, meeting_date_str: str = ""
    ) -> Optional[Tuple[int, str, str]]:
        """
        Matches meeting location string against known 3GPP venue cities/countries.
        Uses Python zoneinfo (exact IANA DST rules) when available, falling back to
        static month-based DST rules if tzdata is missing.
        Returns (offset_minutes, label, iana_tz) or None.
        """
        matched_city = cls._match_city_key(location_str)
        if not matched_city:
            return None

        std_min, dst_min, dst_start, dst_end, tz_name = KNOWN_3GPP_CITIES[matched_city]
        iana_tz = CITY_TO_IANA_TZ.get(matched_city, "")
        ref_utc = cls._parse_reference_utc(meeting_date_str)

        # 1. Preferred: exact IANA zoneinfo evaluation (handles regional DST law changes)
        if iana_tz:
            zi_res = cls.resolve_iana_offset_minutes(iana_tz, ref_utc=ref_utc)
            if zi_res is not None:
                chosen_offset, abbr = zi_res
                suffix = f"{iana_tz} - {abbr}" if abbr and not abbr.startswith(("+", "-")) else iana_tz
                label = f"{format_utc_offset(chosen_offset)} ({suffix})"
                return (chosen_offset, label, iana_tz)

        # 2. Fallback: static month-based DST heuristic
        month = ref_utc.month
        is_dst = False
        if dst_start > 0 and dst_end > 0:
            if dst_start < dst_end:
                is_dst = dst_start <= month <= dst_end
            else:
                is_dst = month >= dst_start or month <= dst_end

        chosen_offset = dst_min if is_dst else std_min
        label = f"{format_utc_offset(chosen_offset)} ({tz_name})"
        logger.debug(f"[SessionPlanTZ] Fallback resolved '{matched_city}' to: {chosen_offset}m ({label})")
        return (chosen_offset, label, iana_tz)

    @classmethod
    def lookup_city_timezone(cls, location_str: str, meeting_date_str: str = "") -> Optional[Tuple[int, str]]:
        """
        Tier 2: Matches meeting location string against known 3GPP venue cities.
        Accounts for Daylight Saving Time based on meeting date (or current date).
        """
        details = cls.lookup_city_timezone_details(location_str, meeting_date_str=meeting_date_str)
        if details is None:
            return None
        offset_min, label, _iana = details
        return (offset_min, label)
