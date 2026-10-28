# --- File: src/modules/meetings/core/session_plan_timezone.py ---
"""
Timezone engine for 3GPP Meeting Session Plans.
Solves the corporate Austrian laptop clock problem (Europe/Vienna CET/CEST)
by strictly anchoring calculations in true UTC and translating between
Venue Local Time and Austrian Laptop Time.
"""
import logging
import re
from datetime import datetime, time, timedelta, timezone
from typing import Dict, Optional, Tuple

logger = logging.getLogger(__name__)

# Known 3GPP meeting host cities with standard and daylight savings UTC offsets (in minutes)
# Format: {city_key: (std_offset_min, dst_offset_min, dst_months_start, dst_months_end, tz_name)}
KNOWN_3GPP_CITIES: Dict[str, Tuple[int, int, int, int, str]] = {
    "prague": (60, 120, 3, 10, "Europe/Prague (CET/CEST)"),
    "vienna": (60, 120, 3, 10, "Europe/Vienna (CET/CEST)"),
    "gothenburg": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "stockholm": (60, 120, 3, 10, "Europe/Stockholm (CET/CEST)"),
    "helsinki": (120, 180, 3, 10, "Europe/Helsinki (EET/EEST)"),
    "athens": (120, 180, 3, 10, "Europe/Athens (EET/EEST)"),
    "dubrovnik": (60, 120, 3, 10, "Europe/Zagreb (CET/CEST)"),
    "valencia": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "barcelona": (60, 120, 3, 10, "Europe/Madrid (CET/CEST)"),
    "sophia antipolis": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "nice": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "paris": (60, 120, 3, 10, "Europe/Paris (CET/CEST)"),
    "fukuoka": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "tokyo": (540, 540, 0, 0, "Asia/Tokyo (JST, UTC+9)"),
    "incheon": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "jeju": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "seoul": (540, 540, 0, 0, "Asia/Seoul (KST, UTC+9)"),
    "shanghai": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "beijing": (480, 480, 0, 0, "Asia/Shanghai (CST, UTC+8)"),
    "taipei": (480, 480, 0, 0, "Asia/Taipei (CST, UTC+8)"),
    "reno": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "san diego": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "sparks": (-480, -420, 3, 11, "America/Los_Angeles (PST/PDT)"),
    "denver": (-420, -360, 3, 11, "America/Denver (MST/MDT)"),
    "chicago": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "austin": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "dallas": (-360, -300, 3, 11, "America/Chicago (CST/CDT)"),
    "orlando": (-300, -240, 3, 11, "America/New_York (EST/EDT)"),
    "cancun": (-300, -300, 0, 0, "America/Cancun (EST, UTC-5)"),
    "gold coast": (600, 600, 0, 0, "Australia/Brisbane (AEST, UTC+10)"),
    "brisbane": (600, 600, 0, 0, "Australia/Brisbane (AEST, UTC+10)"),
    "melbourne": (600, 660, 10, 4, "Australia/Melbourne (AEST/AEDT)"),
}


class SessionPlanTimezone:
    """Manages timezone calculations, dual clocks, and UTC conversions."""

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
        """User-friendly label for Austrian corporate laptop time."""
        mins = cls.get_laptop_tz_offset_minutes()
        hours = mins // 60
        sign = "+" if hours >= 0 else "-"
        return f"Austrian Laptop Time (UTC{sign}{abs(hours):02d}:{abs(mins % 60):02d})"

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
                    # Handle day wrapping (e.g. 01:00 Local vs 23:00 UTC)
                    if diff_min > 720:
                        diff_min -= 1440
                    elif diff_min < -720:
                        diff_min += 1440

                    hours = diff_min // 60
                    mins = abs(diff_min % 60)
                    sign = "+" if diff_min >= 0 else "-"
                    label = f"UTC{sign}{abs(hours):02d}:{mins:02d} (from Agenda Deadlines)"
                    logger.info(f"[SessionPlanTZ] Auto-detected meeting timezone from Agenda: {diff_min}m ({label})")
                    return (diff_min, label)

        return None

    @classmethod
    def lookup_city_timezone(cls, location_str: str, meeting_date_str: str = "") -> Optional[Tuple[int, str]]:
        """
        Tier 2: Matches meeting location string against known 3GPP venue cities.
        Accounts for Daylight Saving Time based on meeting month.
        """
        if not location_str:
            return None

        loc_lower = location_str.lower()
        matched_city = None
        for city_key in KNOWN_3GPP_CITIES:
            if city_key in loc_lower:
                matched_city = city_key
                break

        if not matched_city:
            return None

        std_min, dst_min, dst_start, dst_end, tz_name = KNOWN_3GPP_CITIES[matched_city]

        # Determine month for DST check
        month = 10  # Default autumn
        if meeting_date_str:
            m = re.search(r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b", meeting_date_str, re.I)
            if m:
                month_names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
                month = month_names.index(m.group(1)[:3].lower()) + 1
            else:
                m_num = re.search(r"[-/](\d{1,2})[-/]", meeting_date_str)
                if m_num:
                    month = int(m_num.group(1))

        # Check DST
        is_dst = False
        if dst_start > 0 and dst_end > 0:
            if dst_start < dst_end:
                is_dst = dst_start <= month <= dst_end
            else:
                # Southern hemisphere DST (spans year end)
                is_dst = month >= dst_start or month <= dst_end

        chosen_offset = dst_min if is_dst else std_min
        hours = chosen_offset // 60
        mins = abs(chosen_offset % 60)
        sign = "+" if chosen_offset >= 0 else "-"
        label = f"UTC{sign}{abs(hours):02d}:{mins:02d} ({tz_name})"
        logger.info(f"[SessionPlanTZ] Resolved city '{matched_city}' to: {chosen_offset}m ({label})")
        return (chosen_offset, label)

    @staticmethod
    def _time_str_to_minutes(s: str) -> Optional[int]:
        """Converts '0900', '900', '09:00', '9:00' to minutes since midnight."""
        s = re.sub(r"[^\d:]", "", s.strip())
        if ":" in s:
            parts = s.split(":")
            return int(parts[0]) * 60 + int(parts[1])
        if len(s) in (3, 4):
            hh = int(s[:-2])
            mm = int(s[-2:])
            return hh * 60 + mm
        return None

    @classmethod
    def get_venue_current_datetime(cls, meeting_tz_offset_min: int) -> datetime:
        """
        Converts the current physical instant into Venue Local datetime,
        completely immune to the laptop clock's local timezone.
        """
        utc_now = datetime.now(timezone.utc)
        venue_tz = timezone(timedelta(minutes=meeting_tz_offset_min))
        return utc_now.astimezone(venue_tz)

    @classmethod
    def format_dual_time_range(
        cls, start_hhmm: str, end_hhmm: str, meeting_tz_offset_min: int
    ) -> str:
        """
        Formats a time range for display.
        If laptop timezone differs from meeting venue, appends the Austrian laptop time in parentheses.
        Example: '09:00 – 10:30 (02:00 – 03:30 Vienna)'
        """
        laptop_offset_min = cls.get_laptop_tz_offset_minutes()
        offset_diff = laptop_offset_min - meeting_tz_offset_min

        venue_str = f"{start_hhmm} – {end_hhmm}"
        if offset_diff == 0:
            return venue_str

        # Calculate corresponding Austrian laptop times
        s_min = cls._time_str_to_minutes(start_hhmm)
        e_min = cls._time_str_to_minutes(end_hhmm)
        if s_min is None or e_min is None:
            return venue_str

        lap_s = (s_min + offset_diff) % 1440
        lap_e = (e_min + offset_diff) % 1440

        lap_str = f"{lap_s // 60:02d}:{lap_s % 60:02d} – {lap_e // 60:02d}:{lap_e % 60:02d} Vienna"
        return f"{venue_str} ({lap_str})"

    @classmethod
    def venue_local_to_utc_dt(
        cls, date_str: str, time_hhmm: str, meeting_tz_offset_min: int
    ) -> Optional[datetime]:
        """
        Converts a Venue Local date ('2026-10-12') and time ('09:00') into an aware UTC datetime object.
        Essential for generating accurate RFC 5545 .ics calendar entries.
        """
        try:
            m = cls._time_str_to_minutes(time_hhmm)
            if m is None:
                return None
            hh = m // 60
            mm = m % 60

            # Parse date
            d = datetime.strptime(date_str, "%Y-%m-%d").date()
            venue_dt = datetime(d.year, d.month, d.day, hh, mm)
            venue_tz = timezone(timedelta(minutes=meeting_tz_offset_min))
            venue_aware = venue_dt.replace(tzinfo=venue_tz)
            return venue_aware.astimezone(timezone.utc)
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] Error converting {date_str} {time_hhmm} to UTC: {e}")
            return None
