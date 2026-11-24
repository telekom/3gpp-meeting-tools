"""Per-meeting timezone persistence, dual-clock range formatting, and UTC conversions."""

from __future__ import annotations

import json
import logging
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Tuple

from modules.meetings.core.session_plan.timezone.cities import (
    MEETING_TZ_FILENAME,
    format_utc_offset,
)

logger = logging.getLogger(__name__)


class TimezonePersistenceAndFormatMixin:
    """Mixin providing meeting_timezone.json persistence and dual-time formatting helpers."""

    @classmethod
    def load_meeting_folder_timezone(cls, meeting_dir: Path) -> Optional[dict]:
        """Reads `<meeting_dir>/meeting_timezone.json` if it exists and is valid."""
        if not meeting_dir:
            return None
        tz_file = Path(meeting_dir) / MEETING_TZ_FILENAME
        if not tz_file.is_file():
            return None
        try:
            data = json.loads(tz_file.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or data.get("tz_offset_minutes") is None:
                return None
            return {
                "wg": str(data.get("wg", "")),
                "meeting_number": str(data.get("meeting_number", "")),
                "location": str(data.get("location", "")),
                "tz_offset_minutes": int(data["tz_offset_minutes"]),
                "tz_label": str(data.get("tz_label", "")),
                "iana_tz": str(data.get("iana_tz", "")),
                "source": str(data.get("source", "meeting_folder")),
                "updated_at": str(data.get("updated_at", "")),
            }
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] Failed to load {tz_file}: {e}")
            return None

    @classmethod
    def save_meeting_folder_timezone(
        cls,
        meeting_dir: Path,
        tz_offset_minutes: int,
        tz_label: str,
        source: str = "manual_override",
        location: str = "",
        wg: str = "",
        meeting_number: str = "",
        iana_tz: str = "",
    ) -> bool:
        """
        Persists confirmed/overridden timezone configuration inside `<meeting_dir>/meeting_timezone.json`
        and synchronizes `Agenda/session_plans.db` (`meeting_meta`) if that database exists.
        """
        if not meeting_dir:
            return False
        mtg_path = Path(meeting_dir)
        tz_file = mtg_path / MEETING_TZ_FILENAME
        payload = {
            "wg": str(wg or ""),
            "meeting_number": str(meeting_number or ""),
            "location": str(location or ""),
            "tz_offset_minutes": int(tz_offset_minutes),
            "tz_label": str(tz_label or format_utc_offset(int(tz_offset_minutes))),
            "iana_tz": str(iana_tz or ""),
            "source": str(source or "manual_override"),
            "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        try:
            mtg_path.mkdir(parents=True, exist_ok=True)
            tmp_file = tz_file.with_suffix(".json.tmp")
            tmp_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp_file, tz_file)
        except Exception as e:
            logger.warning(f"[SessionPlanTZ] Could not save {tz_file}: {e}")
            return False

        # Keep existing Agenda/session_plans.db synchronized when present
        sp_db_path = mtg_path / "Agenda" / "session_plans.db"
        if sp_db_path.is_file() and meeting_number:
            try:
                from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase

                sp_db = SessionPlanDatabase(sp_db_path)
                sp_db.set_meeting_meta(
                    meeting_number=str(meeting_number),
                    location=str(location or ""),
                    tz_offset_minutes=int(tz_offset_minutes),
                    tz_name=payload["tz_label"],
                    tz_confirmed=1,
                )
            except Exception as e:
                logger.debug(f"[SessionPlanTZ] Could not sync session_plans.db: {e}")

        return True

    @classmethod
    def clear_meeting_folder_timezone(cls, meeting_dir: Path) -> bool:
        """Removes `<meeting_dir>/meeting_timezone.json` if it exists."""
        if not meeting_dir:
            return False
        tz_file = Path(meeting_dir) / MEETING_TZ_FILENAME
        if tz_file.is_file():
            try:
                tz_file.unlink()
                return True
            except Exception as e:
                logger.debug(f"[SessionPlanTZ] Could not delete {tz_file}: {e}")
                return False
        return False

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
    def format_laptop_time_range(
        cls, start_hhmm: str, end_hhmm: str, meeting_tz_offset_min: int
    ) -> Tuple[str, str]:
        """Returns (local_start_hhmm, local_end_hhmm) when laptop TZ differs from venue TZ, else ('', '')."""
        laptop_offset_min = cls.get_laptop_tz_offset_minutes()
        offset_diff = laptop_offset_min - meeting_tz_offset_min
        if offset_diff == 0:
            return ("", "")

        s_min = cls._time_str_to_minutes(start_hhmm)
        e_min = cls._time_str_to_minutes(end_hhmm)
        if s_min is None or e_min is None:
            return ("", "")

        lap_s = (s_min + offset_diff) % 1440
        lap_e = (e_min + offset_diff) % 1440
        return (f"{lap_s // 60:02d}:{lap_s % 60:02d}", f"{lap_e // 60:02d}:{lap_e % 60:02d}")

    @classmethod
    def format_dual_time_range(
        cls, start_hhmm: str, end_hhmm: str, meeting_tz_offset_min: int
    ) -> str:
        """
        Formats a time range for display.
        If laptop timezone differs from meeting venue, appends the Austrian laptop time in parentheses.
        Example: '09:00 – 10:30 (02:00 – 03:30 Vienna)'
        """
        venue_str = f"{start_hhmm} – {end_hhmm}"
        lap_s_str, lap_e_str = cls.format_laptop_time_range(start_hhmm, end_hhmm, meeting_tz_offset_min)
        if not lap_s_str:
            return venue_str
        return f"{venue_str} ({lap_s_str} – {lap_e_str} Vienna)"

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

            d = datetime.strptime(date_str, "%Y-%m-%d").date()
            venue_dt = datetime(d.year, d.month, d.day, hh, mm)
            venue_tz = timezone(timedelta(minutes=meeting_tz_offset_min))
            venue_aware = venue_dt.replace(tzinfo=venue_tz)
            return venue_aware.astimezone(timezone.utc)
        except Exception as e:
            logger.debug(f"[SessionPlanTZ] Error converting {date_str} {time_hhmm} to UTC: {e}")
            return None
