"""High-level WorldClockService resolving multi-venue clocks and snapshots."""

from __future__ import annotations

import logging
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.ui.world_clock_engine.formatting import WorldClockFormattingMixin
from core.ui.world_clock_engine.models import (
    VenueClockState,
    WorldClockSnapshot,
    _is_web_entry_fresh,
)
from core.utils.paths import get_db_root
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.session_plan.session_plan_timezone import (
    SessionPlanTimezone,
    format_utc_offset,
)
from modules.meetings.core.settings import MeetingsSettings

logger = logging.getLogger(__name__)


class WorldClockService(WorldClockFormattingMixin):
    """
    High-level service managing active 3GPP meeting venue clocks, WG filtering,
    and per-meeting/global timezone overrides.
    """

    def __init__(
        self,
        db_path: Optional[Path] = None,
        settings: Optional[MeetingsSettings] = None,
    ):
        self.db_path = Path(db_path) if db_path is not None else (get_db_root() / "3gpp_data.db")
        self.settings = settings or MeetingsSettings()
        self._db: Optional[MeetingsDatabase] = None

        # Short-lived in-memory cache of DB meeting rows so 1s timer ticks do zero SQLite I/O
        self._cache_ts: float = 0.0
        self._cache_key: Tuple[str, int, bool] = ("", 2, True)
        self._cached_active_rows: List[dict] = []
        self._cached_upcoming_rows: List[dict] = []

    def _get_db(self) -> Optional[MeetingsDatabase]:
        if not self.db_path.is_file():
            return None
        if self._db is None:
            self._db = MeetingsDatabase(self.db_path)
        return self._db

    def invalidate_cache(self) -> None:
        self._cache_ts = 0.0

    def get_meeting_dir_for_row(self, meeting_row: dict) -> Optional[Path]:
        """Returns the local meeting directory `<cache_dir>/<folder_name>` for a meeting row."""
        explicit = meeting_row.get("meeting_dir")
        if explicit:
            return Path(explicit)
        folder_name = str(
            meeting_row.get("folder_name") or meeting_row.get("meeting_number") or ""
        ).strip()
        if not folder_name:
            return None
        return Path(self.settings.cache_dir) / folder_name

    def _get_cached_meetings(
        self, ref_date_str: str, buffer_days: int, include_upcoming: bool
    ) -> Tuple[List[dict], List[dict]]:
        now_mono = time.monotonic()
        key = (ref_date_str, buffer_days, include_upcoming)
        if (now_mono - self._cache_ts) < 60.0 and key == self._cache_key:
            return self._cached_active_rows, self._cached_upcoming_rows

        db = self._get_db()
        if db is None:
            self._cached_active_rows = []
            self._cached_upcoming_rows = []
        else:
            try:
                self._cached_active_rows = db.get_active_meetings_for_date(
                    ref_date_str, buffer_days=buffer_days, wg_filter=None
                )
                if include_upcoming:
                    self._cached_upcoming_rows = db.get_next_upcoming_meetings(
                        ref_date_str, wg_filter=None
                    )
                else:
                    self._cached_upcoming_rows = []
            except Exception as e:
                logger.debug(f"[WorldClockService] DB query failed: {e}")
                self._cached_active_rows = []
                self._cached_upcoming_rows = []

        self._cache_key = key
        self._cache_ts = now_mono
        return self._cached_active_rows, self._cached_upcoming_rows

    def get_available_working_groups(self) -> List[str]:
        """Returns known Working Group names from DB (with standard 3GPP fallback)."""
        db = self._get_db()
        wgs: List[str] = []
        if db is not None:
            try:
                wgs = db.get_working_groups()
            except Exception:
                wgs = []
        if wgs:
            return wgs
        return [
            "SA", "SA1", "SA2", "SA3", "SA4", "SA5", "SA6",
            "RAN", "RAN1", "RAN2", "RAN3", "RAN4",
            "CT", "CT1", "CT3", "CT4", "CT6",
        ]

    def resolve_venue_timezone(
        self,
        location: str,
        meetings: Optional[List[dict]] = None,
        ref_utc: Optional[datetime] = None,
        explicit_meeting_dir: Optional[Path] = None,
    ) -> Tuple[Optional[int], str, str, str]:
        """
        Resolves a venue's UTC offset using the 3-tier priority hierarchy:
        1. Manual override in `<meeting_dir>/meeting_timezone.json`
        2. Manual override in `meetings_config.json` (`venue_overrides`)
        3. Fresh `web_api` verification in `<meeting_dir>/meeting_timezone.json` or `meetings_config.json`
        4. Other confirmed `meeting_timezone.json` (`meeting_folder` / `agenda`)
        5. Offline built-in 3GPP city/country lookup (`zoneinfo` + DST rules)
        6. Stale `web_api` fallback if built-in lookup returned None
        Returns `(offset_minutes, tz_label, iana_tz, source)`.
        """
        clock_cfg = self.settings.get_clock_config()
        overrides = clock_cfg.get("venue_overrides") or {}
        loc_key = (location or "").strip().lower()
        global_entry = overrides.get(loc_key) if isinstance(overrides, dict) else None

        # Collect existing meeting folder timezone configs
        folder_entries: List[dict] = []
        candidate_dirs: List[Path] = []
        if explicit_meeting_dir is not None:
            candidate_dirs.append(Path(explicit_meeting_dir))
        for m in meetings or []:
            d = self.get_meeting_dir_for_row(m)
            if d is not None:
                candidate_dirs.append(d)

        seen_dirs = set()
        for d in candidate_dirs:
            d_key = str(d)
            if d_key in seen_dirs:
                continue
            seen_dirs.add(d_key)
            loaded = SessionPlanTimezone.load_meeting_folder_timezone(d)
            if loaded is not None:
                folder_entries.append(loaded)

        # Tier 1a: Folder manual_override
        for fe in folder_entries:
            if fe.get("source") == "manual_override":
                return (
                    int(fe["tz_offset_minutes"]),
                    fe.get("tz_label") or format_utc_offset(int(fe["tz_offset_minutes"])),
                    fe.get("iana_tz", ""),
                    "manual_override",
                )

        # Tier 1b: Global manual_override
        if isinstance(global_entry, dict) and global_entry.get("source") == "manual_override":
            if global_entry.get("tz_offset_minutes") is not None:
                off = int(global_entry["tz_offset_minutes"])
                return (
                    off,
                    str(global_entry.get("tz_label") or format_utc_offset(off)),
                    str(global_entry.get("iana_tz", "")),
                    "manual_override",
                )

        # Tier 2a: Fresh folder web_api
        for fe in folder_entries:
            if fe.get("source") == "web_api" and _is_web_entry_fresh(fe.get("updated_at", "")):
                off = int(fe["tz_offset_minutes"])
                return (
                    off,
                    fe.get("tz_label") or format_utc_offset(off),
                    fe.get("iana_tz", ""),
                    "web_api",
                )

        # Tier 2b: Fresh global web_api
        if (
            isinstance(global_entry, dict)
            and global_entry.get("source") == "web_api"
            and global_entry.get("tz_offset_minutes") is not None
            and _is_web_entry_fresh(str(global_entry.get("updated_at", "")))
        ):
            off = int(global_entry["tz_offset_minutes"])
            return (
                off,
                str(global_entry.get("tz_label") or format_utc_offset(off)),
                str(global_entry.get("iana_tz", "")),
                "web_api",
            )

        # Tier 2c: Non-manual meeting folder config (e.g. agenda/session_plan)
        for fe in folder_entries:
            if fe.get("source") not in ("web_api",):
                off = int(fe["tz_offset_minutes"])
                return (
                    off,
                    fe.get("tz_label") or format_utc_offset(off),
                    fe.get("iana_tz", ""),
                    "meeting_folder",
                )

        # Tier 3: Offline built-in city/country + zoneinfo
        ref_date_str = (ref_utc or datetime.now(timezone.utc)).strftime("%Y-%m-%d")
        builtin = SessionPlanTimezone.lookup_city_timezone_details(
            location, meeting_date_str=ref_date_str
        )
        if builtin is not None:
            off, lbl, iana = builtin
            return (off, lbl, iana, "builtin")

        # Stale web_api fallback if built-in did not recognize the city
        for fe in folder_entries:
            if fe.get("tz_offset_minutes") is not None:
                off = int(fe["tz_offset_minutes"])
                return (off, fe.get("tz_label") or format_utc_offset(off), fe.get("iana_tz", ""), "web_api")
        if isinstance(global_entry, dict) and global_entry.get("tz_offset_minutes") is not None:
            off = int(global_entry["tz_offset_minutes"])
            return (
                off,
                str(global_entry.get("tz_label") or format_utc_offset(off)),
                str(global_entry.get("iana_tz", "")),
                "web_api",
            )

        return (None, "Unresolved Timezone", "", "unresolved")

    def _build_venue_groups(
        self,
        rows: List[dict],
        ref_utc: datetime,
        laptop_dt: datetime,
        laptop_offset_min: int,
        wg_filter_upper: set,
        is_upcoming: bool,
    ) -> Tuple[List[VenueClockState], List[VenueClockState]]:
        grouped: Dict[str, List[dict]] = {}
        display_names: Dict[str, str] = {}
        for r in rows:
            loc = str(r.get("location", "")).strip()
            if not loc:
                continue
            key = loc.lower()
            grouped.setdefault(key, []).append(r)
            if key not in display_names:
                display_names[key] = loc

        primary: List[VenueClockState] = []
        other: List[VenueClockState] = []

        for key, mtgs in grouped.items():
            loc_display = display_names[key]
            wgs = sorted({str(m.get("wg_name", "")).strip() for m in mtgs if m.get("wg_name")})
            matching_wgs = (
                [w for w in wgs if w.upper() in wg_filter_upper] if wg_filter_upper else list(wgs)
            )
            starts = [str(m.get("start_date", "")).strip() for m in mtgs if m.get("start_date")]
            ends = [str(m.get("end_date", "")).strip() for m in mtgs if m.get("end_date")]
            start_d = min(starts) if starts else ""
            end_d = max(ends) if ends else ""

            eval_utc = ref_utc
            if is_upcoming and start_d:
                try:
                    d_obj = datetime.strptime(start_d[:10], "%Y-%m-%d")
                    eval_utc = d_obj.replace(hour=12, tzinfo=timezone.utc)
                except Exception:
                    pass

            off_min, tz_label, iana_tz, source = self.resolve_venue_timezone(
                loc_display, meetings=mtgs, ref_utc=eval_utc
            )

            if off_min is not None:
                v_tz = timezone(timedelta(minutes=off_min))
                v_dt = ref_utc.astimezone(v_tz)
                day_diff = (v_dt.date() - laptop_dt.date()).days
                off_diff = off_min - laptop_offset_min
            else:
                v_dt = None
                day_diff = 0
                off_diff = 0

            state = VenueClockState(
                location=loc_display,
                wgs=wgs,
                matching_wgs=matching_wgs,
                meetings=mtgs,
                tz_offset_minutes=off_min,
                tz_label=tz_label,
                iana_tz=iana_tz,
                source=source,
                venue_dt=v_dt,
                day_diff_vs_laptop=day_diff,
                offset_diff_vs_laptop_min=off_diff,
                is_upcoming=is_upcoming,
                start_date=start_d,
                end_date=end_d,
            )

            if wg_filter_upper and not matching_wgs:
                other.append(state)
            else:
                primary.append(state)

        return primary, other

    def get_snapshot(
        self,
        ref_utc: Optional[datetime] = None,
        ref_laptop_offset_min: Optional[int] = None,
    ) -> WorldClockSnapshot:
        """Builds a complete WorldClockSnapshot for the current instant (or provided `ref_utc`)."""
        utc_dt = ref_utc or datetime.now(timezone.utc)
        if utc_dt.tzinfo is None:
            utc_dt = utc_dt.replace(tzinfo=timezone.utc)
        else:
            utc_dt = utc_dt.astimezone(timezone.utc)

        laptop_off = (
            int(ref_laptop_offset_min)
            if ref_laptop_offset_min is not None
            else SessionPlanTimezone.get_laptop_tz_offset_minutes()
        )
        laptop_dt = utc_dt.astimezone(timezone(timedelta(minutes=laptop_off)))
        laptop_label = f"Laptop Local ({format_utc_offset(laptop_off)})"

        cfg = self.settings.get_clock_config()
        wg_filter = [str(w).strip() for w in (cfg.get("wg_filter") or []) if str(w).strip()]
        wg_filter_upper = {w.upper() for w in wg_filter}
        buffer_days = int(cfg.get("travel_buffer_days", 2))
        show_upcoming = bool(cfg.get("show_upcoming_when_idle", True))
        show_utc = bool(cfg.get("show_utc_in_bar", True))

        ref_date_str = utc_dt.strftime("%Y-%m-%d")
        active_rows, upcoming_rows = self._get_cached_meetings(
            ref_date_str, buffer_days=buffer_days, include_upcoming=show_upcoming
        )

        primary_venues: List[VenueClockState] = []
        other_venues: List[VenueClockState] = []

        if active_rows:
            p_act, o_act = self._build_venue_groups(
                active_rows, utc_dt, laptop_dt, laptop_off, wg_filter_upper, is_upcoming=False
            )
            primary_venues.extend(p_act)
            other_venues.extend(o_act)

        # If WG filter matched zero active venues (or there are zero active meetings overall),
        # check upcoming meetings when show_upcoming is enabled.
        if not primary_venues and show_upcoming and upcoming_rows:
            if wg_filter_upper:
                # Filter upcoming specifically for the user's selected WGs
                db = self._get_db()
                filtered_upcoming = []
                if db is not None:
                    try:
                        filtered_upcoming = db.get_next_upcoming_meetings(
                            ref_date_str, wg_filter=list(wg_filter_upper)
                        )
                    except Exception:
                        filtered_upcoming = []
                if filtered_upcoming:
                    p_up, _ = self._build_venue_groups(
                        filtered_upcoming, utc_dt, laptop_dt, laptop_off, wg_filter_upper, is_upcoming=True
                    )
                    primary_venues.extend(p_up)
            else:
                p_up, _ = self._build_venue_groups(
                    upcoming_rows, utc_dt, laptop_dt, laptop_off, wg_filter_upper, is_upcoming=True
                )
                primary_venues.extend(p_up)

        return WorldClockSnapshot(
            utc_dt=utc_dt,
            laptop_dt=laptop_dt,
            laptop_offset_minutes=laptop_off,
            laptop_label=laptop_label,
            primary_venues=primary_venues,
            other_venues=other_venues,
            wg_filter=wg_filter,
            show_utc_in_bar=show_utc,
            show_upcoming_when_idle=show_upcoming,
        )

    def resolve_single_meeting(
        self,
        mtg_info: dict,
        meeting_dir: Optional[Path] = None,
        ref_utc: Optional[datetime] = None,
    ) -> VenueClockState:
        """Resolves the live clock state for a single meeting (used by TDocsWindow)."""
        utc_dt = ref_utc or datetime.now(timezone.utc)
        laptop_off = SessionPlanTimezone.get_laptop_tz_offset_minutes()
        laptop_dt = utc_dt.astimezone(timezone(timedelta(minutes=laptop_off)))
        location = str(mtg_info.get("location", "")).strip()
        wg_name = str(mtg_info.get("wg_name", "")).strip()
        is_electronic = bool(mtg_info.get("is_electronic", 0))

        if is_electronic or not location or location.upper() in MeetingsDatabase._NON_PHYSICAL_LOCATIONS:
            return VenueClockState(
                location="Online (UTC)",
                wgs=[wg_name] if wg_name else [],
                meetings=[mtg_info],
                tz_offset_minutes=0,
                tz_label="UTC+00:00 (Electronic Meeting)",
                iana_tz="UTC",
                source="builtin",
                venue_dt=utc_dt,
                day_diff_vs_laptop=(utc_dt.date() - laptop_dt.date()).days,
                offset_diff_vs_laptop_min=-laptop_off,
                start_date=str(mtg_info.get("start_date", "")),
                end_date=str(mtg_info.get("end_date", "")),
            )

        off_min, tz_label, iana_tz, source = self.resolve_venue_timezone(
            location,
            meetings=[mtg_info],
            ref_utc=utc_dt,
            explicit_meeting_dir=meeting_dir,
        )
        if off_min is not None:
            v_dt = utc_dt.astimezone(timezone(timedelta(minutes=off_min)))
            day_diff = (v_dt.date() - laptop_dt.date()).days
            off_diff = off_min - laptop_off
        else:
            v_dt = None
            day_diff = 0
            off_diff = 0

        return VenueClockState(
            location=location,
            wgs=[wg_name] if wg_name else [],
            meetings=[mtg_info],
            tz_offset_minutes=off_min,
            tz_label=tz_label,
            iana_tz=iana_tz,
            source=source,
            venue_dt=v_dt,
            day_diff_vs_laptop=day_diff,
            offset_diff_vs_laptop_min=off_diff,
            start_date=str(mtg_info.get("start_date", "")),
            end_date=str(mtg_info.get("end_date", "")),
        )
