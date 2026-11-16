# --- File: modules/meetings/core/settings.py ---
from datetime import datetime, timezone
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional
import core.utils.paths

logger = logging.getLogger(__name__)

DEFAULT_CLOCK_CONFIG: Dict = {
    "wg_filter": [],
    "travel_buffer_days": 2,
    "show_upcoming_when_idle": True,
    "show_utc_in_bar": True,
    "venue_overrides": {},
}

class MeetingsSettings:
    def __init__(self, config_file: Optional[Path] = None):
        self.config_file = (
            Path(config_file)
            if config_file is not None
            else (core.utils.paths.get_project_root() / "config" / "meetings_config.json")
        )
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        self.cache_dir = self._load_settings()

    def _load_settings(self) -> str:
        fallback = str(core.utils.paths.get_cache_root())
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('download_dir', fallback)
            except Exception:
                pass
        return fallback

    def save_settings(self, download_dir: str):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['download_dir'] = download_dir

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            self.cache_dir = download_dir
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving config: {e}")

    def save_last_meeting(self, mtg_info: dict):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['last_mtg_id'] = mtg_info.get("mtg_id")
            data['last_mtg_number'] = mtg_info.get("meeting_number")
            data['last_mtg_wg'] = mtg_info.get("wg_name") or mtg_info.get("wg", "")

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving last meeting state: {e}")

    def get_last_meeting(self) -> tuple:
        """Returns (last_id, last_num, last_wg). Returns (None, None, None) if missing."""
        if not self.config_file.exists():
            return None, None, None
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("last_mtg_id"), data.get("last_mtg_number"), data.get("last_mtg_wg")
        except Exception:
            return None, None, None

    def save_filters(self, filters: dict):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['filters'] = filters

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving filters: {e}")

    def get_filters(self) -> dict:
        if not self.config_file.exists():
            return {}
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("filters", {})
        except Exception:
            return {}

    def get_include_unassigned_ls_in(self) -> bool:
        if not self.config_file.exists():
            return True
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("include_unassigned_ls_in", True)
        except Exception:
            return True

    def set_include_unassigned_ls_in(self, enabled: bool):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['include_unassigned_ls_in'] = bool(enabled)

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving include_unassigned_ls_in: {e}")

    def get_auto_enrich_wis(self) -> bool:
        if not self.config_file.exists():
            return True
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("auto_enrich_wis", True)
        except Exception:
            return True

    def set_auto_enrich_wis(self, enabled: bool):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['auto_enrich_wis'] = bool(enabled)

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving auto_enrich_wis: {e}")

    def get_include_tentative_sources(self) -> bool:
        if not self.config_file.exists():
            return False
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return bool(data.get("include_tentative_sources", False))
        except Exception:
            return False

    def set_include_tentative_sources(self, enabled: bool):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['include_tentative_sources'] = bool(enabled)

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving include_tentative_sources: {e}")

    def _read_raw_config(self) -> dict:
        if not self.config_file.exists():
            return {}
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _write_raw_config(self, data: dict) -> None:
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)

    def get_clock_config(self) -> dict:
        """Returns persisted World Clock & Venue Timezone configuration with defaults."""
        data = self._read_raw_config()
        raw_clock = data.get("clock_config", {})
        if not isinstance(raw_clock, dict):
            raw_clock = {}
        cfg = {
            "wg_filter": list(raw_clock.get("wg_filter") or []),
            "travel_buffer_days": int(raw_clock.get("travel_buffer_days", 2)),
            "show_upcoming_when_idle": bool(raw_clock.get("show_upcoming_when_idle", True)),
            "show_utc_in_bar": bool(raw_clock.get("show_utc_in_bar", True)),
            "venue_overrides": dict(raw_clock.get("venue_overrides") or {}),
        }
        return cfg

    def save_clock_config(self, clock_cfg: dict) -> None:
        """Persists World Clock & Venue Timezone configuration into meetings_config.json."""
        try:
            data = self._read_raw_config()
            data["clock_config"] = {
                "wg_filter": list(clock_cfg.get("wg_filter") or []),
                "travel_buffer_days": int(clock_cfg.get("travel_buffer_days", 2)),
                "show_upcoming_when_idle": bool(clock_cfg.get("show_upcoming_when_idle", True)),
                "show_utc_in_bar": bool(clock_cfg.get("show_utc_in_bar", True)),
                "venue_overrides": dict(clock_cfg.get("venue_overrides") or {}),
            }
            self._write_raw_config(data)
        except Exception as e:
            logger.error(f"[MeetingsSettings] Error saving clock_config: {e}")

    def _resolve_existing_meeting_dirs(self, matching_meetings: Optional[List[dict]]) -> List[tuple]:
        """Returns [(Path, wg, meeting_number, location), ...] for existing meeting folders."""
        if not matching_meetings:
            return []
        base = Path(self.cache_dir)
        found: List[tuple] = []
        seen_paths = set()
        for m in matching_meetings:
            explicit_dir = m.get("meeting_dir")
            if explicit_dir:
                mtg_dir = Path(explicit_dir)
            else:
                folder_name = str(m.get("folder_name") or m.get("meeting_number") or "").strip()
                if not folder_name:
                    continue
                mtg_dir = base / folder_name
            key = str(mtg_dir.resolve()) if mtg_dir.exists() else str(mtg_dir)
            if mtg_dir.is_dir() and key not in seen_paths:
                seen_paths.add(key)
                found.append((
                    mtg_dir,
                    str(m.get("wg_name") or m.get("wg") or ""),
                    str(m.get("meeting_number") or ""),
                    str(m.get("location") or ""),
                ))
        return found

    def set_venue_timezone_override(
        self,
        location: str,
        tz_offset_minutes: int,
        tz_label: str,
        source: str = "manual_override",
        iana_tz: str = "",
        matching_meetings: Optional[List[dict]] = None,
    ) -> None:
        """
        Saves a venue timezone override/verification into `meetings_config.json`
        and writes `meeting_timezone.json` into any existing matching meeting folders.
        """
        loc_key = (location or "").strip().lower()
        if not loc_key:
            return

        cfg = self.get_clock_config()
        overrides = dict(cfg.get("venue_overrides") or {})
        existing = overrides.get(loc_key)
        # Never let automatic web_api overwrite a locked manual_override
        if (
            source != "manual_override"
            and isinstance(existing, dict)
            and existing.get("source") == "manual_override"
        ):
            return

        overrides[loc_key] = {
            "location": location.strip(),
            "tz_offset_minutes": int(tz_offset_minutes),
            "tz_label": str(tz_label),
            "iana_tz": str(iana_tz or ""),
            "source": str(source or "manual_override"),
            "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        cfg["venue_overrides"] = overrides
        self.save_clock_config(cfg)

        from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone

        for mtg_dir, wg, m_num, m_loc in self._resolve_existing_meeting_dirs(matching_meetings):
            if source != "manual_override":
                cur_folder = SessionPlanTimezone.load_meeting_folder_timezone(mtg_dir)
                if cur_folder and cur_folder.get("source") == "manual_override":
                    continue
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=mtg_dir,
                tz_offset_minutes=int(tz_offset_minutes),
                tz_label=str(tz_label),
                source=source,
                location=m_loc or location.strip(),
                wg=wg,
                meeting_number=m_num,
                iana_tz=iana_tz,
            )

    def clear_venue_timezone_override(
        self,
        location: str,
        matching_meetings: Optional[List[dict]] = None,
    ) -> None:
        """Removes a venue timezone override from config and matching meeting folders."""
        loc_key = (location or "").strip().lower()
        if loc_key:
            cfg = self.get_clock_config()
            overrides = dict(cfg.get("venue_overrides") or {})
            if loc_key in overrides:
                overrides.pop(loc_key, None)
                cfg["venue_overrides"] = overrides
                self.save_clock_config(cfg)

        from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone

        for mtg_dir, _wg, _m_num, _m_loc in self._resolve_existing_meeting_dirs(matching_meetings):
            SessionPlanTimezone.clear_meeting_folder_timezone(mtg_dir)


