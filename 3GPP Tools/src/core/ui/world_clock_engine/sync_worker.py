"""Asynchronous QThread worker for public city/IP timezone synchronization."""

from __future__ import annotations

import logging
import urllib.parse
from typing import List, Optional

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.ui.world_clock_engine.models import _clean_city_query
from modules.meetings.core.session_plan.session_plan_timezone import (
    SessionPlanTimezone,
    format_utc_offset,
)

logger = logging.getLogger(__name__)


class WebTimezoneSyncWorker(QThread):
    """
    Asynchronous worker that queries public geocoding/timezone APIs using
    `NetworkSession` (respecting corporate proxies) without blocking the GUI.
    """

    city_resolved = pyqtSignal(str, int, str, str, list)  # location, offset_min, label, iana_tz, meetings
    ip_resolved = pyqtSignal(str, int, str, str)  # city, offset_min, label, iana_tz
    sync_failed = pyqtSignal(str, str)  # target, error_msg

    def __init__(
        self,
        mode: str = "city",
        location: str = "",
        meetings: Optional[List[dict]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.mode = mode
        self.location = location
        self.meetings = list(meetings or [])

    def run(self) -> None:
        try:
            if self.mode == "ip":
                self._run_ip_lookup()
            else:
                self._run_city_lookup()
        except Exception as e:
            if not self.isInterruptionRequested():
                logger.debug(f"[WorldClockSync] Sync failed for {self.location or 'IP'}: {e}")
                self.sync_failed.emit(self.location or "ip", str(e))

    def _run_city_lookup(self) -> None:
        query = _clean_city_query(self.location)
        if not query:
            self.sync_failed.emit(self.location, "Empty location name")
            return

        session = NetworkSession.get_instance()
        headers = NetworkSession.get_request_headers()

        # 1. Geocode city -> IANA timezone via Open-Meteo
        geo_url = (
            "https://geocoding-api.open-meteo.com/v1/search"
            f"?name={urllib.parse.quote(query)}&count=1&language=en&format=json"
        )
        resp = session.get(geo_url, headers=headers, timeout=6)
        resp.raise_for_status()
        geo_data = resp.json() or {}
        results = geo_data.get("results") or []
        if not results:
            raise ValueError(f"No geocoding match found for '{query}'")

        first = results[0]
        iana_tz = str(first.get("timezone", "")).strip()
        if not iana_tz:
            raise ValueError(f"Geocoding result for '{query}' did not include a timezone")

        if self.isInterruptionRequested():
            return

        # 2. Query live UTC offset & DST status from TimeAPI.io (with WorldTimeAPI / zoneinfo fallback)
        offset_min: Optional[int] = None
        dst_active: Optional[bool] = None

        try:
            t_url = f"https://timeapi.io/api/time/current/zone?timeZone={urllib.parse.quote(iana_tz)}"
            t_resp = session.get(t_url, headers=headers, timeout=6)
            if t_resp.ok:
                t_data = t_resp.json() or {}
                cur_off = t_data.get("currentUtcOffset") or {}
                if isinstance(cur_off, dict) and cur_off.get("seconds") is not None:
                    offset_min = int(round(float(cur_off["seconds"]) / 60))
                if t_data.get("dstActive") is not None:
                    dst_active = bool(t_data.get("dstActive"))
        except Exception as e:
            logger.debug(f"[WorldClockSync] TimeAPI.io zone check failed for {iana_tz}: {e}")

        if offset_min is None and not self.isInterruptionRequested():
            try:
                wt_url = f"https://worldtimeapi.org/api/timezone/{urllib.parse.quote(iana_tz)}"
                wt_resp = session.get(wt_url, headers=headers, timeout=6)
                if wt_resp.ok:
                    wt_data = wt_resp.json() or {}
                    raw_off = wt_data.get("raw_offset")
                    dst_off = wt_data.get("dst_offset", 0)
                    if raw_off is not None:
                        offset_min = int(round((float(raw_off) + float(dst_off or 0)) / 60))
                    if wt_data.get("dst") is not None:
                        dst_active = bool(wt_data.get("dst"))
            except Exception as e:
                logger.debug(f"[WorldClockSync] WorldTimeAPI check failed for {iana_tz}: {e}")

        if offset_min is None:
            zi_res = SessionPlanTimezone.resolve_iana_offset_minutes(iana_tz)
            if zi_res is not None:
                offset_min = zi_res[0]

        if offset_min is None:
            raise ValueError(f"Could not determine UTC offset for '{iana_tz}'")

        dst_note = " - DST" if dst_active is True else ""
        label = f"{format_utc_offset(offset_min)} ({iana_tz}{dst_note})"
        if not self.isInterruptionRequested():
            self.city_resolved.emit(self.location, offset_min, label, iana_tz, self.meetings)

    def _run_ip_lookup(self) -> None:
        session = NetworkSession.get_instance()
        headers = NetworkSession.get_request_headers()

        city = ""
        country = ""
        iana_tz = ""
        offset_min: Optional[int] = None

        # 1. Primary: ipwho.is (returns city, country, IANA timezone, and UTC offset in one call)
        try:
            resp = session.get("https://ipwho.is/", headers=headers, timeout=6)
            if resp.ok:
                data = resp.json() or {}
                if data.get("success", True):
                    city = str(data.get("city", "")).strip()
                    country = str(data.get("country", "")).strip()
                    tz_obj = data.get("timezone") or {}
                    if isinstance(tz_obj, dict):
                        iana_tz = str(tz_obj.get("id", "")).strip()
                        if tz_obj.get("offset") is not None:
                            offset_min = int(round(float(tz_obj["offset"]) / 60))
        except Exception as e:
            logger.debug(f"[WorldClockSync] ipwho.is failed: {e}")

        # 2. Fallback: worldtimeapi.org/api/ip
        if (offset_min is None or not iana_tz) and not self.isInterruptionRequested():
            try:
                resp = session.get("https://worldtimeapi.org/api/ip", headers=headers, timeout=6)
                if resp.ok:
                    data = resp.json() or {}
                    iana_tz = iana_tz or str(data.get("timezone", "")).strip()
                    if data.get("raw_offset") is not None:
                        offset_min = int(
                            round((float(data["raw_offset"]) + float(data.get("dst_offset") or 0)) / 60)
                        )
            except Exception as e:
                logger.debug(f"[WorldClockSync] worldtimeapi IP check failed: {e}")

        if offset_min is None and iana_tz:
            zi_res = SessionPlanTimezone.resolve_iana_offset_minutes(iana_tz)
            if zi_res is not None:
                offset_min = zi_res[0]

        if offset_min is None:
            raise ValueError("Could not detect timezone from public IP endpoints")

        place = ", ".join(p for p in (city, country) if p) or iana_tz or "Detected Location"
        tz_part = f" ({iana_tz})" if iana_tz else ""
        label = f"{format_utc_offset(offset_min)}{tz_part}"
        if not self.isInterruptionRequested():
            self.ip_resolved.emit(place, offset_min, label, iana_tz)
