# --- File: tests/test_world_clock.py ---
"""Unit tests for the 3GPP Multi-Venue & World Clock system."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication, QPushButton, QWidget

from core.ui.system_status_bar import SystemStatusBar
from core.ui.world_clock_dialog import WorldClockDialog
from core.ui.world_clock_service import WebTimezoneSyncWorker, WorldClockService
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.session_plan.session_plan_timezone import (
    SessionPlanTimezone,
    format_utc_offset,
)
from modules.meetings.core.settings import MeetingsSettings


class TestWorldClockAndTimezones(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="test_world_clock_"))
        self.db_path = self.tmp_dir / "3gpp_data.db"
        self.cfg_path = self.tmp_dir / "meetings_config.json"
        self.cache_dir = self.tmp_dir / "cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        self.db = MeetingsDatabase(self.db_path)
        self.settings = MeetingsSettings(config_file=self.cfg_path)
        self.settings.save_settings(str(self.cache_dir))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _seed_meeting(
        self,
        wg: str,
        num: str,
        location: str,
        start_date: str,
        end_date: str,
        folder_name: str = "",
        is_electronic: int = 0,
    ) -> None:
        self.db.upsert_single_meeting(
            {
                "wg_name": wg,
                "meeting_number": num,
                "folder_name": folder_name or f"{wg}_{num}",
                "location": location,
                "start_date": start_date,
                "end_date": end_date,
                "is_electronic": is_electronic,
                "url_key": f"Meetings_3GPP_SYNC/{wg}/{num}",
                "mtg_id": f"{wg}-{num}",
            }
        )

    def test_all_2025_2027_3gpp_locations_resolve(self) -> None:
        locations = [
            "Calgary",
            "Boston",
            "Baltimore",
            "Montreal",
            "Osaka",
            "Incheon, KR",
            "South Korea",
            "Prague",
            "Maastricht",
            "Madrid",
            "Singapore",
            "China",
            "Wuhan",
            "Saint Julian's",
            "Saint Paul's Bay",
            "Valletta",
            "Gothenburg Metropolitan Area",
            "Stor-Göteborg",
            "India",
            "Bengaluru",
            "Fukuoka",
            "Dallas",
            "Athens",
            "Geneva",
            "Sophia-Antipolis",
            "Bratislava",
        ]
        for loc in locations:
            res = SessionPlanTimezone.lookup_city_timezone(loc, "2026-11-16")
            self.assertIsNotNone(res, msg=f"Location '{loc}' failed to resolve")
            off, lbl = res
            self.assertIsInstance(off, int)
            self.assertTrue(lbl.startswith("UTC"), msg=f"Unexpected label '{lbl}' for '{loc}'")

        # Negative half-hour offset formatting check
        self.assertEqual(format_utc_offset(-330), "UTC-05:30")
        self.assertEqual(format_utc_offset(330), "UTC+05:30")
        self.assertEqual(format_utc_offset(0), "UTC+00:00")

    def test_colocated_grouping_vs_split_week_and_wg_filter(self) -> None:
        # Co-located Calgary week (2026-11-16 .. 2026-11-20)
        for wg, num in [
            ("SA1", "116"), ("SA2", "178"), ("SA3", "131"),
            ("RAN1", "127"), ("RAN2", "136"), ("CT1", "164"),
        ]:
            self._seed_meeting(wg, num, "Calgary", "2026-11-16", "2026-11-20")

        # Split week (2026-10-12 .. 2026-10-16): SA2/CT1 in Prague, RAN2/RAN3 in South Korea
        self._seed_meeting("SA2", "177", "Prague", "2026-10-12", "2026-10-16")
        self._seed_meeting("CT1", "163", "Prague", "2026-10-12", "2026-10-16")
        self._seed_meeting("RAN2", "135BIS", "South Korea", "2026-10-12", "2026-10-16")
        self._seed_meeting("RAN3", "133", "South Korea", "2026-10-12", "2026-10-16")

        # Electronic meeting overlapping split week must be ignored
        self._seed_meeting("SA4", "135E", "Online", "2026-10-12", "2026-10-16", is_electronic=1)

        service = WorldClockService(db_path=self.db_path, settings=self.settings)

        # 1. Co-located Calgary: 6 meetings collapse into 1 Calgary venue
        snap_calgary = service.get_snapshot(
            ref_utc=datetime(2026, 11, 16, 12, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        self.assertEqual(len(snap_calgary.primary_venues), 1)
        self.assertEqual(snap_calgary.primary_venues[0].location, "Calgary")
        self.assertEqual(len(snap_calgary.primary_venues[0].wgs), 6)

        # 2. Split week with All Active WGs: shows both Prague and South Korea
        snap_split_all = service.get_snapshot(
            ref_utc=datetime(2026, 10, 14, 10, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=120,
        )
        locs = {v.location for v in snap_split_all.primary_venues}
        self.assertEqual(locs, {"Prague", "South Korea"})
        self.assertEqual(snap_split_all.other_venues, [])

        # 3. Split week filtered to SA2 only: Prague primary, South Korea in other_venues
        cfg = self.settings.get_clock_config()
        cfg["wg_filter"] = ["SA2"]
        self.settings.save_clock_config(cfg)
        service.invalidate_cache()

        snap_split_sa2 = service.get_snapshot(
            ref_utc=datetime(2026, 10, 14, 10, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=120,
        )
        self.assertEqual([v.location for v in snap_split_sa2.primary_venues], ["Prague"])
        self.assertEqual([v.location for v in snap_split_sa2.other_venues], ["South Korea"])
        bar_txt = service.format_status_bar_text(snap_split_sa2)
        self.assertIn("Prague", bar_txt)
        self.assertIn("(SA2)", bar_txt)
        self.assertNotIn("South Korea", bar_txt)

        compact_bar_txt = service.format_status_bar_text(snap_split_sa2, compact=True)
        self.assertIn("Prague", compact_bar_txt)
        self.assertIn("Local", compact_bar_txt)
        self.assertNotIn("(SA2)", compact_bar_txt)

    def test_meeting_folder_persistence_and_override_priority(self) -> None:
        folder = "TSGS2_178_Calgary_2026-11"
        mtg_dir = self.cache_dir / folder
        mtg_dir.mkdir(parents=True, exist_ok=True)
        self._seed_meeting("SA2", "178", "Calgary", "2026-11-16", "2026-11-20", folder_name=folder)

        service = WorldClockService(db_path=self.db_path, settings=self.settings)

        # Save web_api first
        self.settings.set_venue_timezone_override(
            location="Calgary",
            tz_offset_minutes=-360,
            tz_label="UTC-06:00 (America/Edmonton)",
            source="web_api",
            iana_tz="America/Edmonton",
            matching_meetings=[{"folder_name": folder, "wg_name": "SA2", "meeting_number": "178", "location": "Calgary"}],
        )
        tz_file = mtg_dir / "meeting_timezone.json"
        self.assertTrue(tz_file.is_file())
        data = json.loads(tz_file.read_text(encoding="utf-8"))
        self.assertEqual(data["source"], "web_api")
        self.assertEqual(data["tz_offset_minutes"], -360)

        # Save manual_override -> must win over web_api and builtin
        self.settings.set_venue_timezone_override(
            location="Calgary",
            tz_offset_minutes=-420,
            tz_label="UTC-07:00 (Manual Override)",
            source="manual_override",
            matching_meetings=[{"folder_name": folder, "wg_name": "SA2", "meeting_number": "178", "location": "Calgary"}],
        )
        snap = service.get_snapshot(
            ref_utc=datetime(2026, 11, 16, 12, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        self.assertEqual(snap.primary_venues[0].source, "manual_override")
        self.assertEqual(snap.primary_venues[0].tz_offset_minutes, -420)

        # Background web_api update must NOT overwrite locked manual_override
        self.settings.set_venue_timezone_override(
            location="Calgary",
            tz_offset_minutes=-360,
            tz_label="UTC-06:00 (America/Edmonton)",
            source="web_api",
            matching_meetings=[{"folder_name": folder, "wg_name": "SA2", "meeting_number": "178", "location": "Calgary"}],
        )
        data_after = json.loads(tz_file.read_text(encoding="utf-8"))
        self.assertEqual(data_after["source"], "manual_override")
        self.assertEqual(data_after["tz_offset_minutes"], -420)

        # Clear override -> reverts to builtin
        self.settings.clear_venue_timezone_override(
            location="Calgary",
            matching_meetings=[{"folder_name": folder}],
        )
        self.assertFalse(tz_file.exists())
        service.invalidate_cache()
        snap_reset = service.get_snapshot(
            ref_utc=datetime(2026, 11, 16, 12, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        self.assertEqual(snap_reset.primary_venues[0].source, "builtin")

    def test_day_rollover_indicators_and_offweek_upcoming(self) -> None:
        self._seed_meeting("SA2", "178", "Calgary", "2026-11-16", "2026-11-20")
        self._seed_meeting("RAN2", "137", "Osaka", "2026-11-16", "2026-11-20")
        self._seed_meeting("SA", "114", "Boston", "2026-12-07", "2026-12-11")

        service = WorldClockService(db_path=self.db_path, settings=self.settings)

        # At 23:30 UTC on Nov 16 -> Laptop (UTC+1) is 00:30 on Nov 17
        # Osaka (UTC+9) is 08:30 on Nov 17 (day_diff = 0)
        # Calgary (UTC-6 or -7) is 16:30/17:30 on Nov 16 (day_diff = -1)
        snap = service.get_snapshot(
            ref_utc=datetime(2026, 11, 16, 23, 30, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        by_loc = {v.location: v for v in snap.primary_venues}
        self.assertEqual(by_loc["Calgary"].day_diff_vs_laptop, -1)
        self.assertEqual(by_loc["Osaka"].day_diff_vs_laptop, 0)
        self.assertIn("(-1d)", service.format_status_bar_text(snap))

        # At 16:00 UTC on Nov 16 -> Laptop (UTC+1) is 17:00 Nov 16, Osaka (UTC+9) is 01:00 Nov 17 (+1d)
        snap_plus = service.get_snapshot(
            ref_utc=datetime(2026, 11, 16, 16, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        by_loc_plus = {v.location: v for v in snap_plus.primary_venues}
        self.assertEqual(by_loc_plus["Osaka"].day_diff_vs_laptop, 1)
        self.assertIn("(+1d)", service.format_status_bar_text(snap_plus))

        # Off-week fallback (Nov 28, 2026 -> Next upcoming is Boston Dec 7)
        snap_off = service.get_snapshot(
            ref_utc=datetime(2026, 11, 28, 12, 0, tzinfo=timezone.utc),
            ref_laptop_offset_min=60,
        )
        self.assertEqual(len(snap_off.primary_venues), 1)
        self.assertEqual(snap_off.primary_venues[0].location, "Boston")
        self.assertTrue(snap_off.primary_venues[0].is_upcoming)
        self.assertIn("⏭️", service.format_status_bar_text(snap_off))

    def test_status_bar_and_world_clock_dialog_smoke(self) -> None:
        # Anchor dates to current UTC date so live snapshots without ref_utc remain deterministic
        today_utc = datetime.now(timezone.utc).date()
        start_d = (today_utc - timedelta(days=1)).isoformat()
        end_d = (today_utc + timedelta(days=3)).isoformat()
        self._seed_meeting("SA2", "178", "Calgary", start_d, end_d)
        service = WorldClockService(db_path=self.db_path, settings=self.settings)

        bar = SystemStatusBar()
        try:
            bar._world_clock_service = service
            bar.refresh_world_clock()
            self.assertIn("Local", bar.clock_indicator.text())
            self.assertIn("UTC", bar.clock_indicator.text())
            self.assertIn("Calgary", bar.clock_indicator.text())
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()

        dlg = WorldClockDialog(service)
        try:
            self.assertGreaterEqual(len(dlg._venue_cards), 1)
            # Nudge +1h and verify override persists
            dlg._on_nudge_offset("Calgary", 60, dlg._venue_cards[0][2].venue.meetings, None)
            cfg = self.settings.get_clock_config()
            self.assertIn("calgary", cfg["venue_overrides"])
            self.assertEqual(cfg["venue_overrides"]["calgary"]["source"], "manual_override")
        finally:
            dlg.close()

    def test_web_timezone_sync_worker_city_and_ip_mocked(self) -> None:
        def _fake_get(url: str, **_kwargs):
            resp = MagicMock()
            resp.ok = True
            resp.raise_for_status = MagicMock()
            if "geocoding-api.open-meteo.com" in url:
                resp.json.return_value = {"results": [{"name": "Calgary", "timezone": "America/Edmonton"}]}
            elif "timeapi.io/api/time/current/zone" in url:
                resp.json.return_value = {"currentUtcOffset": {"seconds": -25200}, "dstActive": False}
            elif "ipwho.is" in url:
                resp.json.return_value = {
                    "success": True,
                    "city": "Calgary",
                    "country": "Canada",
                    "timezone": {"id": "America/Edmonton", "offset": -25200},
                }
            else:
                resp.ok = False
                resp.json.return_value = {}
            return resp

        fake_session = MagicMock()
        fake_session.get.side_effect = _fake_get

        with patch("core.ui.world_clock_service.NetworkSession.get_instance", return_value=fake_session):
            city_results = []
            w_city = WebTimezoneSyncWorker(mode="city", location="Calgary")
            w_city.city_resolved.connect(lambda loc, off, lbl, iana, mtgs: city_results.append((loc, off, lbl, iana)))
            w_city.run()
            self.assertEqual(len(city_results), 1)
            self.assertEqual(city_results[0][0], "Calgary")
            self.assertEqual(city_results[0][1], -420)
            self.assertIn("America/Edmonton", city_results[0][2])
            self.assertEqual(city_results[0][3], "America/Edmonton")

            ip_results = []
            w_ip = WebTimezoneSyncWorker(mode="ip")
            w_ip.ip_resolved.connect(lambda place, off, lbl, iana: ip_results.append((place, off, lbl, iana)))
            w_ip.run()
            self.assertEqual(len(ip_results), 1)
            self.assertEqual(ip_results[0][0], "Calgary, Canada")
            self.assertEqual(ip_results[0][1], -420)
            self.assertEqual(ip_results[0][3], "America/Edmonton")

    def test_tdocs_window_contextual_venue_clock(self) -> None:
        from modules.meetings.ui.tdocs_window import TDocsWindow

        folder = "TSGS2_178_Calgary_2026-11"
        mtg_dir = self.cache_dir / folder
        mtg_dir.mkdir(parents=True, exist_ok=True)
        SessionPlanTimezone.save_meeting_folder_timezone(
            meeting_dir=mtg_dir,
            tz_offset_minutes=-420,
            tz_label="UTC-07:00 (Manual Override)",
            source="manual_override",
            location="Calgary",
            wg="SA2",
            meeting_number="178",
        )

        class _DummyTDocs(QWidget):
            _update_venue_clock = TDocsWindow._update_venue_clock

        w = _DummyTDocs()
        w.mtg_info = {"wg_name": "SA2", "meeting_number": "178", "location": "Calgary", "is_electronic": 0}
        w.meeting_dir = mtg_dir
        w._clock_service = WorldClockService(db_path=self.db_path, settings=self.settings)
        w.venue_clock_btn = QPushButton(w)

        w._update_venue_clock()
        self.assertIn("Calgary", w.venue_clock_btn.text())
        self.assertIn("UTC-07", w.venue_clock_btn.text())
        self.assertIn("manual_override", w.venue_clock_btn.toolTip())
        w.close()


if __name__ == "__main__":
    unittest.main()
