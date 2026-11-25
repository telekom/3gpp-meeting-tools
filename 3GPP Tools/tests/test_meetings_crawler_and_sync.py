"""Unit and UI tests for meeting parsing, fast-first sync order, and sync UI controls."""

from __future__ import annotations

import datetime
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PyQt5.QtWidgets import QApplication, QMainWindow

from core.ui.toast import ToastManager
from modules.meetings.core.crawler.crawler_thread import MeetingsCrawlerThread
from modules.meetings.core.crawler.dynareport import process_dynareport
from modules.meetings.core.crawler.sources import (
    extract_fallback_date_from_folder,
    extract_meeting_number,
    is_meeting_name,
    normalize_meeting_ftp_path,
)
from modules.meetings.core.db import MeetingsDatabase
from modules.meetings.ui.meetings_sync_dialog import MeetingsSyncDialog
from modules.meetings.ui.ui_tabs import MeetingsTab


class TestMeetingsParsingHelpers(unittest.TestCase):
    def test_extract_meeting_number_and_deny_list(self) -> None:
        cases = [
            ("TSGS2_178-AH-e_Electronic_2027-01", "178AHE"),
            ("TSGR3_135-bis", "135BIS"),
            ("TSGS4_127-bis-e", "127BISE"),
            ("TSGS1_112-Ad_Hoc-e", "112AHE"),
            ("TSGS3_110_AdHoc-e", "110AHE"),
            ("TSGR4_94_e", "94E"),
            ("TSGR4_94_eBis", "94EBIS"),
            ("TSGS2_178_Calgary_2026-11", "178"),
            ("TSGS1_60_Edinburgh", "60"),
            ("TSGS2_90a-SIRIG_AH", "90A"),
        ]
        for folder, expected in cases:
            self.assertEqual(extract_meeting_number(folder), expected, folder)

        for container in ("TSGR_AHs", "TSGR1_AH", "TSGR2_AHs", "TSGR3_AHGs", "TSGR4_AHs"):
            self.assertFalse(is_meeting_name(container))

    def test_fallback_date_and_ftp_path_normalization(self) -> None:
        self.assertEqual(
            extract_fallback_date_from_folder("TSGS2_178-AH-e_Electronic_2027-01"),
            "2027-01-01",
        )
        self.assertEqual(extract_fallback_date_from_folder("TSGS1_60_Edinburgh"), "")
        self.assertEqual(
            normalize_meeting_ftp_path("tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/Invitation/"),
            "tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11",
        )
        self.assertEqual(
            normalize_meeting_ftp_path("tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/Inbox/Revisions"),
            "tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11",
        )

    def test_process_dynareport_parses_ahe_ad_hoc_and_normalizes_url_key(self) -> None:
        sample_html = """
        <table>
          <tr>
            <td>S2-181</td>
            <td><a href="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=82767">3GPPSA2#181</a></td>
            <td>China</td>
            <td>2027-05-24</td>
            <td>2027-05-28</td>
          </tr>
          <tr>
            <td>S2-178-AH-e</td>
            <td><a href="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=86724">3GPPSA2#178-AH-e</a></td>
            <td>Online</td>
            <td>2027-01-25</td>
            <td>2027-01-29</td>
          </tr>
          <tr>
            <td>S2-178</td>
            <td><a href="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60769">3GPPSA2#178</a></td>
            <td>Calgary</td>
            <td>2026-11-16</td>
            <td>2026-11-20</td>
            <td><a href="/../../../\\ftp\\tsg_sa\\WG2_Arch\\TSGS2_178_Calgary_2026-11\\Invitation/">Invitation</a></td>
            <td><a href="/../../../\\ftp\\tsg_sa\\WG2_Arch\\TSGS2_178_Calgary_2026-11\\\\docs\\">Docs</a></td>
          </tr>
          <tr>
            <td>S2-166-Ad Hoc-e</td>
            <td><a href="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60815">3GPPSA2#166-Ad Hoc-e</a></td>
            <td>Online</td>
            <td>2025-01-20</td>
            <td>2025-01-24</td>
          </tr>
        </table>
        """
        with patch(
            "modules.meetings.core.crawler.dynareport.NetworkSession.get_html",
            return_value=sample_html,
        ):
            rows = process_dynareport("SA2", "https://www.3gpp.org/dynareport?code=Meetings-S2.htm")

        by_num = {r[1]: r for r in rows}
        self.assertIn("181", by_num)
        self.assertEqual(by_num["181"][2], "")
        self.assertEqual(by_num["181"][3], "82767")

        self.assertIn("178AHE", by_num)
        self.assertEqual(by_num["178AHE"][3], "86724")
        self.assertEqual(by_num["178AHE"][6], "2027-01-25")

        self.assertIn("178", by_num)
        self.assertEqual(by_num["178"][2], "tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11")

        self.assertIn("166AHE", by_num)


class TestMeetingsDbUpsertsAndQueries(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.db_path = Path(self._tmpdir.name) / "test_meetings.db"
        self.db = MeetingsDatabase(self.db_path)

    def tearDown(self) -> None:
        self._tmpdir.cleanup()

    def test_future_dynareport_insert_merges_with_later_ftp_folder(self) -> None:
        # 1. DynaReport discovers scheduled future SA2#181 without FTP folder
        p2 = self.db.update_meeting_metadata_bulk(
            [
                ("SA2", "181", "", "82767", "S2-181", "China", "2027-05-24", "2027-05-28", "")
            ]
        )
        self.assertEqual(p2["inserted"], 1)

        rows = self.db.search_meetings(wg_name="SA2", search_term="181")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["mtg_id"], "82767")
        self.assertEqual(rows[0]["start_date"], "2027-05-24")
        self.assertIsNone(rows[0]["url_key"])

        # 2. Later FTP folder is created and linked in-place without duplication
        p1 = self.db.insert_meetings_bulk(
            [
                {
                    "wg_name": "SA2",
                    "folder_name": "TSGS2_181_China_2027-05",
                    "meeting_num": "181",
                    "url_key": "tsg_sa/WG2_Arch/TSGS2_181_China_2027-05",
                    "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_181_China_2027-05/Docs/",
                    "fallback_start_date": "2027-05-01",
                }
            ]
        )
        self.assertEqual(p1["linked"], 1)
        self.assertEqual(p1["inserted"], 0)

        rows_after = self.db.search_meetings(wg_name="SA2", search_term="181")
        self.assertEqual(len(rows_after), 1)
        self.assertEqual(rows_after[0]["url_key"], "tsg_sa/WG2_Arch/TSGS2_181_China_2027-05")
        # Official DynaReport date is preserved over fallback
        self.assertEqual(rows_after[0]["start_date"], "2027-05-24")

    def test_fallback_date_overwritten_by_official_and_unknown_date_isolation(self) -> None:
        self.db.insert_meetings_bulk(
            [
                {
                    "wg_name": "SA2",
                    "folder_name": "TSGS2_178-AH-e_Electronic_2027-01",
                    "meeting_num": "178AHE",
                    "url_key": "tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01",
                    "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01/Docs/",
                    "fallback_start_date": "2027-01-01",
                },
                {
                    "wg_name": "RAN1",
                    "folder_name": "TSGR1_AH",
                    "meeting_num": "AH",
                    "url_key": "tsg_ran/WG1_RL1/TSGR1_AH",
                    "docs_url": "https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_AH/Docs/",
                    "fallback_start_date": "",
                },
                {
                    "wg_name": "SA2",
                    "folder_name": "TSGS2_90a-SIRIG_AH",
                    "meeting_num": "90A",
                    "url_key": "tsg_sa/WG2_Arch/TSGS2_90a-SIRIG_AH",
                    "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_90a-SIRIG_AH/Docs/",
                    "fallback_start_date": "",
                },
            ]
        )
        # Cleanup container row
        self.assertEqual(self.db.cleanup_container_rows(), 1)

        # Official DynaReport updates fallback 2027-01-01 -> 2027-01-25..29
        self.db.update_meeting_metadata_bulk(
            [
                ("SA2", "178AHE", "", "86724", "S2-178-AH-e", "Online", "2027-01-25", "2027-01-29", "")
            ]
        )
        # Re-running FTP insert does NOT overwrite official start_date back to fallback
        self.db.insert_meetings_bulk(
            [
                {
                    "wg_name": "SA2",
                    "folder_name": "TSGS2_178-AH-e_Electronic_2027-01",
                    "meeting_num": "178AHE",
                    "url_key": "tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01",
                    "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01/Docs/",
                    "fallback_start_date": "2027-01-01",
                }
            ]
        )
        dated = self.db.search_meetings(date_mode="dated")
        self.assertEqual([r["meeting_number"] for r in dated], ["178AHE"])
        self.assertEqual(dated[0]["start_date"], "2027-01-25")

        # Undated row stays hidden from dated/upcoming and appears only under unknown
        upcoming = self.db.search_meetings(date_from="2026-01-01", date_mode="dated")
        self.assertEqual(len(upcoming), 1)
        unknown = self.db.search_meetings(date_mode="unknown")
        self.assertEqual([r["meeting_number"] for r in unknown], ["90A"])
        self.assertEqual(self.db.count_unknown_date_meetings(), 1)

    def test_smart_docs_candidates_filters_old_indexed_meetings(self) -> None:
        today = datetime.date.today()
        recent_str = (today - datetime.timedelta(days=30)).isoformat()
        old_str = (today - datetime.timedelta(days=900)).isoformat()

        self.db.upsert_single_meeting(
            {
                "wg_name": "SA2",
                "meeting_number": "178",
                "url_key": "tsg_sa/WG2_Arch/TSGS2_178_Calgary",
                "start_date": recent_str,
                "end_date": recent_str,
                "first_tdoc": "S2-2611404",
                "last_tdoc": "S2-2612678",
            }
        )
        self.db.upsert_single_meeting(
            {
                "wg_name": "SA2",
                "meeting_number": "140",
                "url_key": "tsg_sa/WG2_Arch/TSGS2_140_Old",
                "start_date": old_str,
                "end_date": old_str,
                "first_tdoc": "S2-2001000",
                "last_tdoc": "S2-2002000",
            }
        )
        candidates = self.db.get_smart_docs_candidates(wg_names=["SA2"])
        self.assertEqual([c["meeting_num"] for c in candidates], ["178"])


class TestMeetingsCrawlerOrderAndUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_crawler_runs_ftp_then_dyna_then_smart_docs_for_target_wgs(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "crawl.db"
            thread = MeetingsCrawlerThread(
                db_path=db_path,
                target_wgs=["SA2"],
                docs_mode="smart",
            )
            call_order = []
            emitted_phases = []
            summaries = []

            thread.finished_path.connect(emitted_phases.append)
            thread.task_summary.connect(summaries.append)

            def fake_fetch(wg_name, url, is_ah=False):
                call_order.append(("ftp", wg_name))
                return [
                    {
                        "wg_name": wg_name,
                        "folder_name": "TSGS2_178-AH-e_Electronic_2027-01",
                        "meeting_num": "178AHE",
                        "url_key": "tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01",
                        "absolute_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01",
                        "is_ad_hoc": False,
                        "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01/Docs/",
                        "fallback_start_date": "2027-01-01",
                    }
                ]

            def fake_dyna(wg_name, dyna_url):
                call_order.append(("dyna", wg_name))
                return [
                    ("SA2", "178AHE", "", "86724", "S2-178-AH-e", "Online", "2027-01-25", "2027-01-29", "")
                ]

            def fake_docs(task):
                call_order.append(("docs", task["meeting_num"]))
                return (
                    (task["docs_url"], "S2-2700001", "S2", 2700001, "S2-2700100", "S2", 2700100, task["url_key"]),
                    100,
                )

            thread.fetch_wg_directories = fake_fetch
            thread.process_dynareport = fake_dyna
            thread.process_individual_meeting = fake_docs

            thread.run()

            self.assertEqual(
                [step for step, _ in call_order],
                ["ftp", "dyna", "docs"],
            )
            self.assertEqual(
                emitted_phases,
                ["MEETINGS_DB_PHASE_1", "MEETINGS_DB_PHASE_2", "MEETINGS_DB_PHASE_3"],
            )
            self.assertEqual(len(summaries), 1)
            self.assertEqual(summaries[0]["new_meetings"], 1)
            self.assertEqual(summaries[0]["docs_scanned"], 1)
            self.assertFalse(summaries[0]["cancelled"])

    def test_meetings_tab_unknown_dates_and_sync_controls(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "ui.db"
            cfg_path = Path(tmpdir) / "meetings_config.json"
            db = MeetingsDatabase(db_path)
            db.upsert_single_meeting(
                {
                    "wg_name": "SA2",
                    "meeting_number": "178AHE",
                    "url_key": "tsg_sa/WG2_Arch/TSGS2_178-AH-e_Electronic_2027-01",
                    "mtg_id": "86724",
                    "name": "S2-178-AH-e",
                    "location": "Online",
                    "start_date": "2027-01-25",
                    "end_date": "2027-01-29",
                    "is_electronic": 1,
                    "is_ad_hoc": 1,
                }
            )
            db.insert_meetings_bulk(
                [
                    {
                        "wg_name": "SA2",
                        "folder_name": "TSGS2_90a-SIRIG_AH",
                        "meeting_num": "90A",
                        "url_key": "tsg_sa/WG2_Arch/TSGS2_90a-SIRIG_AH",
                        "docs_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_90a-SIRIG_AH/Docs/",
                        "fallback_start_date": "",
                    }
                ]
            )

            host = QMainWindow()
            host.toast_manager = ToastManager(host)
            tab = MeetingsTab(db_path, config_file=cfg_path)
            host.setCentralWidget(tab)
            try:
                tab.wg_filter.setCheckedItems(["SA2"])
                tab.refresh_table()

                # Default "All Dates" excludes undated rows
                self.assertEqual(tab.table_model.rowCount(), 1)
                self.assertEqual(tab.table_model._data[0]["meeting_number"], "178AHE")
                self.assertIn("1 undated", tab.status_lbl.text())

                # Switching to "Unknown Dates" shows only the undated row
                idx_unk = tab.date_preset_combo.findText("Unknown Dates")
                self.assertGreaterEqual(idx_unk, 0)
                tab.date_preset_combo.setCurrentIndex(idx_unk)
                self.assertEqual(tab.table_model.rowCount(), 1)
                self.assertEqual(tab.table_model._data[0]["meeting_number"], "90A")

                # Inline sync banner lifecycle
                self.assertTrue(tab.sync_status_banner.isHidden())
                tab.on_sync_started("3GPP_Meetings")
                self.assertFalse(tab.sync_status_banner.isHidden())
                tab.on_sync_progress("Step 2/3: Updating dates & Portal IDs (1/1)...", 1, 1)
                self.assertEqual(tab.sync_progress_bar.maximum(), 1)
                tab.on_sync_finished({"new_meetings": 1, "updated_meetings": 1, "docs_scanned": 0, "elapsed_s": 1.2})
                self.assertTrue(tab.sync_status_banner.isHidden())

                # Sync menu includes scoped WG option when SA2 is selected
                tab.wg_filter.setCheckedItems(["SA2"])
                tab._refresh_sync_menu()
                action_texts = [a.text() for a in tab.sync_menu.actions() if a.text()]
                self.assertTrue(any("Sync SA2 Only" in t for t in action_texts), action_texts)

                # Custom sync dialog respects initial WG scope
                dlg = MeetingsSyncDialog(initial_wgs=["SA2"], parent=tab)
                opts = dlg.get_sync_options()
                self.assertEqual(opts["target_wgs"], ["SA2"])
                self.assertEqual(opts["docs_mode"], "smart")
                dlg.close()
            finally:
                host.toast_manager.dismiss_all()
                tab.shutdown()
                tab.close()
                host.close()


if __name__ == "__main__":
    unittest.main()
