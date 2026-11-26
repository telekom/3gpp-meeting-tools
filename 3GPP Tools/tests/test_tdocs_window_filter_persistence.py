"""
Regression tests for per-meeting TDocsWindow filter persistence and MeetingsTab
close-flush behavior.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication

from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase
from modules.meetings.ui.tdocs_window import TDocsWindow
from modules.meetings.ui.ui_tabs import MeetingsTab


class TestTDocsWindowFilterPersistence(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def _sample_rows(self) -> list[dict]:
        return [
            {
                "TDoc": "S2-2611863",
                "Title": "NAS Routing Alt 2 baseline",
                "Source": "Nokia",
                "Type": "pCR",
                "For": "Approval",
                "Agenda Item": "20.6.1",
                "Sub-Item": "NAS Routing",
                "Related WIs": "FS_6G_ARC",
                "TDoc Status": "",
                "Secretary Remarks": "Baseline comment",
                "My Status": "",
            },
            {
                "TDoc": "S2-2611895",
                "Title": "IMS handling clarification",
                "Source": "Ericsson",
                "Type": "CR",
                "For": "Discussion",
                "Agenda Item": "20.1.2",
                "Sub-Item": "B2BUA",
                "Related WIs": "IMS_ENH",
                "TDoc Status": "revised",
                "Secretary Remarks": "",
                "My Status": "🟢 Support",
            },
            {
                "TDoc": "S2-2611900",
                "Title": "Policy draft update",
                "Source": "Huawei",
                "Type": "discussion",
                "For": "Information",
                "Agenda Item": "20.6.6",
                "Sub-Item": "Policy",
                "Related WIs": "FS_6G_ARC",
                "TDoc Status": "available",
                "Secretary Remarks": "Pending review",
                "My Status": "⚪ Neutral",
            },
        ]

    def test_tdocs_window_persists_and_restores_all_filters_per_meeting(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            meeting_dir = tmp_path / "SA2_178"
            agenda_dir = meeting_dir / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            dummy_xls = agenda_dir / "TDocs.xlsx"
            dummy_xls.write_text("placeholder", encoding="utf-8")

            # Seed user stances and tag so my_status_combo and tags_combo have distinct options
            tdb = TDocsDatabase(agenda_dir / "user_tdocs.db")
            tdb.upsert("S2-2611863", "⚪ Neutral", "note", tags="Priority")
            tdb.upsert("S2-2611895", "🟢 Support", "note2", tags="")

            ini_path = tmp_path / "tdocs_filters.ini"
            qsettings = QSettings(str(ini_path), QSettings.IniFormat)
            mtg_sa2 = {"wg_name": "SA2", "meeting_number": "178", "location": "Calgary"}

            win = TDocsWindow(mtg_sa2, self._sample_rows(), str(dummy_xls), filter_settings=qsettings)

            # Set search text WITHOUT waiting 300ms debounce timer to test flush on close
            win.search_input.setText("NAS")
            win.company_combo.setCheckedItems(["Nokia", "Huawei"])
            win._on_company_changed(win.company_combo.getCheckedItems())
            win.type_combo.setCheckedItems(["pCR"])
            win._on_type_changed(win.type_combo.getCheckedItems())
            win.sub_item_combo.setCheckedItems(["NAS Routing"])
            win._on_sub_item_changed(win.sub_item_combo.getCheckedItems())
            # Ensure empty string ("(Empty)") selection survives round-trip on TDoc Status
            win.status_combo.setCheckedItems([""])
            win._on_status_changed(win.status_combo.getCheckedItems())
            win.my_status_combo.setCheckedItems(["⚪ Neutral"])
            win._on_my_status_changed(win.my_status_combo.getCheckedItems())
            win.tags_combo.setCheckedItems(["Priority"])
            win._on_tags_changed(win.tags_combo.getCheckedItems())
            win.chk_no_comments.setChecked(True)
            win.btn_unread_emails.click()
            win.order_combo.setCurrentIndex(win.order_combo.findData("TDOC"))
            win.row_mode_combo.setCurrentIndex(win.row_mode_combo.findData("TDOCS"))
            self.app.processEvents()
            win.close()

            # Reopen same meeting with fresh QSettings reading same INI file
            qsettings_reloaded = QSettings(str(ini_path), QSettings.IniFormat)
            win2 = TDocsWindow(mtg_sa2, self._sample_rows(), str(dummy_xls), filter_settings=qsettings_reloaded)
            self.app.processEvents()

            self.assertEqual(win2.search_input.text(), "NAS")
            self.assertEqual(win2.proxy.include_terms, ["nas"])
            self.assertEqual(win2.proxy.rowCount(), 0)
            # Clear unread-email & no-comments filters so S2-2611863 (matching all other filters) appears
            win2.btn_unread_emails.click()
            win2.chk_no_comments.setChecked(False)
            self.assertEqual(win2.proxy.rowCount(), 1)
            win2.chk_no_comments.setChecked(True)
            win2.btn_unread_emails.click()
            self.assertEqual(win2.company_combo.getCheckedItems(), ["Huawei", "Nokia"])
            self.assertEqual(win2.type_combo.getCheckedItems(), ["pCR"])
            self.assertEqual(win2.sub_item_combo.getCheckedItems(), ["NAS Routing"])
            self.assertEqual(win2.status_combo.getCheckedItems(), [""])
            self.assertEqual(win2.my_status_combo.getCheckedItems(), ["⚪ Neutral"])
            self.assertEqual(win2.tags_combo.getCheckedItems(), ["Priority"])
            self.assertTrue(win2.chk_no_comments.isChecked())
            self.assertTrue(win2.btn_unread_emails.isChecked())
            self.assertEqual(win2.order_combo.currentData(), "TDOC")
            self.assertEqual(win2.row_mode_combo.currentData(), "TDOCS")

            # Delayed _refresh_comboboxes (e.g. background enrichment) must keep saved subsets
            win2._refresh_comboboxes()
            self.assertEqual(win2.company_combo.getCheckedItems(), ["Huawei", "Nokia"])
            self.assertEqual(win2.status_combo.getCheckedItems(), [""])
            self.assertEqual(win2.my_status_combo.getCheckedItems(), ["⚪ Neutral"])

            # Different meeting remains completely isolated
            mtg_ran2 = {"wg_name": "RAN2", "meeting_number": "134", "location": "Prague"}
            win_other = TDocsWindow(
                mtg_ran2,
                self._sample_rows(),
                str(dummy_xls),
                filter_settings=QSettings(str(ini_path), QSettings.IniFormat),
            )
            self.assertEqual(win_other.search_input.text(), "")
            self.assertTrue(win_other.company_combo.isAllChecked())
            self.assertFalse(win_other.btn_unread_emails.isChecked())
            win_other.close()

            # Clearing filters updates persisted state so next open is unfiltered
            win2._clear_all_filters()
            win2.close()

            win3 = TDocsWindow(
                mtg_sa2,
                self._sample_rows(),
                str(dummy_xls),
                filter_settings=QSettings(str(ini_path), QSettings.IniFormat),
            )
            self.assertEqual(win3.search_input.text(), "")
            self.assertTrue(win3.company_combo.isAllChecked())
            self.assertTrue(win3.status_combo.isAllChecked())
            self.assertTrue(win3.my_status_combo.isAllChecked())
            self.assertFalse(win3.chk_no_comments.isChecked())
            self.assertFalse(win3.btn_unread_emails.isChecked())
            win3.close()

    def test_meetings_tab_shutdown_flushes_pending_filter_save(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            cfg = Path(tmp_dir) / "meetings_config.json"
            tab = MeetingsTab(Path("src/db/3gpp_data.db"), config_file=cfg)
            tab.search_input.setText("Prague")
            # Debounce timer is active (1000ms); shutdown immediately before it fires
            self.assertTrue(tab.save_filters_timer.isActive())
            tab.shutdown()
            self.assertFalse(tab.save_filters_timer.isActive())

            tab2 = MeetingsTab(Path("src/db/3gpp_data.db"), config_file=cfg)
            self.assertEqual(tab2.search_input.text(), "Prague")
            tab2.shutdown()


if __name__ == "__main__":
    unittest.main()
