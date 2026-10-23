# --- File: tests/test_meetings_ui_redesign.py ---
import sys
import unittest
from pathlib import Path
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import QApplication

from modules.meetings.ui.models import MeetingsTableModel, format_meeting_dates
from modules.meetings.ui.meetings_delegates import WGTagDelegate
from modules.meetings.ui.ui_tabs import MeetingsTab
from modules.meetings.ui.dialogs import MeetingsConfigDialog


class TestMeetingsUiRedesign(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
        cls.db_path = Path("src/db/3gpp_data.db")

    def test_table_headers_and_column_count(self):
        model = MeetingsTableModel()
        self.assertEqual(model.columnCount(), 7)
        expected_headers = ["", "Action", "WG", "Meeting", "Location", "Dates", "Allocated TDocs"]
        for col, expected in enumerate(expected_headers):
            self.assertEqual(model.headerData(col, Qt.Horizontal, Qt.DisplayRole), expected)

    def test_date_formatting_logic(self):
        # Same month & year
        self.assertEqual(format_meeting_dates("2026-10-12", "2026-10-16"), "12 – 16 Oct 2026")
        # Different month, same year
        self.assertEqual(format_meeting_dates("2026-10-28", "2026-11-02"), "28 Oct – 2 Nov 2026")
        # Different year
        self.assertEqual(format_meeting_dates("2026-12-28", "2027-01-03"), "28 Dec 2026 – 3 Jan 2027")
        # Single date
        self.assertEqual(format_meeting_dates("2026-10-12", ""), "12 Oct 2026")
        # Empty
        self.assertEqual(format_meeting_dates("", ""), "—")

    def test_model_data_rendering_and_full_prefixes(self):
        test_rows = [
            {
                "wg_name": "CT4",
                "meeting_number": "139",
                "location": "South Korea",
                "start_date": "2027-02-22",
                "end_date": "2027-02-26",
                "is_electronic": 0,
                "first_tdoc": "C4-264009",
                "last_tdoc": "C4-264023",
            },
            {
                "wg_name": "SA2",
                "meeting_number": "175AHE",
                "location": "Online",
                "start_date": "2026-06-24",
                "end_date": "2026-07-01",
                "is_electronic": 1,
                "first_tdoc": "S2-2605675",
                "last_tdoc": "S2-2605759",
            },
            {
                "wg_name": "RAN2",
                "meeting_number": "138",
                "location": "Calgary",
                "start_date": "2026-11-16",
                "end_date": "2026-11-20",
                "is_electronic": 0,
                "first_tdoc": "",
                "last_tdoc": "",
            }
        ]
        model = MeetingsTableModel(test_rows)

        # Row 0: CT4 (In-Person)
        idx_wg = model.index(0, 2)
        idx_mtg = model.index(0, 3)
        idx_loc = model.index(0, 4)
        idx_dates = model.index(0, 5)
        idx_tdocs = model.index(0, 6)

        self.assertEqual(model.data(idx_wg, Qt.DisplayRole), "CT4")
        self.assertEqual(model.data(idx_mtg, Qt.DisplayRole), "139")
        self.assertEqual(model.data(idx_loc, Qt.DisplayRole), "📍 South Korea")
        self.assertEqual(model.data(idx_dates, Qt.DisplayRole), "22 – 26 Feb 2027")
        # Full prefix must be preserved
        self.assertEqual(model.data(idx_tdocs, Qt.DisplayRole), "C4-264009 → C4-264023")

        # Row 1: SA2 (Electronic)
        idx_loc1 = model.index(1, 4)
        idx_tdocs1 = model.index(1, 6)
        self.assertEqual(model.data(idx_loc1, Qt.DisplayRole), "🌐 Online")
        self.assertEqual(model.data(idx_tdocs1, Qt.DisplayRole), "S2-2605675 → S2-2605759")

        # Row 2: Upcoming with no allocated TDocs yet
        idx_tdocs2 = model.index(2, 6)
        self.assertEqual(model.data(idx_tdocs2, Qt.DisplayRole), "—")

    def test_meetings_tab_initialization_and_layout(self):
        tab = MeetingsTab(self.db_path)

        # 1. Full-width table with delegates
        self.assertEqual(tab.table.model().columnCount(), 7)
        self.assertIsInstance(tab.table.itemDelegateForColumn(2), WGTagDelegate)

        # 2. Floating comparison cart starts hidden when empty
        self.assertTrue(tab.cart_frame.isHidden())

        # When slot updated, cart becomes visible
        tab._update_cart_ui({"name": "S2-2605740"}, None)
        self.assertFalse(tab.cart_frame.isHidden())
        self.assertIn("S2-2605740", tab.lbl_slot_a.text())
        self.assertFalse(tab.btn_compare.isEnabled())

        # When both slots populated, compare button is enabled
        tab._update_cart_ui({"name": "S2-2605740"}, {"name": "S2-2605741"})
        self.assertFalse(tab.cart_frame.isHidden())
        self.assertTrue(tab.btn_compare.isEnabled())

        # When cleared, cart hides
        tab._update_cart_ui(None, None)
        self.assertTrue(tab.cart_frame.isHidden())

        # 3. TSG pills exist and function
        self.assertIn("SA", tab.tsg_buttons)
        self.assertIn("RAN", tab.tsg_buttons)
        self.assertIn("CT", tab.tsg_buttons)
        self.assertIn("", tab.tsg_buttons)  # "All"

        # 4. Danger Zone: Clear button is in MeetingsConfigDialog
        self.assertIsNotNone(tab.config_dialog.btn_clear)
        self.assertIn("Clear All Meetings", tab.config_dialog.btn_clear.text())

        # 5. Status summary label exists
        self.assertTrue(hasattr(tab, "status_lbl"))
        self.assertIn("Showing", tab.status_lbl.text())

        tab.close()


if __name__ == "__main__":
    unittest.main()
