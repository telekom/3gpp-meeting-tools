import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QApplication, QMenu


class TestMeetingPortalMenuActions(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_add_meeting_portal_actions_skips_when_no_mtg_id(self):
        from modules.meetings.ui.meetings_menu import add_meeting_portal_actions

        menu = QMenu()
        added = add_meeting_portal_actions(menu, {"wg_name": "SA2"})
        self.assertFalse(added)
        self.assertEqual(menu.actions(), [])

    @patch("modules.meetings.ui.meetings_menu.open_portal_create_tdoc")
    @patch("modules.meetings.ui.meetings_menu.open_portal_meeting")
    def test_add_meeting_portal_actions_active_and_past_meetings(self, mock_open_mtg, mock_open_create):
        from modules.meetings.ui.meetings_menu import add_meeting_portal_actions

        future_date = QDate.currentDate().addDays(5).toString("yyyy-MM-dd")
        menu_active = QMenu()
        added_active = add_meeting_portal_actions(
            menu_active, {"mtg_id": 3210, "end_date": future_date}
        )
        self.assertTrue(added_active)
        texts_active = [a.text() for a in menu_active.actions()]
        self.assertEqual(texts_active, ["🖥️ 3GU Meeting Portal", "📝 Reserve / Contribute TDoc"])

        menu_active.actions()[0].trigger()
        mock_open_mtg.assert_called_once_with(3210)
        menu_active.actions()[1].trigger()
        mock_open_create.assert_called_once_with(3210)

        past_date = QDate.currentDate().addDays(-5).toString("yyyy-MM-dd")
        menu_past = QMenu()
        added_past = add_meeting_portal_actions(
            menu_past, {"mtg_id": 3210, "end_date": past_date}
        )
        self.assertTrue(added_past)
        texts_past = [a.text() for a in menu_past.actions()]
        self.assertEqual(texts_past, ["🖥️ 3GU Meeting Portal"])

    def test_populate_meeting_context_menu_includes_portal_actions(self):
        from modules.meetings.ui.meetings_menu import populate_meeting_context_menu

        tab = MagicMock()
        tab.table.selectionModel().selectedRows.return_value = []
        tab.db.is_active_sync_meeting.return_value = False
        tab.settings.cache_dir = tempfile.gettempdir()
        del tab.dl_dir_input

        menu = QMenu()
        future_date = QDate.currentDate().addDays(2).toString("yyyy-MM-dd")
        row_data = {
            "mtg_id": 4567,
            "wg_name": "SA2",
            "meeting_number": "178",
            "end_date": future_date,
        }
        populate_meeting_context_menu(tab, menu, row_data, 0)
        texts = [a.text() for a in menu.actions()]
        self.assertIn("🖥️ 3GU Meeting Portal", texts)
        self.assertIn("📝 Reserve / Contribute TDoc", texts)

    def test_tdocs_window_resources_menu_includes_portal_actions(self):
        from modules.meetings.ui.tdocs_window import TDocsWindow

        with tempfile.TemporaryDirectory() as tmp_dir:
            meeting_dir = Path(tmp_dir)
            agenda_dir = meeting_dir / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)

            window = TDocsWindow.__new__(TDocsWindow)
            super(TDocsWindow, window).__init__()
            window.meeting_dir = meeting_dir
            window.agenda_dir = agenda_dir
            window.is_sa2 = True
            window.main_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_Calgary"
            window.docs_ftp_url = ""
            window.revisions_url = ""
            window.folder_menu = QMenu()
            window._open_meeting_folder = MagicMock()
            window._open_agenda_file = MagicMock()
            window._show_unmatched_companies = MagicMock()
            window._export_reports = MagicMock()

            future_date = QDate.currentDate().addDays(2).toString("yyyy-MM-dd")
            window.mtg_info = {
                "wg_name": "SA2",
                "is_electronic": 0,
                "mtg_id": 4567,
                "end_date": future_date,
            }
            TDocsWindow._populate_resources_menu(window)
            texts = [a.text() for a in window.folder_menu.actions()]
            self.assertIn("🌐 Web FTP: Main Folder", texts)
            self.assertIn("🖥️ 3GU Meeting Portal", texts)
            self.assertIn("📝 Reserve / Contribute TDoc", texts)
            self.assertGreater(
                texts.index("🖥️ 3GU Meeting Portal"),
                texts.index("🌐 Web FTP: Main Folder"),
            )

            # Past meeting keeps 3GU Meeting Portal, omits Reserve / Contribute TDoc
            window.mtg_info["end_date"] = QDate.currentDate().addDays(-10).toString("yyyy-MM-dd")
            TDocsWindow._populate_resources_menu(window)
            past_texts = [a.text() for a in window.folder_menu.actions()]
            self.assertIn("🖥️ 3GU Meeting Portal", past_texts)
            self.assertNotIn("📝 Reserve / Contribute TDoc", past_texts)


if __name__ == "__main__":
    unittest.main()
