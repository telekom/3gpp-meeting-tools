# --- File: tests/test_jump_to_tdoc_clipboard.py ---
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QApplication, QMenu

from modules.meetings.ui.tdocs_menus import build_quick_copy_menu, populate_quick_copy_menu
from modules.meetings.ui.ui_tabs import MeetingsTab


class TestJumpToTDocClipboard(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
        cls.db_path = Path("src/db/3gpp_data.db")

    def test_populate_quick_copy_menu_baseline(self):
        """Verifies actions in the Quick Copy menu for a baseline TDoc with full metadata."""
        menu = QMenu()
        row_data = {
            "TDoc": "S2-2605740",
            "Title": "Solution on Edge Architecture",
            "Source": "Deutsche Telekom",
        }
        docs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs"

        populate_quick_copy_menu(
            menu=menu,
            tdoc_id="S2-2605740",
            row_data=row_data,
            docs_ftp_url=docs_url,
        )

        action_texts = [act.text() for act in menu.actions()]
        self.assertIn("📋 Copy TDoc Number (S2-2605740)", action_texts)
        self.assertIn("📋 Copy Title", action_texts)
        self.assertIn("📋 Copy Full Citation", action_texts)
        self.assertIn("📋 Copy Document Download URL", action_texts)
        self.assertIn("🌐 Copy 3GU Portal URL", action_texts)
        self.assertIn("🔗 Copy Markdown Link", action_texts)

    def test_populate_quick_copy_menu_revision(self):
        """Verifies that for a revision TDoc, both revision and baseline download URLs appear."""
        menu = QMenu()
        row_data = {
            "TDoc": "S2-2605740r01",
            "Title": "Revised Solution on Edge Architecture",
            "Source": "Deutsche Telekom",
        }
        docs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs"

        populate_quick_copy_menu(
            menu=menu,
            tdoc_id="S2-2605740r01",
            row_data=row_data,
            docs_ftp_url=docs_url,
        )

        action_texts = [act.text() for act in menu.actions()]
        self.assertIn("📋 Copy TDoc Number (S2-2605740r01)", action_texts)
        self.assertIn("📋 Copy Revision Download URL", action_texts)
        self.assertIn("📋 Copy Baseline Document URL", action_texts)
        self.assertIn("🌐 Copy 3GU Portal URL", action_texts)
        self.assertIn("🔗 Copy Markdown Link", action_texts)

    def test_populate_quick_copy_menu_no_metadata(self):
        """Verifies that when only TDoc ID and URL are known, TDoc, URL and Portal copy options still populate."""
        menu = QMenu()
        docs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs"

        populate_quick_copy_menu(
            menu=menu,
            tdoc_id="S2-2605740",
            row_data={},
            docs_ftp_url=docs_url,
        )

        action_texts = [act.text() for act in menu.actions()]
        self.assertIn("📋 Copy TDoc Number (S2-2605740)", action_texts)
        self.assertNotIn("📋 Copy Title", action_texts)
        self.assertNotIn("📋 Copy Full Citation", action_texts)
        self.assertIn("📋 Copy Document Download URL", action_texts)
        self.assertIn("🌐 Copy 3GU Portal URL", action_texts)
        self.assertIn("🔗 Copy Markdown Link", action_texts)

    def test_quick_bar_btn_copy_visibility_and_wiring(self):
        """Tests that btn_copy_tdoc is hidden initially and becomes visible on valid TDoc lookup."""
        tab = MeetingsTab(self.db_path)
        self.assertTrue(hasattr(tab, "btn_copy_tdoc"))
        self.assertTrue(tab.btn_copy_tdoc.isHidden())

        # Mock find_meeting_by_tdoc to return a synthetic meeting
        fake_meeting = {
            "mtg_id": 9999,
            "wg_name": "SA2",
            "meeting_number": "168",
            "docs_folder_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs",
            "url_key": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju",
        }

        with patch.object(tab.db, "find_meeting_by_tdoc", return_value=fake_meeting):
            tab.global_tdoc_input.setText("S2-2605740")
            self.assertFalse(tab.btn_copy_tdoc.isHidden())
            self.assertIn("S2-2605740", tab.btn_copy_tdoc.toolTip())

            # Clear input -> button should hide
            tab.global_tdoc_input.setText("")
            self.assertTrue(tab.btn_copy_tdoc.isHidden())

        tab.close()

    def test_action_show_copy_menu_executes(self):
        """Tests that action_show_copy_menu constructs the menu without exceptions."""
        tab = MeetingsTab(self.db_path)
        fake_meeting = {
            "mtg_id": 9999,
            "wg_name": "SA2",
            "meeting_number": "168",
            "docs_folder_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs",
            "url_key": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju",
        }

        with patch.object(tab.db, "find_meeting_by_tdoc", return_value=fake_meeting):
            tab.global_tdoc_input.setText("S2-2605740")
            self.assertFalse(tab.btn_copy_tdoc.isHidden())

            # Mock menu.exec_ so it doesn't block the test
            with patch.object(QMenu, "exec_", return_value=None) as mock_exec:
                tab.search_controller.action_show_copy_menu()
                self.assertTrue(mock_exec.called)

        tab.close()

    def test_clipboard_mime_data_and_urls(self):
        """Verifies clipboard contents (plain text and rich HTML) when actions are triggered."""
        menu = QMenu()
        row_data = {
            "TDoc": "S2-2605740",
            "Title": "Solution on Edge Architecture",
            "Source": "Deutsche Telekom",
        }
        docs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs"

        populate_quick_copy_menu(
            menu=menu,
            tdoc_id="S2-2605740",
            row_data=row_data,
            docs_ftp_url=docs_url,
        )

        actions_by_text = {act.text(): act for act in menu.actions()}

        # 1. Test Document Download URL copy
        act_dl = actions_by_text["📋 Copy Document Download URL"]
        act_dl.trigger()
        expected_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168_Jeju/Docs/S2-2605740.zip"
        self.assertEqual(QApplication.clipboard().text(), expected_url)

        # 2. Test Markdown Link copy
        act_md = actions_by_text["🔗 Copy Markdown Link"]
        act_md.trigger()
        expected_md = f"[S2-2605740]({expected_url})"
        self.assertEqual(QApplication.clipboard().text(), expected_md)

        # 3. Test 3GU Portal URL copy
        act_portal = actions_by_text["🌐 Copy 3GU Portal URL"]
        act_portal.trigger()
        self.assertEqual(
            QApplication.clipboard().text(),
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId=S2-2605740"
        )

        # 4. Test TDoc Number copy (plain text + rich HTML)
        act_id = actions_by_text["📋 Copy TDoc Number (S2-2605740)"]
        act_id.trigger()
        self.assertEqual(QApplication.clipboard().text(), "S2-2605740")
        mime = QApplication.clipboard().mimeData()
        self.assertIn(expected_url, mime.html())


if __name__ == "__main__":
    unittest.main()
