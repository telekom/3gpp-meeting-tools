import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication, QMenu, QAction


class TestTDocWindowExportShortcut(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.meeting_dir = Path(self.temp_dir.name)
        self.agenda_dir = self.meeting_dir / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)

        from modules.meetings.ui.tdocs_window import TDocsWindow

        # Create a lightweight dummy TDocsWindow instance without full heavy UI setup
        self.window = TDocsWindow.__new__(TDocsWindow)
        super(TDocsWindow, self.window).__init__()
        self.window.meeting_dir = self.meeting_dir
        self.window.agenda_dir = self.agenda_dir
        self.window.mtg_info = {"wg_name": "SA2", "is_electronic": 0}
        self.window.is_sa2 = True
        self.window.main_ftp_url = ""
        self.window.docs_ftp_url = ""
        self.window.revisions_url = ""

        # Set up menus and actions
        self.window.folder_menu = QMenu()
        self.window.export_menu = QMenu()
        self.window.export_menu_sep = self.window.export_menu.addSeparator()
        self.window.act_open_export_dir = self.window.export_menu.addAction("📂 Open Export Folder")
        self.window._open_meeting_folder = MagicMock()
        self.window._open_agenda_file = MagicMock()
        self.window._show_unmatched_companies = MagicMock()
        self.window._export_reports = MagicMock()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_export_menu_visibility_when_export_dir_missing(self):
        from modules.meetings.ui.tdocs_window import TDocsWindow

        TDocsWindow._on_export_menu_about_to_show(self.window)
        self.assertFalse(self.window.act_open_export_dir.isVisible())
        self.assertFalse(self.window.export_menu_sep.isVisible())

    def test_export_menu_visibility_when_export_dir_present(self):
        from modules.meetings.ui.tdocs_window import TDocsWindow

        export_dir = self.meeting_dir / "Export"
        export_dir.mkdir(parents=True, exist_ok=True)

        TDocsWindow._on_export_menu_about_to_show(self.window)
        self.assertTrue(self.window.act_open_export_dir.isVisible())
        self.assertTrue(self.window.export_menu_sep.isVisible())

    def test_resources_menu_includes_export_folder_when_present(self):
        from modules.meetings.ui.tdocs_window import TDocsWindow

        # 1. When Export does not exist
        TDocsWindow._populate_resources_menu(self.window)
        action_texts = [a.text() for a in self.window.folder_menu.actions()]
        self.assertNotIn("📂 Local: Export Folder", action_texts)

        # 2. When Export exists
        export_dir = self.meeting_dir / "Export"
        export_dir.mkdir(parents=True, exist_ok=True)

        TDocsWindow._populate_resources_menu(self.window)
        action_texts = [a.text() for a in self.window.folder_menu.actions()]
        self.assertIn("📂 Local: Export Folder", action_texts)

    @patch("modules.meetings.ui.tdocs_window._open_folder")
    def test_open_export_folder_calls_open_folder(self, mock_open_folder):
        from modules.meetings.ui.tdocs_window import TDocsWindow

        export_dir = self.meeting_dir / "Export"
        export_dir.mkdir(parents=True, exist_ok=True)

        TDocsWindow._open_export_folder(self.window)
        mock_open_folder.assert_called_once_with(export_dir)
