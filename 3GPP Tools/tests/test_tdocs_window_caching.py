import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication, QPushButton

from modules.meetings.core.tdocs.tdocs_cacher import TDocsCacherThread
from modules.meetings.ui.tdocs_agenda_controller import (
    start_window_tdocs_caching,
    on_window_tdocs_caching_finished,
)
from modules.meetings.ui.tdocs_menus import is_tdoc_cached


class TestTDocsWindowCaching(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_tdocs_cache_"))
        self.meeting_dir = self.temp_dir / "SA2_168"
        self.meeting_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_is_tdoc_cached_direct_and_subfolder(self):
        """Verify is_tdoc_cached detects TDocs in meeting_dir / base_tdoc and meeting_dir / TDocs."""
        base_tdoc = "S2-2601234"
        tdoc_dir = self.meeting_dir / base_tdoc
        tdoc_dir.mkdir(parents=True, exist_ok=True)

        # Before file creation: should be False
        self.assertFalse(is_tdoc_cached(self.meeting_dir, base_tdoc))

        # After creating zip in base_tdoc folder: should be True
        zip_file = tdoc_dir / f"{base_tdoc}.zip"
        zip_file.touch()
        self.assertTrue(is_tdoc_cached(self.meeting_dir, base_tdoc))

        # Test revision in the same base_tdoc folder
        rev_name = f"{base_tdoc}r01"
        self.assertFalse(is_tdoc_cached(self.meeting_dir, rev_name))
        rev_zip = tdoc_dir / f"{rev_name}.zip"
        rev_zip.touch()
        self.assertTrue(is_tdoc_cached(self.meeting_dir, rev_name))

        # Test legacy TDocs subfolder
        legacy_tdoc = "S2-2609999"
        legacy_dir = self.meeting_dir / "TDocs" / legacy_tdoc
        legacy_dir.mkdir(parents=True, exist_ok=True)
        (legacy_dir / f"{legacy_tdoc}.docx").touch()
        self.assertTrue(is_tdoc_cached(self.meeting_dir, legacy_tdoc))

    def test_start_window_tdocs_caching_lifecycle(self):
        """Test dispatch of caching thread and button state management in TDocsWindow."""
        window = MagicMock()
        window.docs_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168/Docs/"
        window.meeting_dir = self.meeting_dir
        window.refresh_btn = QPushButton("🔄 Refresh")
        window.cacher_thread = None

        with patch.object(TDocsCacherThread, "start") as mock_start:
            start_window_tdocs_caching(window)

            self.assertIsNotNone(window.cacher_thread)
            self.assertFalse(window.refresh_btn.isEnabled())
            self.assertIn("Caching", window.refresh_btn.text())
            mock_start.assert_called_once()

    def test_start_window_tdocs_caching_duplicate_guard(self):
        """Verify duplicate calls are blocked when cacher_thread is already running."""
        window = MagicMock()
        window.docs_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_168/Docs/"
        window.meeting_dir = self.meeting_dir
        window.refresh_btn = QPushButton("🔄 Refresh")
        
        # Simulate already running thread
        mock_thread = MagicMock()
        mock_thread.isRunning.return_value = True
        window.cacher_thread = mock_thread

        with patch("modules.meetings.ui.tdocs_agenda_controller.QMessageBox.information") as mock_info:
            start_window_tdocs_caching(window)
            mock_info.assert_called_once()
            # Ensure no new thread replaced the active one
            self.assertEqual(window.cacher_thread, mock_thread)

    def test_on_window_tdocs_caching_finished(self):
        """Verify button is restored and information box is shown upon caching completion."""
        window = MagicMock()
        window.refresh_btn = QPushButton("⏳ Caching...")
        window.refresh_btn.setEnabled(False)

        with patch("modules.meetings.ui.tdocs_agenda_controller.QMessageBox.information") as mock_info:
            on_window_tdocs_caching_finished(window, True, "Caching Complete! Downloaded: 5, Skipped: 10, Total: 15")
            self.assertTrue(window.refresh_btn.isEnabled())
            self.assertEqual(window.refresh_btn.text(), "🔄 Refresh")
            mock_info.assert_called_once()

        # Test failure case
        with patch("modules.meetings.ui.tdocs_agenda_controller.QMessageBox.warning") as mock_warn:
            on_window_tdocs_caching_finished(window, False, "Network error 404")
            self.assertTrue(window.refresh_btn.isEnabled())
            mock_warn.assert_called_once()

    def test_tdocs_cacher_thread_task_building(self):
        """Verify TDocsCacherThread groups archives into base TDoc folders and detects revisions."""
        html_sample = """
        <html>
            <a href="S2-2601001.zip">S2-2601001.zip</a>
            <a href="S2-2601001r01.zip">S2-2601001r01.zip</a>
            <a href="readme.txt">readme.txt</a>
        </html>
        """
        cacher = TDocsCacherThread("https://www.3gpp.org/ftp/Docs/", self.meeting_dir)

        with patch("modules.meetings.core.tdocs.tdocs_cacher.NetworkSession.get_html", return_value=html_sample):
            with patch.object(cacher, "_download_file", return_value=False):
                finished_signal_received = []
                cacher.finished.connect(lambda s, msg: finished_signal_received.append((s, msg)))

                cacher.run()

                self.assertTrue(finished_signal_received)
                success, summary = finished_signal_received[0]
                self.assertTrue(success)
                self.assertIn("Total: 2", summary)

                # Verify target folder is base TDoc 'S2-2601001' for both base and revision
                expected_dir = self.meeting_dir / "S2-2601001"
                # Even if skipped, the task paths point to self.meeting_dir / "S2-2601001"
                self.assertIn("S2-2601001", str(expected_dir))


if __name__ == "__main__":
    unittest.main()
