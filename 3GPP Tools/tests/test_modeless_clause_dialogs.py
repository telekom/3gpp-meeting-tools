"""Tests for modeless top-level independent window behavior of clause excerpt and contribution dialogs."""

import os
import sys
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget

from modules.specifications.ui.spec_clause_dialog import SpecClauseExcerptDialog, SpecClauseDialog
from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

# Ensure single QApplication instance
app = QApplication.instance() or QApplication(sys.argv)


class TestModelessClauseDialogs(unittest.TestCase):
    """Verifies that clause excerpt and contribution dialogs are independent modeless windows."""

    def test_spec_clause_dialog_window_flags_and_modality(self):
        """Verify SpecClauseExcerptDialog is an independent top-level modeless window without staying on top."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            dlg = SpecClauseExcerptDialog(
                spec_num="23.501",
                spec_type="TS",
                version="17.4.0",
                fname="23501-h40.zip",
                url="https://www.3gpp.org/ftp/Specs/archive/23_series/23.501/23501-h40.zip",
                spec_dl_dir=Path(tmpdir),
                parent=parent_widget,
                auto_start=False,
            )
            # Verify Win32 parent is None so Win32 GW_OWNER does not force always-on-top over main app window
            self.assertIsNone(dlg.parent())
            self.assertIs(dlg._parent_ref, parent_widget)

            # Flags must contain Qt.Window and NOT Qt.WindowStaysOnTopHint
            self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
            self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))

            # Window type mask must be Qt.Window (0x1) and NOT Qt.Dialog (0x3)
            self.assertEqual(dlg.windowFlags() & Qt.WindowType_Mask, Qt.Window)

            # Window modality must be NonModal
            self.assertEqual(dlg.windowModality(), Qt.NonModal)
            dlg.close()

    def test_cr_pcr_dialog_window_flags_and_modality(self):
        """Verify CrPcrDialog is an independent top-level modeless window without staying on top."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            dlg = CrPcrDialog(
                spec_num="23.501",
                spec_type="TS",
                version="17.4.0",
                fname="23501-h40.zip",
                url="https://www.3gpp.org/ftp/Specs/archive/23_series/23.501/23501-h40.zip",
                spec_dl_dir=Path(tmpdir),
                parent=parent_widget,
                auto_start=False,
            )
            self.assertIsNone(dlg.parent())
            self.assertIs(dlg._parent_ref, parent_widget)
            self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
            self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
            self.assertEqual(dlg.windowFlags() & Qt.WindowType_Mask, Qt.Window)
            self.assertEqual(dlg.windowModality(), Qt.NonModal)
            dlg.close()


if __name__ == "__main__":
    unittest.main()
