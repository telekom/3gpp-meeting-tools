"""Tests for modeless top-level independent window behavior of utility dialogs and windows.

Verifies that modeless tool windows (clause excerpt, contribution generator,
database maintenance, process manager, inspector, email reader/emails/summary,
triage dialog, call flow generator, and session plan) do NOT become owned Win32 child
windows (GW_OWNER) and never stay forced on top of the main application window or external applications.
"""

import sys
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget

from core.ui.database_inspector import DatabaseInspectorDialog
from core.ui.database_maintenance_dialog import DatabaseMaintenanceDialog
from core.ui.process_manager_dialog import ProcessManagerDialog
from modules.emails.ui.email_summary_dialog import EmailSummaryDialog
from modules.emails.ui.tdoc_emails_dialog import StandaloneEmailReaderWindow, TDocEmailsDialog
from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow
from modules.meetings.ui.tdoc_triage_dialog import TDocTriageDialog
from modules.puml2visio.ui.callflow_dialog import CallFlowDialog
from modules.specifications.ui.spec_clause_dialog import SpecClauseExcerptDialog
from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

# Ensure single QApplication instance
app = QApplication.instance() or QApplication(sys.argv)


class TestModelessUtilityWindows(unittest.TestCase):
    """Verifies that all modeless utility dialogs and windows are independent top-level windows."""

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

    def test_database_maintenance_dialog_flags(self):
        """Verify DatabaseMaintenanceDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        dlg = DatabaseMaintenanceDialog(parent=parent_widget)
        self.assertIsNone(dlg.parent())
        self.assertIs(dlg._parent_ref, parent_widget)
        self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
        self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
        self.assertEqual(dlg.windowModality(), Qt.NonModal)
        dlg.close()

    def test_process_manager_dialog_flags(self):
        """Verify ProcessManagerDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        dlg = ProcessManagerDialog(parent=parent_widget)
        self.assertIsNone(dlg.parent())
        self.assertIs(dlg._parent_ref, parent_widget)
        self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
        self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
        self.assertEqual(dlg.windowModality(), Qt.NonModal)
        dlg.close()

    def test_database_inspector_dialog_flags(self):
        """Verify DatabaseInspectorDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            dummy_db = Path(tmpdir) / "test.db"
            dummy_db.touch()
            dlg = DatabaseInspectorDialog(db_path=dummy_db, display_name="Test DB", parent=parent_widget)
            self.assertIsNone(dlg.parent())
            self.assertIs(dlg._parent_ref, parent_widget)
            self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
            self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
            self.assertEqual(dlg.windowModality(), Qt.NonModal)
            dlg.close()

    def test_standalone_email_reader_window_flags(self):
        """Verify StandaloneEmailReaderWindow is an independent top-level window."""
        parent_widget = QWidget()
        win = StandaloneEmailReaderWindow(parent=parent_widget)
        self.assertIsNone(win.parent())
        self.assertIs(win._parent_ref, parent_widget)
        self.assertTrue(bool(win.windowFlags() & Qt.Window))
        self.assertFalse(bool(win.windowFlags() & Qt.WindowStaysOnTopHint))
        win.close()

    def test_tdoc_emails_dialog_flags(self):
        """Verify TDocEmailsDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            dummy_db = Path(tmpdir) / "emails.db"
            dummy_db.touch()
            dlg = TDocEmailsDialog(
                target_tdoc="S2-2401234",
                family_tdocs=["S2-2401234"],
                db_path=dummy_db,
                parent=parent_widget,
            )
            self.assertIsNone(dlg.parent())
            self.assertIs(dlg._parent_ref, parent_widget)
            self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
            self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
            self.assertEqual(dlg.windowModality(), Qt.NonModal)
            dlg.close()

    def test_tdoc_triage_dialog_flags(self):
        """Verify TDocTriageDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            dlg = TDocTriageDialog(
                tdoc_data={"TDoc": "S2-2401234"},
                meeting_dir=Path(tmpdir),
                parent=parent_widget,
                auto_start=False,
            )
            self.assertIsNone(dlg.parent())
            self.assertIs(dlg._parent_ref, parent_widget)
            self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
            self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
            self.assertFalse(dlg.isModal())
            dlg.close()

    def test_callflow_dialog_flags(self):
        """Verify CallFlowDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        dlg = CallFlowDialog(parent=parent_widget)
        self.assertIsNone(dlg.parent())
        self.assertIs(dlg._parent_ref, parent_widget)
        self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
        self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
        self.assertFalse(dlg.isModal())
        dlg.close()

    def test_email_summary_dialog_flags(self):
        """Verify EmailSummaryDialog is an independent modeless top-level window."""
        parent_widget = QWidget()
        dlg = EmailSummaryDialog(
            tdoc_id="S2-2401234",
            family_tdocs=[],
            emails_data=[],
            parent=parent_widget,
            auto_start=False,
        )
        self.assertIsNone(dlg.parent())
        self.assertIs(dlg._parent_ref, parent_widget)
        self.assertTrue(bool(dlg.windowFlags() & Qt.Window))
        self.assertFalse(bool(dlg.windowFlags() & Qt.WindowStaysOnTopHint))
        self.assertFalse(dlg.isModal())
        dlg.close()

    def test_session_plan_window_flags(self):
        """Verify SessionPlanWindow is an independent top-level window."""
        parent_widget = QWidget()
        with tempfile.TemporaryDirectory() as tmpdir:
            mtg_dir = Path(tmpdir)
            agenda_dir = mtg_dir / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            win = SessionPlanWindow(
                meeting_dir=mtg_dir,
                mtg_info={"meeting_number": "177", "location": "Prague"},
                parent=parent_widget,
            )
            self.assertIsNone(win.parent())
            self.assertIs(win._parent_ref, parent_widget)
            self.assertTrue(bool(win.windowFlags() & Qt.Window))
            self.assertFalse(bool(win.windowFlags() & Qt.WindowStaysOnTopHint))
            win.close()


if __name__ == "__main__":
    unittest.main()
