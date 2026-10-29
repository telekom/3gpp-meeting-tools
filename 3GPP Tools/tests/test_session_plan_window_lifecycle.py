# --- File: tests/test_session_plan_window_lifecycle.py ---
import unittest
from pathlib import Path
import tempfile
import shutil
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication, QWidget, QDialog
from PyQt5.QtGui import QCloseEvent

# Ensure QApplication exists in test environment
app = QApplication.instance()
if not app:
    app = QApplication([])

from modules.meetings.ui.session_plan_window import SessionPlanWindow
from modules.meetings.ui.ui_tabs import MeetingsTab


class TestSessionPlanWindowLifecycle(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_sp_lifecycle_"))
        self.agenda_dir = self.temp_dir / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)
        self.mtg_info = {
            "meeting_number": "177",
            "location": "Test Venue",
            "start_date": "2026-10-12",
        }

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_session_plan_window_clock_timer_stops_on_close_and_resumes_on_show(self):
        """Verify clock_timer starts on init, stops on close, and resumes on show."""
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            win.show()
            self.assertTrue(win.clock_timer.isActive(), "Clock timer should be active after window is shown")

            # Close the window
            win.close()
            self.assertFalse(win.clock_timer.isActive(), "Clock timer should stop when window is closed")

            # Show again
            win.show()
            self.assertTrue(win.clock_timer.isActive(), "Clock timer should resume when window is shown again")
        finally:
            win.close()

    def test_session_plan_window_closes_child_dialogs_on_close(self):
        """Verify modal or modeless child dialogs parented to SessionPlanWindow are closed on closeEvent."""
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            win.show()
            child_dlg = QDialog(win)
            child_dlg.show()
            self.assertTrue(child_dlg.isVisible())

            win.close()
            self.assertFalse(child_dlg.isVisible(), "Child dialog should be closed when SessionPlanWindow is closed")
        finally:
            win.close()

    def test_tdocs_window_close_closes_session_plan_window(self):
        """Verify that calling TDocsWindow.closeEvent() explicitly closes _session_plan_window."""
        from modules.meetings.ui.tdocs_window import TDocsWindow

        class DummyTDocsWindow(TDocsWindow):
            def __init__(self):
                QWidget.__init__(self)
                self._session_plan_window = MagicMock()
                self.email_window = None

        tdocs_win = DummyTDocsWindow()
        mock_sp_win = tdocs_win._session_plan_window

        # Trigger closeEvent
        event = QCloseEvent()
        tdocs_win.closeEvent(event)

        mock_sp_win.close.assert_called_once()
        self.assertIsNone(tdocs_win._session_plan_window)

    def test_meetings_tab_shutdown_closes_session_plan_and_tdoc_windows(self):
        """Verify MeetingsTab.shutdown closes all TDoc windows and their SessionPlanWindows."""
        class DummyMeetingsTab(QWidget):
            def __init__(self):
                super().__init__()
                self.tdoc_windows = {}

            shutdown = MeetingsTab.shutdown

        tab = DummyMeetingsTab()
        mock_tdoc_win = MagicMock()
        mock_sp_win = MagicMock()
        mock_tdoc_win._session_plan_window = mock_sp_win
        tab.tdoc_windows["SP-177"] = mock_tdoc_win

        tab.shutdown()

        mock_sp_win.close.assert_called_once()
        mock_tdoc_win.close.assert_called_once()
        self.assertEqual(len(tab.tdoc_windows), 0)

    def test_main_window_close_event_closes_all_top_level_windows(self):
        """Verify MainWindow.closeEvent iterates topLevelWidgets and closes any remaining windows."""
        from main_window import DragDropUI

        class DummyMainWindow(DragDropUI):
            def __init__(self):
                QWidget.__init__(self)
                self.status_bar = MagicMock()
                self.save_cache = MagicMock()

        main_win = DummyMainWindow()
        floating_win = QWidget()
        floating_win.show()
        self.assertTrue(floating_win.isVisible())

        event = QCloseEvent()
        main_win.closeEvent(event)

        self.assertFalse(floating_win.isVisible(), "Standalone top-level window should be closed by MainWindow.closeEvent")

    def test_rich_slot_tooltip_formatting(self):
        """Verify _build_rich_slot_tooltip produces modern, well-sectioned HTML."""
        from modules.meetings.core.session_plan_models import (
            SessionSlot,
            AgendaSlotItem,
            MeetingStreamInfo,
            SessionPlanSnapshot,
            UserSessionAnnotation,
            DelegationMember,
            TopicOwnership,
        )

        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            snap = SessionPlanSnapshot(meeting_number="177")
            snap.streams[2] = MeetingStreamInfo(
                stream_number=2,
                stream_name="Stream 2",
                room_name="Breakout 1",
                capacity=100,
                convener="Mike",
                tohru_room="SA2_Brk1",
            )
            win.active_snapshot = snap

            # Add a delegate and topic ownership
            member = DelegationMember(name="Dieter", color="#2563EB")
            win.delegation_plan.members["Dieter"] = member
            win.delegation_plan.topic_owners["20.6.12"] = TopicOwnership(
                meeting_number="177", ai_number="20.6.12", owner_name="Dieter"
            )

            slot = SessionSlot(
                day_index=2,
                day_name="Wednesday",
                quarter="Q2",
                sub_slot=1,
                stream_number=2,
                raw_text="KI#12 – KI#16 (20.6.12 – 16)",
                items=[
                    AgendaSlotItem(
                        ai_number="20.6.12",
                        display_expr="KI#12 (Voice Services) (20.6.12) [5]",
                        topic_label="Voice Services for 6G",
                        tdoc_count_str="5",
                    )
                ],
            )

            # Add annotation
            win.annotations[slot.slot_key] = UserSessionAnnotation(
                target_type="slot",
                target_key=slot.slot_key,
                comment="Check CR from Nokia",
                assigned_colleague="Colleague A",
            )

            tt = win._build_rich_slot_tooltip(slot, day_idx=2, period="Q2", stream_num=2, tu=1)

            self.assertIn("Breakout 1 (100 seats)", tt)
            self.assertIn("Venue Time", tt)
            self.assertIn("Chair:</b> Mike", tt)
            self.assertIn("TOHRU:</b> SA2_Brk1", tt)
            self.assertIn("Chair's Slot Note:", tt)
            self.assertIn("KI#12 – KI#16", tt)
            self.assertIn("Agenda Items Covered", tt)
            self.assertIn("AI 20.6.12:", tt)
            self.assertIn("Voice Services for 6G", tt)
            self.assertIn("Delegation:", tt)
            self.assertIn("Dieter", tt)
            self.assertIn("Topic Owner", tt)
            self.assertIn("Personal Note:", tt)
            self.assertIn("Check CR from Nokia", tt)
        finally:
            win.close()


if __name__ == "__main__":
    unittest.main()
