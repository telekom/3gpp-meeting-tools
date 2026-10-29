# --- File: tests/test_session_plan_delegation.py ---
"""
Unit and integration tests for Session Plan Delegation & Topic Distribution:
1. DelegationMember, TopicOwnership, SessionAppointment, and DelegationPlan domain models.
2. SQLite persistence in SessionPlanDatabase (roster, topic owners, session appointments, import/export dict).
3. Exporter serialization, embedded HTML parsing, and email summary generation.
4. Real-time availability calculation and parallel stream collision detection.
5. SessionPlanWindow UI components and 5-view stack integration.
"""
import shutil
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan_models import (
    DELEGATE_PALETTE,
    AgendaSlotItem,
    DelegationMember,
    DelegationPlan,
    MeetingStreamInfo,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
)
from modules.meetings.ui.session_plan_dialogs import (
    DelegationImportDialog,
    DelegationMemberDialog,
    SessionDispatchDialog,
)
from modules.meetings.ui.session_plan_window import SessionPlanWindow

# Ensure QApplication singleton for GUI tests
_app = QApplication.instance() or QApplication([])


class TestSessionPlanDelegationModelsAndDB(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.db_path = self.temp_dir / "session_plans.db"
        self.db = SessionPlanDatabase(self.db_path)
        self.meeting_number = "177"

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_delegation_member_initials(self):
        m1 = DelegationMember(name="Martin", color="#2563EB")
        self.assertEqual(m1.initials, "MA")

        m2 = DelegationMember(name="Anna Schmidt", color="#059669")
        self.assertEqual(m2.initials, "AS")

        m3 = DelegationMember(name="Thomas Alva Edison", color="#D97706")
        self.assertEqual(m3.initials, "TE")

    def test_delegation_roster_crud(self):
        m1 = DelegationMember(name="Martin", email="martin@telekom.com", color="#2563EB", is_self=True, notes="Lead")
        m2 = DelegationMember(name="Anna", email="anna@telekom.com", color="#059669", is_self=False, notes="Core")

        self.db.save_delegation_member(self.meeting_number, m1)
        self.db.save_delegation_member(self.meeting_number, m2)

        roster = self.db.get_delegation_members(self.meeting_number)
        self.assertEqual(len(roster), 2)
        self.assertIn("Martin", roster)
        self.assertIn("Anna", roster)
        self.assertTrue(roster["Martin"].is_self)
        self.assertFalse(roster["Anna"].is_self)

        # Switch self identity to Anna
        self.db.set_self_delegate(self.meeting_number, "Anna")
        roster2 = self.db.get_delegation_members(self.meeting_number)
        self.assertFalse(roster2["Martin"].is_self)
        self.assertTrue(roster2["Anna"].is_self)

        # Delete Martin
        self.db.delete_delegation_member(self.meeting_number, "Martin")
        roster3 = self.db.get_delegation_members(self.meeting_number)
        self.assertEqual(len(roster3), 1)
        self.assertNotIn("Martin", roster3)

    def test_topic_ownership_crud(self):
        self.db.set_topic_owner(self.meeting_number, "20.6.18", "Martin", notes="Primary lead")
        self.db.set_topic_owner(self.meeting_number, "5.2.1", "Anna", notes="Architecture")

        topics = self.db.get_topic_owners(self.meeting_number)
        self.assertEqual(len(topics), 2)
        self.assertEqual(topics["20.6.18"].owner_name, "Martin")
        self.assertEqual(topics["5.2.1"].owner_name, "Anna")

        # Update owner
        self.db.set_topic_owner(self.meeting_number, "20.6.18", "Thomas")
        topics2 = self.db.get_topic_owners(self.meeting_number)
        self.assertEqual(topics2["20.6.18"].owner_name, "Thomas")

        # Delete
        self.db.delete_topic_owner(self.meeting_number, "5.2.1")
        topics3 = self.db.get_topic_owners(self.meeting_number)
        self.assertNotIn("5.2.1", topics3)

    def test_session_appointments_crud(self):
        slot_key = "Tuesday|Q2|0|2"
        self.db.set_session_appointment(
            meeting_number=self.meeting_number,
            slot_key=slot_key,
            appointed_delegate="Thomas",
            original_owner="Martin",
            reason="Collision with Stream 1",
        )

        appts = self.db.get_session_appointments(self.meeting_number)
        self.assertEqual(len(appts), 1)
        self.assertIn(slot_key, appts)
        self.assertEqual(appts[slot_key].appointed_delegate, "Thomas")
        self.assertEqual(appts[slot_key].original_owner, "Martin")

        # Delete appointment
        self.db.delete_session_appointment(self.meeting_number, slot_key)
        self.assertEqual(len(self.db.get_session_appointments(self.meeting_number)), 0)

    def test_plan_export_and_import_dict(self):
        # Create full plan
        m1 = DelegationMember(name="Martin", color="#2563EB", is_self=True)
        m2 = DelegationMember(name="Anna", color="#059669", is_self=False)
        self.db.save_delegation_member(self.meeting_number, m1)
        self.db.save_delegation_member(self.meeting_number, m2)
        self.db.set_topic_owner(self.meeting_number, "20.6.18", "Martin")
        self.db.set_session_appointment(self.meeting_number, "Tue|Q2|0|2", "Anna", "Martin", "Collision")

        data = self.db.export_delegation_plan_dict(self.meeting_number)
        self.assertEqual(data["meeting_number"], "177")
        self.assertEqual(len(data["members"]), 2)
        self.assertIn("20.6.18", data["topic_owners"])
        self.assertIn("Tue|Q2|0|2", data["appointments"])

        # Import into clean meeting
        mtg2 = "178"
        count = self.db.import_delegation_plan_dict(mtg2, data)
        self.assertEqual(count, 4)  # 2 members + 1 topic + 1 appt

        plan2 = self.db.get_delegation_plan(mtg2)
        self.assertEqual(len(plan2.members), 2)
        self.assertEqual(plan2.topic_owners["20.6.18"].owner_name, "Martin")
        self.assertEqual(plan2.appointments["Tue|Q2|0|2"].appointed_delegate, "Anna")


class TestSessionPlanExporterDelegation(unittest.TestCase):
    def test_json_serialization_and_parsing(self):
        plan = DelegationPlan(
            meeting_number="177",
            members={
                "Martin": DelegationMember(name="Martin", color="#2563EB", is_self=True),
                "Anna": DelegationMember(name="Anna", color="#059669", is_self=False),
            },
            topic_owners={
                "20.6.18": TopicOwnership(meeting_number="177", ai_number="20.6.18", owner_name="Martin"),
            },
            appointments={
                "Tue|Q2|0|2": SessionAppointment(
                    meeting_number="177", slot_key="Tue|Q2|0|2", appointed_delegate="Anna", original_owner="Martin", reason="Clash"
                ),
            },
        )

        json_str = SessionPlanExporter.serialize_delegation_plan_json(plan)
        self.assertIn("3gpp_delegation_plan_version", json_str)
        self.assertIn("Martin", json_str)

        # Parse from raw JSON
        parsed_raw = SessionPlanExporter.parse_delegation_plan_json(json_str)
        self.assertIsNotNone(parsed_raw)
        self.assertEqual(parsed_raw["meeting_number"], "177")
        self.assertEqual(len(parsed_raw["members"]), 2)

        # Parse from embedded HTML comment
        email_with_comment = f"""
        <html><body><p>Here is the plan.</p>
        <!-- 3GPP-DELEGATION-PLAN: {json_str} -->
        </body></html>
        """
        parsed_comment = SessionPlanExporter.parse_delegation_plan_json(email_with_comment)
        self.assertIsNotNone(parsed_comment)
        self.assertEqual(parsed_comment["meeting_number"], "177")

    def test_email_summary_generation(self):
        snapshot = SessionPlanSnapshot(
            meeting_number="177",
            tdoc_id="S2-2610195",
            streams={
                1: MeetingStreamInfo(stream_number=1, stream_name="Stream 1", room_name="Room B2"),
                2: MeetingStreamInfo(stream_number=2, stream_name="Stream 2", room_name="Room B3"),
            },
            slots=[
                SessionSlot(
                    day_name="Tuesday",
                    quarter="Q1",
                    stream_number=1,
                    raw_text="AI 20.6.18 Network AI",
                    items=[AgendaSlotItem(ai_number="20.6.18", display_expr="AI 20.6.18", topic_label="Network AI")],
                ),
                SessionSlot(
                    day_name="Tuesday",
                    quarter="Q2",
                    stream_number=2,
                    raw_text="AI 20.6.18 Network AI (Cont)",
                    items=[AgendaSlotItem(ai_number="20.6.18", display_expr="AI 20.6.18", topic_label="Network AI")],
                ),
            ],
        )

        plan = DelegationPlan(
            meeting_number="177",
            members={
                "Martin": DelegationMember(name="Martin", color="#2563EB", is_self=True, notes="Lead"),
                "Anna": DelegationMember(name="Anna", color="#059669", is_self=False, notes="Delegate"),
            },
            topic_owners={
                "20.6.18": TopicOwnership(meeting_number="177", ai_number="20.6.18", owner_name="Martin"),
            },
            appointments={
                "Tuesday|Q2|0|2": SessionAppointment(
                    meeting_number="177",
                    slot_key="Tuesday|Q2|0|2",
                    appointed_delegate="Anna",
                    original_owner="Martin",
                    reason="Covering for Martin due to parallel meeting",
                ),
            },
        )

        subject, html_body, plain_text = SessionPlanExporter.generate_delegation_email_summary(
            snapshot=snapshot,
            plan=plan,
            meeting_tz_offset_min=120,
            tz_name="Europe/Prague",
        )

        self.assertIn("[SA2#177]", subject)
        self.assertIn("Martin", html_body)
        self.assertIn("Anna", html_body)
        self.assertIn("<!-- 3GPP-DELEGATION-PLAN:", html_body)
        self.assertIn("Covering for Martin", html_body)
        self.assertIn("--- ROSTER ---", plain_text)


class TestSessionPlanWindowDelegationIntegration(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.mtg_info = {
            "meeting_number": "177",
            "location": "Prague, Czechia",
            "start_date": "2026-10-12",
        }

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_window_initialization_with_5_views(self):
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            self.assertEqual(win.view_stack.count(), 5)
            self.assertEqual(win.btn_team.text(), "👥 Team Allocations (Clean)")
            self.assertIn("Delegate Agenda", win.btn_tracked.text())
        finally:
            win.close()

    def test_collision_detection_and_dispatch(self):
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            # 1. Setup roster with Martin, Anna, and Thomas
            m_martin = DelegationMember(name="Martin", color="#2563EB", is_self=True)
            m_anna = DelegationMember(name="Anna", color="#059669")
            m_thomas = DelegationMember(name="Thomas", color="#D97706")

            win.db.save_delegation_member(win.meeting_number, m_martin)
            win.db.save_delegation_member(win.meeting_number, m_anna)
            win.db.save_delegation_member(win.meeting_number, m_thomas)

            # 2. Setup topic owners: Martin owns both 20.6.18 and 5.2.1
            win.db.set_topic_owner(win.meeting_number, "20.6.18", "Martin")
            win.db.set_topic_owner(win.meeting_number, "5.2.1", "Martin")
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            # 3. Simulate active snapshot with parallel sessions on Tuesday Q2 in Stream 1 and Stream 2
            slot1 = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                stream_number=1,
                raw_text="AI 5.2.1 (Arch)",
                items=[AgendaSlotItem(ai_number="5.2.1", display_expr="5.2.1", topic_label="Arch")],
            )
            slot2 = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                stream_number=2,
                raw_text="AI 20.6.18 (Network AI)",
                items=[AgendaSlotItem(ai_number="20.6.18", display_expr="20.6.18", topic_label="Network AI")],
            )
            win.active_snapshot = SessionPlanSnapshot(
                meeting_number=win.meeting_number,
                tdoc_id="S2-2610195",
                streams={1: MeetingStreamInfo(1, "Str 1", "B2"), 2: MeetingStreamInfo(2, "Str 2", "B3")},
                slots=[slot1, slot2],
            )

            # Check effective delegates: both Martin
            del1, _ = win._get_slot_effective_delegate(slot1)
            del2, _ = win._get_slot_effective_delegate(slot2)
            self.assertEqual(del1, "Martin")
            self.assertEqual(del2, "Martin")

            # Detect collisions: Martin is double-booked on (Martin, 1, 'Q2')
            collisions = win._detect_delegation_collisions()
            self.assertEqual(len(collisions), 1)
            self.assertIn(("Martin", 1, "Q2"), collisions)

            # Check team availability for slot2:
            # Martin is busy in Stream 1; Anna and Thomas are free!
            avail = win._compute_delegate_availability(slot2)
            self.assertEqual(avail["Martin"]["status"], "busy")
            self.assertEqual(avail["Anna"]["status"], "free")
            self.assertEqual(avail["Thomas"]["status"], "free")

            # Appoint Thomas on slot2
            win.db.set_session_appointment(
                meeting_number=win.meeting_number,
                slot_key=slot2.slot_key,
                appointed_delegate="Thomas",
                original_owner="Martin",
                reason="Parallel stream collision",
            )
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            # Now effective delegate on slot2 is Thomas (is_cov=True)
            del2_after, is_cov = win._get_slot_effective_delegate(slot2)
            self.assertEqual(del2_after, "Thomas")
            self.assertTrue(is_cov)

            # Re-check collisions: completely resolved!
            collisions_after = win._detect_delegation_collisions()
            self.assertEqual(len(collisions_after), 0)

        finally:
            win.close()


if __name__ == "__main__":
    unittest.main()
