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

    def test_delegation_palette_and_solid_tints(self):
        # 1. Verify 10 distinct, valid hex colors in DELEGATE_PALETTE
        self.assertEqual(len(DELEGATE_PALETTE), 10)
        self.assertEqual(len(set(DELEGATE_PALETTE)), 10, "Palette must not contain duplicate hues")
        for col in DELEGATE_PALETTE:
            self.assertTrue(col.startswith("#"))
            self.assertEqual(len(col), 7)
            int(col[1:], 16)  # Must be valid hex

        # 2. Verify DelegationMember bg_tint and border_tint
        m = DelegationMember(name="Martin", color="#2563EB")
        self.assertTrue(m.bg_tint.startswith("#"))
        self.assertEqual(len(m.bg_tint), 7)
        self.assertTrue(m.border_tint.startswith("#"))
        self.assertEqual(len(m.border_tint), 7)

        # 3. Verify all palette colors generate valid solid pastel tints
        for col in DELEGATE_PALETTE:
            bg = DelegationMember.compute_tint(col, 0.12)
            bdr = DelegationMember.compute_tint(col, 0.35)
            self.assertTrue(bg.startswith("#") and len(bg) == 7)
            self.assertTrue(bdr.startswith("#") and len(bdr) == 7)
            # Tint blended with white should be lighter than original color
            r_orig, g_orig, b_orig = int(col[1:3], 16), int(col[3:5], 16), int(col[5:7], 16)
            r_bg, g_bg, b_bg = int(bg[1:3], 16), int(bg[3:5], 16), int(bg[5:7], 16)
            self.assertGreaterEqual(r_bg + g_bg + b_bg, r_orig + g_orig + b_orig)

        # 4. Graceful fallback on malformed hex
        self.assertEqual(DelegationMember.compute_tint("not-a-color"), "#F1F5F9")
        self.assertEqual(DelegationMember.compute_tint(""), "#F1F5F9")

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

    def test_half_quarter_collision_avoidance_and_availability(self):
        """Verifies that sequential half-quarters (TU1 and TU2) in parallel streams do not falsely collide."""
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            m_martin = DelegationMember(name="Martin", color="#2563EB", is_self=True)
            m_thomas = DelegationMember(name="Thomas", color="#D97706")
            win.db.save_delegation_member(win.meeting_number, m_martin)
            win.db.save_delegation_member(win.meeting_number, m_thomas)

            win.db.set_topic_owner(win.meeting_number, "20.6.18", "Martin")
            win.db.set_topic_owner(win.meeting_number, "5.2.1", "Martin")
            win.db.set_topic_owner(win.meeting_number, "7.1.0", "Thomas")
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            # Slot A: Martin in Stream 1, TU 1 (sub_slot=1)
            slot_tu1 = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                sub_slot=1,
                stream_number=1,
                raw_text="AI 5.2.1 (TU1)",
                items=[AgendaSlotItem(ai_number="5.2.1", display_expr="5.2.1", topic_label="Arch TU1")],
            )
            # Slot B: Martin in Stream 2, TU 2 (sub_slot=2) -> sequential, NO physical time overlap!
            slot_tu2 = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                sub_slot=2,
                stream_number=2,
                raw_text="AI 20.6.18 (TU2)",
                items=[AgendaSlotItem(ai_number="20.6.18", display_expr="20.6.18", topic_label="Network AI TU2")],
            )
            # Slot C: Thomas in Stream 1, TU 2 (sub_slot=2)
            slot_thomas_tu2 = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                sub_slot=2,
                stream_number=1,
                raw_text="AI 7.1.0 (Thomas TU2)",
                items=[AgendaSlotItem(ai_number="7.1.0", display_expr="7.1.0", topic_label="Slicing TU2")],
            )

            # Test helper _slots_time_overlap directly
            self.assertFalse(win._slots_time_overlap(slot_tu1, slot_tu2))
            self.assertTrue(win._slots_time_overlap(slot_tu2, slot_thomas_tu2))

            # Full quarter slot overlaps both TU1 and TU2
            slot_full_q = SessionSlot(day_index=1, quarter="Q2", sub_slot=0, stream_number=3)
            self.assertTrue(win._slots_time_overlap(slot_tu1, slot_full_q))
            self.assertTrue(win._slots_time_overlap(slot_tu2, slot_full_q))

            # Mount snapshot with sequential TU1 and TU2
            win.active_snapshot = SessionPlanSnapshot(
                meeting_number=win.meeting_number,
                tdoc_id="S2-2610195",
                streams={1: MeetingStreamInfo(1, "Str 1", "B2"), 2: MeetingStreamInfo(2, "Str 2", "B3")},
                slots=[slot_tu1, slot_tu2, slot_thomas_tu2],
            )

            # Martin is NOT colliding because TU1 and TU2 are distinct time intervals!
            collisions = win._detect_delegation_collisions()
            self.assertEqual(len(collisions), 0, "Sequential half-quarters must not be flagged as collisions!")

            # Availability test: Thomas is busy in TU2, so Thomas MUST BE FREE for slot_tu1 (TU1)!
            avail_for_tu1 = win._compute_delegate_availability(slot_tu1)
            self.assertEqual(avail_for_tu1["Thomas"]["status"], "free", "Thomas in TU2 is available to cover TU1!")

            # But for slot_tu2 (TU2), Thomas IS busy in Stream 1!
            avail_for_tu2 = win._compute_delegate_availability(slot_tu2)
            self.assertEqual(avail_for_tu2["Thomas"]["status"], "busy", "Thomas in TU2 is busy in Stream 1!")

            # Now test true collision in same sub-slot:
            # Change slot_tu2 to sub_slot=1 (both Martin in TU1 on Stream 1 and Stream 2)
            slot_tu2_clash = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                sub_slot=1,
                stream_number=2,
                raw_text="AI 20.6.18 (TU1 Clash)",
                items=[AgendaSlotItem(ai_number="20.6.18", display_expr="20.6.18", topic_label="Network AI TU1")],
            )
            win.active_snapshot.slots = [slot_tu1, slot_tu2_clash]
            collisions_clash = win._detect_delegation_collisions()
            self.assertEqual(len(collisions_clash), 1, "Simultaneous TU1 in streams 1 & 2 must be detected as collision!")

        finally:
            win.close()

    def test_revert_dispatch_restores_topic_owner(self):
        """Verifies _on_revert_dispatch removes session appointment and restores original topic owner."""
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            m_martin = DelegationMember(name="Martin", color="#2563EB", is_self=True)
            m_thomas = DelegationMember(name="Thomas", color="#D97706")
            win.db.save_delegation_member(win.meeting_number, m_martin)
            win.db.save_delegation_member(win.meeting_number, m_thomas)
            win.db.set_topic_owner(win.meeting_number, "20.6.18", "Martin")
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            slot = SessionSlot(
                day_index=1,
                day_name="Tuesday",
                quarter="Q2",
                sub_slot=0,
                stream_number=2,
                raw_text="AI 20.6.18 (Network AI)",
                items=[AgendaSlotItem(ai_number="20.6.18", display_expr="20.6.18", topic_label="Network AI")],
            )
            win.active_snapshot = SessionPlanSnapshot(
                meeting_number=win.meeting_number,
                tdoc_id="S2-2610195",
                streams={2: MeetingStreamInfo(2, "Str 2", "B3")},
                slots=[slot],
            )

            # Initially Martin owns it
            del_init, is_cov_init = win._get_slot_effective_delegate(slot)
            self.assertEqual(del_init, "Martin")
            self.assertFalse(is_cov_init)

            # Appoint Thomas
            win.db.set_session_appointment(
                meeting_number=win.meeting_number,
                slot_key=slot.slot_key,
                appointed_delegate="Thomas",
                original_owner="Martin",
                reason="Ad-hoc cover",
            )
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)
            del_cov, is_cov = win._get_slot_effective_delegate(slot)
            self.assertEqual(del_cov, "Thomas")
            self.assertTrue(is_cov)

            # Revert dispatch
            win._on_revert_dispatch(slot)

            # Must be back to Martin
            del_rev, is_cov_rev = win._get_slot_effective_delegate(slot)
            self.assertEqual(del_rev, "Martin")
            self.assertFalse(is_cov_rev)
            self.assertNotIn(slot.slot_key, win.delegation_plan.appointments)

        finally:
            win.close()

    def test_export_ics_with_delegation_and_delegate_scoping(self):
        """Verifies export_ics properly annotates delegate assignments and filters by target_delegate."""
        slot1 = SessionSlot(
            day_index=0,
            day_name="Monday",
            quarter="Q1",
            start_time_local="09:00",
            end_time_local="10:30",
            stream_number=1,
            raw_text="AI 20.6.18 Network AI Procedures",
            items=[AgendaSlotItem(ai_number="20.6.18", display_expr="20.6.18", topic_label="Network AI")],
        )
        slot2 = SessionSlot(
            day_index=0,
            day_name="Monday",
            quarter="Q2",
            start_time_local="11:00",
            end_time_local="12:30",
            stream_number=2,
            raw_text="AI 5.2.1 Core Architecture",
            items=[AgendaSlotItem(ai_number="5.2.1", display_expr="5.2.1", topic_label="Arch")],
        )
        snap = SessionPlanSnapshot(
            meeting_number="177",
            tdoc_id="S2-2610195",
            streams={1: MeetingStreamInfo(1, "Main", "B2"), 2: MeetingStreamInfo(2, "Breakout 1", "B3")},
            slots=[slot1, slot2],
        )

        plan = DelegationPlan(
            meeting_number="177",
            members={
                "Martin": DelegationMember(name="Martin", color="#2563EB", is_self=True),
                "Anna": DelegationMember(name="Anna", color="#059669"),
            },
            topic_owners={
                "20.6.18": TopicOwnership("177", "20.6.18", "Martin"),
                "5.2.1": TopicOwnership("177", "5.2.1", "Anna"),
            },
            appointments={
                slot2.slot_key: SessionAppointment("177", slot2.slot_key, "Martin", "Anna", "Coverage"),
            },
        )

        # 1. Full export with delegation_plan: both events should contain delegate annotations
        full_ics_path = self.temp_dir / "full_schedule.ics"
        SessionPlanExporter.export_ics(
            snapshot=snap,
            output_path=full_ics_path,
            meeting_start_date_str="2026-10-12",
            meeting_tz_offset_min=120,
            delegation_plan=plan,
        )
        content_full = full_ics_path.read_text(encoding="utf-8")
        self.assertIn("SUMMARY:[Str 1] [Martin] AI 20.6.18", content_full)
        self.assertIn("SUMMARY:[Str 2] [Martin (cov)] AI 5.2.1", content_full)
        self.assertIn("Delegate: Martin (Topic Owner)", content_full)
        self.assertIn("Delegate: Martin (Ad-hoc coverage)", content_full)

        # 2. Scoped export for Anna: slot2 was covered by Martin, so Anna has 0 assigned slots
        anna_ics_path = self.temp_dir / "anna_schedule.ics"
        SessionPlanExporter.export_ics(
            snapshot=snap,
            output_path=anna_ics_path,
            meeting_start_date_str="2026-10-12",
            meeting_tz_offset_min=120,
            delegation_plan=plan,
            target_delegate="Anna",
        )
        content_anna = anna_ics_path.read_text(encoding="utf-8")
        self.assertNotIn("BEGIN:VEVENT", content_anna, "Anna has no scheduled slots since slot2 is covered by Martin")
        self.assertIn("X-WR-CALNAME:SA2 #177 - Anna (S2-2610195)", content_anna)

        # 3. Scoped export for Martin: both slot1 and slot2 are assigned to Martin
        martin_ics_path = self.temp_dir / "martin_schedule.ics"
        SessionPlanExporter.export_ics(
            snapshot=snap,
            output_path=martin_ics_path,
            meeting_start_date_str="2026-10-12",
            meeting_tz_offset_min=120,
            delegation_plan=plan,
            target_delegate="Martin",
        )
        content_martin = martin_ics_path.read_text(encoding="utf-8")
        self.assertEqual(content_martin.count("BEGIN:VEVENT"), 2)
        self.assertIn("X-WR-CALNAME:SA2 #177 - Martin (S2-2610195)", content_martin)



if __name__ == "__main__":
    unittest.main()
