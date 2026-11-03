# --- File: tests/test_session_plan_delegation.py ---
"""
Unit and integration tests for Session Plan Delegation & Topic Distribution:
1. DelegationMember, TopicOwnership, SessionAppointment, and DelegationPlan domain models.
2. SQLite persistence in SessionPlanDatabase (roster, topic owners, session appointments, import/export dict).
3. Exporter serialization, embedded HTML parsing, and email summary generation.
4. Real-time availability calculation and parallel stream collision detection.
5. SessionPlanWindow UI components and 5-view stack integration.
"""
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan.session_plan_models import (
    DELEGATE_PALETTE,
    AgendaSlotItem,
    DelegationMember,
    DelegationPlan,
    MeetingStreamInfo,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
    ai_number_in_text,
    compute_unique_initials,
    resolve_slot_effective_delegates,
)
from modules.meetings.ui.delegate_owner_picker import DelegateOwnerPicker
from modules.meetings.ui.session_plan_dialogs import (
    MultiDelegateSelectDialog,
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

    def test_multi_delegate_models_and_all_expansion(self):
        # 1. TopicOwnership with single, multi, and ALL
        to1 = TopicOwnership("177", "20.6.18", "Martin")
        self.assertFalse(to1.is_all())
        self.assertEqual(to1.get_owners(), ["Martin"])

        to2 = TopicOwnership("177", "20.6.0", "Martin, Anna")
        self.assertFalse(to2.is_all())
        self.assertEqual(to2.get_owners(), ["Martin", "Anna"])

        to3 = TopicOwnership("177", "20.6.0", "ALL")
        self.assertTrue(to3.is_all())
        self.assertEqual(to3.get_owners(["Martin", "Anna", "Thomas"]), ["Martin", "Anna", "Thomas"])

        to4 = TopicOwnership("177", "20.6.0", "👥 All Team")
        self.assertTrue(to4.is_all())
        self.assertEqual(to4.get_owners(["Martin", "Anna"]), ["Martin", "Anna"])

        # 2. SessionAppointment with single, multi, and ALL
        ap1 = SessionAppointment("177", "Tuesday|Q1|0|1", "Thomas")
        self.assertFalse(ap1.is_all())
        self.assertEqual(ap1.get_appointed_delegates(), ["Thomas"])

        ap2 = SessionAppointment("177", "Tuesday|Q1|0|1", "Martin, Anna")
        self.assertFalse(ap2.is_all())
        self.assertEqual(ap2.get_appointed_delegates(), ["Martin", "Anna"])

        ap3 = SessionAppointment("177", "Tuesday|Q1|0|1", "ALL")
        self.assertTrue(ap3.is_all())
        self.assertEqual(ap3.get_appointed_delegates(["Martin", "Anna"]), ["Martin", "Anna"])

    def test_multi_delegate_select_dialog(self):
        members = [
            DelegationMember(name="Martin", color="#2563EB", is_self=True),
            DelegationMember(name="Anna", color="#047857"),
            DelegationMember(name="Thomas", color="#7C3AED"),
        ]

        # 1. Initialize with Martin and Anna
        dlg = MultiDelegateSelectDialog("AI 20.6.0: General Topics", members, "Martin, Anna")
        self.assertIn("Martin", dlg.get_selected_delegates())
        self.assertIn("Anna", dlg.get_selected_delegates())
        self.assertNotIn("Thomas", dlg.get_selected_delegates())
        self.assertEqual(dlg.get_owner_string(), "Martin, Anna")
        self.assertFalse(dlg.chk_all.isChecked())

        # 2. Toggle All Team
        dlg.chk_all.setChecked(True)
        self.assertEqual(len(dlg.get_selected_delegates()), 3)
        self.assertEqual(dlg.get_owner_string(), "ALL")

        # 3. Uncheck Thomas -> should revert master checkbox and return "Martin, Anna"
        dlg._checkboxes["Thomas"].setChecked(False)
        self.assertFalse(dlg.chk_all.isChecked())
        self.assertEqual(dlg.get_owner_string(), "Martin, Anna")

        # 4. Clear all
        dlg._set_all_individual_checked(False)
        self.assertEqual(dlg.get_owner_string(), "")

    def test_window_multi_delegate_slot_and_clash_detection(self):
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            mtg_num = win.meeting_number
            slot1 = SessionSlot(
                id=101,
                day_index=1,
                day_name="Tuesday",
                quarter="Q1",
                sub_slot=1,  # TU1
                stream_number=1,
                raw_text="AI 20.6.0 General Architecture Wrap-up",
                items=[AgendaSlotItem(ai_number="20.6.0", display_expr="AI 20.6.0", topic_label="General")],
            )
            slot2 = SessionSlot(
                id=102,
                day_index=1,
                day_name="Tuesday",
                quarter="Q1",
                sub_slot=1,  # TU1 - SIMULTANEOUS COLLISION with Slot 1!
                stream_number=2,
                raw_text="AI 5.2.1 Core Network Features",
                items=[AgendaSlotItem(ai_number="5.2.1", display_expr="AI 5.2.1", topic_label="Core")],
            )

            snap = SessionPlanSnapshot(
                id=1,
                meeting_number=mtg_num,
                tdoc_id="S2-2610195",
                slots=[slot1, slot2],
            )
            win.active_snapshot = snap

            # Setup delegation plan: Martin and Anna in roster
            # AI 20.6.0 is assigned to ALL (both Martin and Anna)
            # AI 5.2.1 is assigned to Martin
            m1 = DelegationMember(name="Martin", color="#2563EB", is_self=True)
            m2 = DelegationMember(name="Anna", color="#047857")
            win.db.save_delegation_member(mtg_num, m1)
            win.db.save_delegation_member(mtg_num, m2)
            win.db.set_topic_owner(mtg_num, "20.6.0", "ALL")
            win.db.set_topic_owner(mtg_num, "5.2.1", "Martin")
            win.delegation_plan = win.db.get_delegation_plan(mtg_num)

            # 1. Effective delegate resolution
            dels1, is_cov1, is_all1 = win._get_slot_effective_delegates(slot1)
            self.assertTrue(is_all1)
            self.assertIn("Martin", dels1)
            self.assertIn("Anna", dels1)

            dels2, is_cov2, is_all2 = win._get_slot_effective_delegates(slot2)
            self.assertFalse(is_all2)
            self.assertEqual(dels2, ["Martin"])

            # 2. Collision detection
            # Martin has both Slot 1 (Str 1 TU1) and Slot 2 (Str 2 TU1) -> Martin has a collision!
            # Anna is only on Slot 1 (Str 1 TU1) -> Anna has NO collision!
            collisions = win._detect_delegation_collisions()
            martin_clashes = collisions.get(("Martin", 1, "Q1"), [])
            self.assertEqual(len(martin_clashes), 2, "Martin must have a parallel collision between Str 1 and Str 2")

            anna_clashes = collisions.get(("Anna", 1, "Q1"), [])
            self.assertEqual(len(anna_clashes), 0, "Anna must NOT have a collision since she is only in Str 1")

            # 3. Availability engine for appointing coverage for Slot 2
            avail = win._compute_delegate_availability(slot2)
            self.assertEqual(avail["Martin"]["status"], "busy", "Martin is already on Str 1 during this quarter")
            self.assertEqual(avail["Anna"]["status"], "busy", "Anna is also attending Str 1 (20.6.0 ALL) during this quarter")

            # 4. Delegate Agenda view filtering
            # First populate combo with new delegation roster
            win._render_tracked_topics()

            # Viewing 'Anna' must include Slot 1 (because 20.6.0 is for ALL), but NOT Slot 2
            idx_anna = win.delegate_agenda_combo.findData("Anna")
            self.assertGreaterEqual(idx_anna, 0)
            win.delegate_agenda_combo.setCurrentIndex(idx_anna)
            win._render_tracked_topics()
            self.assertEqual(win.tracked_table.rowCount(), 1)
            self.assertIn("All Team", win.tracked_table.item(0, 4).text())

            # Viewing 'Martin' must include both Slot 1 and Slot 2
            idx_martin = win.delegate_agenda_combo.findData("Martin")
            self.assertGreaterEqual(idx_martin, 0)
            win.delegate_agenda_combo.setCurrentIndex(idx_martin)
            win._render_tracked_topics()
            self.assertEqual(win.tracked_table.rowCount(), 2)

            # 5. Email Export includes All Team in master topic ownership
            subj, html_body, plain = SessionPlanExporter.generate_delegation_email_summary(
                snapshot=snap,
                plan=win.delegation_plan,
                meeting_tz_offset_min=120,
                tz_name="Venue Time",
            )
            self.assertIn("All Team", html_body)
            self.assertIn("AI 20.6.0", html_body)
            self.assertIn("AI 20.6.0: Agenda Item 20.6.0 -> All Team", plain)
        finally:
            win.close()

    def test_boundary_aware_ai_matching_prevents_false_clashes(self):
        """
        Regression test for the 20.6.2 vs 20.6.24 false clash bug:
        Assigning 20.6.2 (or 20.6.1) to Josep must NEVER attribute unassigned
        20.6.24, 20.6.22, or 20.6.17 slots to Josep via raw_text substring matching.
        """
        self.assertTrue(ai_number_in_text("20.6.2", "KI#2 (SBA) (20.6.2) [7]"))
        self.assertFalse(ai_number_in_text("20.6.2", "KI#24 (IoT) (20.6.24)"))
        self.assertFalse(ai_number_in_text("20.6.2", "KI#22 (Computing) (20.6.22)"))
        self.assertFalse(ai_number_in_text("20.6.1", "KI#17 (IWK) (20.6.17)"))
        self.assertFalse(ai_number_in_text("20.6.1", "KI#1.1 (NAS) (20.6.1.1)"))

        plan = DelegationPlan(
            meeting_number="177",
            members={"Josep": DelegationMember(name="Josep", color="#7C3AED", is_self=True)},
            topic_owners={
                "20.6.1": TopicOwnership("177", "20.6.1", "Josep"),
                "20.6.2": TopicOwnership("177", "20.6.2", "Josep"),
            },
        )
        slot_sba = SessionSlot(
            day_index=2,
            day_name="Wednesday",
            quarter="Q2",
            sub_slot=2,
            stream_number=1,
            raw_text="KI#2 (SBA) (20.6.2) [7]",
            items=[AgendaSlotItem(ai_number="20.6.2", display_expr="KI#2", topic_label="SBA")],
        )
        slot_iot = SessionSlot(
            day_index=2,
            day_name="Wednesday",
            quarter="Q2",
            sub_slot=2,
            stream_number=2,
            raw_text="KI#24 (IoT) (20.6.24)",
            items=[AgendaSlotItem(ai_number="20.6.24", display_expr="KI#24", topic_label="IoT")],
        )

        dels_sba, _, _ = resolve_slot_effective_delegates(slot_sba, plan)
        dels_iot, _, _ = resolve_slot_effective_delegates(slot_iot, plan)
        self.assertEqual(dels_sba, ["Josep"])
        self.assertEqual(dels_iot, [], "Unassigned 20.6.24 must not inherit owner from 20.6.2!")

    def test_compute_unique_initials_and_picker_widget(self):
        ini = compute_unique_initials(["Dieter", "Josep", "Piotr", "Jordi"])
        self.assertEqual(ini["Dieter"], "DI")
        self.assertEqual(ini["Piotr"], "PI")
        self.assertNotEqual(ini["Josep"], ini["Jordi"], "Colliding 2-letter initials must be disambiguated")

        members = [
            DelegationMember(name="Dieter", color="#2563EB"),
            DelegationMember(name="Josep", color="#7C3AED", is_self=True),
            DelegationMember(name="Piotr", color="#047857"),
        ]
        picker = DelegateOwnerPicker(members, "Josep", initials=ini)
        self.assertEqual(picker.owner_string(), "Josep")
        self.assertIn("Josep", picker.currentText())

        # Multi-select via checkboxes
        picker._checkboxes["Dieter"].setChecked(True)
        self.assertIn("Dieter", picker.selected_delegates())
        self.assertIn("Josep", picker.selected_delegates())
        self.assertEqual(picker.owner_string(), "Dieter, Josep")

        # Checking all 3 collapses to 'ALL'
        picker._checkboxes["Piotr"].setChecked(True)
        self.assertEqual(picker.owner_string(), "ALL")

        # Clear action resets to unassigned
        picker._on_clear()
        self.assertEqual(picker.owner_string(), "")

    def test_overview_multi_select_focus_and_initials_badges(self):
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            m1 = DelegationMember(name="Josep", color="#7C3AED", is_self=True)
            m2 = DelegationMember(name="Dieter", color="#2563EB")
            m3 = DelegationMember(name="Piotr", color="#047857")
            for m in (m1, m2, m3):
                win.db.save_delegation_member(win.meeting_number, m)
            win.db.set_topic_owner(win.meeting_number, "20.6.1", "Josep")
            win.db.set_topic_owner(win.meeting_number, "20.6.20", "Dieter")
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            slot = SessionSlot(
                day_index=0,
                day_name="Monday",
                quarter="Q2",
                sub_slot=1,
                stream_number=1,
                raw_text="KI#1.1 (NAS) (20.6.1) [9]",
                items=[AgendaSlotItem(ai_number="20.6.1", display_expr="KI#1.1", topic_label="NAS")],
            )
            win.active_snapshot = SessionPlanSnapshot(
                meeting_number=win.meeting_number,
                tdoc_id="S2-2610195",
                slots=[slot],
            )
            win._render_current_views()

            # Checkboxes created for each member
            self.assertEqual(len(win.dynamic_delegate_chips), 3)

            # Multi-select Josep + Dieter simultaneously
            win._on_focus_delegate_toggled("Josep", True)
            win._on_focus_delegate_toggled("Dieter", True)
            self.assertEqual(win.grid_focus_delegates, {"Josep", "Dieter"})
            self.assertEqual(win.grid_highlight_delegate, "MULTI")

            # Reset via All Team
            win._on_delegate_highlight_clicked("ALL")
            self.assertEqual(win.grid_focus_delegates, set())
            self.assertEqual(win.grid_highlight_delegate, "ALL")
        finally:
            win.close()

    def test_interactive_html_export_overview_and_cards(self):
        """Verifies exported HTML contains both 5-day Overview Grid and Mobile Day Cards with delegate toggling."""
        snap = SessionPlanSnapshot(
            meeting_number="177",
            tdoc_id="S2-2610195",
            snapshot_label="Snap 1",
            slots=[
                SessionSlot(
                    day_index=0,
                    day_name="Monday",
                    quarter="Q1",
                    sub_slot=0,
                    stream_number=1,
                    raw_text="AI 20.6.0 General Architecture Wrap-up",
                    items=[AgendaSlotItem(ai_number="20.6.0", display_expr="AI 20.6.0", topic_label="General")],
                ),
                SessionSlot(
                    day_index=1,
                    day_name="Tuesday",
                    quarter="Q2",
                    sub_slot=1,
                    stream_number=2,
                    raw_text="AI 5.2.1 Core Network Features",
                    items=[AgendaSlotItem(ai_number="5.2.1", display_expr="AI 5.2.1", topic_label="Core")],
                ),
            ],
        )

        plan = DelegationPlan(
            meeting_number="177",
            members={
                "Anna": DelegationMember(name="Anna", color="#059669", is_self=True),
                "Thomas": DelegationMember(name="Thomas", color="#7C3AED", is_self=False),
            },
            topic_owners={
                "20.6.0": TopicOwnership("177", "20.6.0", "ALL"),
                "5.2.1": TopicOwnership("177", "5.2.1", "Anna"),
            },
        )

        out_html = self.temp_dir / "test_session_plan.html"
        SessionPlanExporter.export_mobile_html(
            snapshot=snap,
            output_path=out_html,
            meeting_tz_offset_min=120,
            tz_name="Venue Time (Madrid)",
            meeting_dates_str="2026-08-17",
            delegation_plan=plan,
        )

        self.assertTrue(out_html.exists())
        content = out_html.read_text(encoding="utf-8")

        # 1. Dual View Architecture
        self.assertIn('id="overviewView"', content)
        self.assertIn('class="overview-table"', content)
        self.assertIn('id="cardsView"', content)
        self.assertIn('class="cards-content"', content)
        self.assertIn("switchView('grid')", content)
        self.assertIn("switchView('cards')", content)

        # 2. Interactive Delegate Toggling Pills
        self.assertIn('data-del="ALL"', content)
        self.assertIn('data-del="Anna"', content)
        self.assertIn('data-del="Thomas"', content)
        self.assertIn("selectDelegate(this.getAttribute('data-del'))", content)

        # 3. Data attributes for client-side filtering
        self.assertIn('data-delegates="Anna"', content)
        self.assertIn('data-is-all="1"', content)

        # 4. No "myself" or "(Me)" tags anywhere in the exported HTML
        self.assertNotIn("(Myself)", content)
        self.assertNotIn("(Me)", content)
        self.assertNotIn("myself", content.lower())

        # 5. Client-Side RFC 5545 .ics Calendar Download
        self.assertIn('id="btnDownloadIcs"', content)
        self.assertIn("downloadSelectedCalendar()", content)
        self.assertIn("const CAL_EVENTS =", content)

        # 6. Parallel Stream Clash Detection
        self.assertIn('id="clashWarningBadge"', content)
        self.assertIn("detectDelegateClashes", content)

        # 7. Live Venue Status Indicator ("Now / Up Next")
        self.assertIn('id="liveVenueBanner"', content)
        self.assertIn("checkLiveVenueStatus", content)

        # 8. Automated JavaScript Syntax and Interactive Function Integrity Verification
        script_matches = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
        self.assertGreaterEqual(len(script_matches), 2)
        full_js = "\n".join(script_matches)

        node_bin = shutil.which("node")
        if node_bin:
            harness = f"""
const mockElement = () => ({{
  classList: {{ toggle: () => {{}}, remove: () => {{}}, add: () => {{}} }},
  style: {{}},
  getAttribute: () => '',
  querySelectorAll: () => [],
  querySelector: () => null,
  addEventListener: () => {{}},
  value: ''
}});
global.document = {{
  getElementById: () => mockElement(),
  querySelectorAll: () => [],
  querySelector: () => mockElement(),
  body: {{ appendChild: () => {{}}, removeChild: () => {{}} }}
}};
global.window = {{ innerWidth: 1024 }};
global.setInterval = () => {{}};
global.Blob = class {{}};
global.URL = {{ createObjectURL: () => '', revokeObjectURL: () => {{}} }};

{full_js}

if (typeof switchView !== 'function') throw new Error('switchView is not defined');
if (typeof selectDelegate !== 'function') throw new Error('selectDelegate is not defined');
if (typeof switchDay !== 'function') throw new Error('switchDay is not defined');
if (typeof downloadSelectedCalendar !== 'function') throw new Error('downloadSelectedCalendar is not defined');
if (typeof detectDelegateClashes !== 'function') throw new Error('detectDelegateClashes is not defined');
if (typeof toggleStream !== 'function') throw new Error('toggleStream is not defined');
if (typeof toggleBreaks !== 'function') throw new Error('toggleBreaks is not defined');
if (typeof toggleEmptyBreaks !== 'function') throw new Error('toggleEmptyBreaks is not defined');

// Invoke interactive handlers to ensure zero runtime scope exceptions
switchView('cards');
selectDelegate('ALL');
switchDay('Monday');
detectDelegateClashes('ALL');
"""
            temp_js = self.temp_dir / "verify_bundle.js"
            temp_js.write_text(harness, encoding="utf-8")
            res = subprocess.run([node_bin, str(temp_js)], capture_output=True, text=True, encoding="utf-8", errors="replace")
            self.assertEqual(res.returncode, 0, f"Node.js execution failed:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}")

    def test_overview_grid_topic_owner_direct_assignment(self):
        """Verifies right-click topic owner assignment from the desktop Overview Grid cell context menu."""
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            m1 = DelegationMember(name="Anna", color="#059669")
            m2 = DelegationMember(name="Thomas", color="#7C3AED")
            win.db.save_delegation_member(win.meeting_number, m1)
            win.db.save_delegation_member(win.meeting_number, m2)
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            # Assign topic 5.2.1 to Anna directly
            win._set_topic_owner_direct("5.2.1", "Anna")
            plan = win.db.get_delegation_plan(win.meeting_number)
            self.assertIn("5.2.1", plan.topic_owners)
            self.assertEqual(plan.topic_owners["5.2.1"].owner_name, "Anna")

            # Assign topic 5.2.1 to Whole Team (ALL)
            win._set_topic_owner_direct("5.2.1", "ALL")
            plan = win.db.get_delegation_plan(win.meeting_number)
            self.assertTrue(plan.topic_owners["5.2.1"].is_all())

            # Unassign topic 5.2.1
            win._set_topic_owner_direct("5.2.1", "")
            plan = win.db.get_delegation_plan(win.meeting_number)
            self.assertNotIn("5.2.1", plan.topic_owners)
        finally:
            win.close()

    def test_delegation_no_myself_tags_in_ui(self):
        """Verifies UI dialogs and pickers do not append (Me) or (Myself) tags to delegate names."""
        from modules.meetings.ui.session_plan_dialogs import DelegationMemberDialog, MultiDelegateSelectDialog

        m1 = DelegationMember(name="Anna", color="#059669", is_self=True)
        m2 = DelegationMember(name="Thomas", color="#7C3AED", is_self=False)

        # MultiDelegateSelectDialog
        multi_dlg = MultiDelegateSelectDialog("AI 20.6.0", [m1, m2], "Anna")
        self.assertEqual(multi_dlg._checkboxes["Anna"].text(), "Anna")
        self.assertEqual(multi_dlg._checkboxes["Thomas"].text(), "Thomas")
        self.assertNotIn("(Me)", multi_dlg._checkboxes["Anna"].text())
        self.assertNotIn("(Myself)", multi_dlg._checkboxes["Anna"].text())

        # DelegationMemberDialog
        mem_dlg = DelegationMemberDialog(parent=None, member=m1, existing_names=[])
        self.assertEqual(mem_dlg.preview_lbl.text().strip(), "● Anna")
        self.assertNotIn("(Me)", mem_dlg.preview_lbl.text())
        self.assertNotIn("(Myself)", mem_dlg.preview_lbl.text())
        self.assertFalse(hasattr(mem_dlg, "self_chk"))

    def test_persistent_ai_assignments_across_meetings_and_session_overrides(self):
        """
        Verifies that Delegation Roster and AI Topic Ownership persist across different meetings
        (both within the same DB and across separate meeting folders) as the Default Assignment,
        can be changed between meetings, and that session_appointments act strictly as per-meeting
        session-level overrides.
        """
        dir_177 = self.temp_dir / "SA2_177"
        dir_178 = self.temp_dir / "SA2_178"
        (dir_177 / "Agenda").mkdir(parents=True, exist_ok=True)
        (dir_178 / "Agenda").mkdir(parents=True, exist_ok=True)

        db_177 = SessionPlanDatabase(dir_177 / "Agenda" / "session_plans.db")
        m_martin = DelegationMember(name="Martin", color="#2563EB", is_self=True)
        m_anna = DelegationMember(name="Anna", color="#059669", is_self=False)
        db_177.save_delegation_member("177", m_martin)
        db_177.save_delegation_member("177", m_anna)
        db_177.set_topic_owner("177", "20.6.1", "Martin")
        db_177.set_topic_owner("177", "20.6.2", "Anna")
        db_177.set_session_appointment(
            meeting_number="177",
            slot_key="Monday|Q2|1|2",
            appointed_delegate="Anna",
            original_owner="Martin",
            reason="Covering parallel collision in SA2#177",
        )

        # Open a brand-new meeting folder (SA2 #178) with its own SQLite database
        db_178 = SessionPlanDatabase(dir_178 / "Agenda" / "session_plans.db")
        plan_178 = db_178.get_delegation_plan("178")

        # 1. Roster and Default AI Assignments are automatically inherited!
        self.assertIn("Martin", plan_178.members)
        self.assertIn("Anna", plan_178.members)
        self.assertEqual(plan_178.topic_owners["20.6.1"].owner_name, "Martin")
        self.assertEqual(plan_178.topic_owners["20.6.2"].owner_name, "Anna")

        # 2. Session-level overrides from #177 do NOT leak into #178
        self.assertEqual(len(plan_178.appointments), 0)

        # 3. Changing default assignment in #178 updates the default for future meetings (#179)
        db_178.set_topic_owner("178", "20.6.1", "Anna")
        db_178.set_session_appointment(
            meeting_number="178",
            slot_key="Wednesday|Q3|0|1",
            appointed_delegate="Martin",
            original_owner="Anna",
            reason="Session override in #178",
        )
        plan_178_updated = db_178.get_delegation_plan("178")
        self.assertEqual(plan_178_updated.topic_owners["20.6.1"].owner_name, "Anna")
        self.assertIn("Wednesday|Q3|0|1", plan_178_updated.appointments)

        # 4. Clearing session overrides in #178 leaves persistent Default AI Assignments intact
        cleared = db_178.clear_session_appointments("178")
        self.assertEqual(cleared, 1)
        plan_178_after_clear = db_178.get_delegation_plan("178")
        self.assertEqual(len(plan_178_after_clear.appointments), 0)
        self.assertEqual(plan_178_after_clear.topic_owners["20.6.1"].owner_name, "Anna")

    def test_html_export_unassigned_conflicts_sticky_and_highlight_now(self):
        """
        Verifies the 5 HTML export enhancements:
        1. No 'Starred Topics' button (`btnTracked`)
        2. Unassigned AIs view (`#unassignedView`, `#btnViewUnassigned`, `#btnUnassigned`)
        3. Locked Day header row on vertical scroll (`position: sticky; top: 0; z-index: 20;`)
        4. Conflicts tab (`#conflictsView`, `#btnViewConflicts`, `renderConflictsView`)
        5. Highlight Now dispatch board (`#btnHighlightNow`, `#nowDispatchPanel`, `highlightNow`)
        """
        snap = SessionPlanSnapshot(
            meeting_number="177",
            tdoc_id="S2-2610195",
            snapshot_label="Rev 2",
            slots=[
                SessionSlot(
                    day_index=0,
                    day_name="Monday",
                    quarter="Q1",
                    sub_slot=1,
                    stream_number=1,
                    raw_text="AI 20.6.1 NAS Evolution",
                    items=[AgendaSlotItem(ai_number="20.6.1", display_expr="AI 20.6.1", topic_label="NAS")],
                ),
                SessionSlot(
                    day_index=0,
                    day_name="Monday",
                    quarter="Q1",
                    sub_slot=1,
                    stream_number=2,
                    raw_text="AI 20.6.2 SBA Enhancements",
                    items=[AgendaSlotItem(ai_number="20.6.2", display_expr="AI 20.6.2", topic_label="SBA")],
                ),
                SessionSlot(
                    day_index=1,
                    day_name="Tuesday",
                    quarter="Q2",
                    sub_slot=0,
                    stream_number=3,
                    raw_text="AI 20.6.99 Unassigned Mystery Topic",
                    items=[AgendaSlotItem(ai_number="20.6.99", display_expr="AI 20.6.99", topic_label="Mystery")],
                ),
            ],
        )
        plan = DelegationPlan(
            meeting_number="177",
            members={
                "Martin": DelegationMember(name="Martin", color="#2563EB", is_self=True),
                "Anna": DelegationMember(name="Anna", color="#059669", is_self=False),
            },
            topic_owners={
                "20.6.1": TopicOwnership("177", "20.6.1", "Martin"),
                "20.6.2": TopicOwnership("177", "20.6.2", "Martin"),  # Collides with 20.6.1 on Mon Q1 TU1!
            },
        )
        out_html = self.temp_dir / "enhanced_session_plan.html"
        SessionPlanExporter.export_mobile_html(
            snapshot=snap,
            output_path=out_html,
            meeting_tz_offset_min=120,
            tz_name="Venue Time (Madrid)",
            meeting_dates_str="2026-08-17",
            delegation_plan=plan,
        )
        content = out_html.read_text(encoding="utf-8")

        # 1. Starred topics view removed
        self.assertNotIn('id="btnTracked"', content)
        self.assertNotIn("Starred Topics", content)

        # 2. Unassigned AIs view & quick filter present
        self.assertIn('id="unassignedView"', content)
        self.assertIn('id="btnViewUnassigned"', content)
        self.assertIn('id="btnUnassigned"', content)
        self.assertIn("const UNASSIGNED_AIS =", content)
        self.assertIn("20.6.99", content)

        # 3. Sticky Day row & 2D scroll container
        self.assertIn("max-height: calc(100vh - 175px)", content)
        self.assertIn("position: sticky; top: 0; z-index: 20;", content)
        self.assertIn("position: sticky; top: 0; left: 0; z-index: 30", content)

        # 4. Conflicts view tab
        self.assertIn('id="conflictsView"', content)
        self.assertIn('id="btnViewConflicts"', content)
        self.assertIn("renderConflictsView", content)

        # 5. Highlight Now dispatch guide
        self.assertIn('id="btnHighlightNow"', content)
        self.assertIn('id="nowDispatchPanel"', content)
        self.assertIn("highlightNow", content)
        self.assertIn("Where Each Delegate Should Go", content)

        # Execute new interactive views in Node.js to verify zero runtime exceptions
        script_matches = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
        full_js = "\n".join(script_matches)
        node_bin = shutil.which("node")
        if node_bin:
            harness = f"""
const mockElement = () => ({{
  classList: {{ toggle: () => {{}}, remove: () => {{}}, add: () => {{}} }},
  style: {{ setProperty: () => {{}} }},
  getAttribute: () => '',
  querySelectorAll: () => [],
  querySelector: () => null,
  addEventListener: () => {{}},
  scrollIntoView: () => {{}},
  innerHTML: '',
  textContent: '',
  value: ''
}});
global.document = {{
  documentElement: {{ style: {{ setProperty: () => {{}} }} }},
  getElementById: () => mockElement(),
  querySelectorAll: () => [],
  querySelector: () => mockElement(),
  body: {{ appendChild: () => {{}}, removeChild: () => {{}} }}
}};
global.window = {{ innerWidth: 1280, innerHeight: 900, addEventListener: () => {{}} }};
global.setInterval = () => {{}};
global.setTimeout = () => {{}};

{full_js}

switchView('conflicts');
switchView('unassigned');
toggleUnassigned();
highlightNow(true);
stepNowSlot(1);
resetNowToLiveClock();
highlightNow(false);
"""
            temp_js = self.temp_dir / "verify_enhanced_bundle.js"
            temp_js.write_text(harness, encoding="utf-8")
            res = subprocess.run([node_bin, str(temp_js)], capture_output=True, text=True, encoding="utf-8", errors="replace")
            self.assertEqual(res.returncode, 0, f"Node.js execution failed:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}")

    def test_enhanced_conflicts_free_colleagues_and_quick_appoint(self):
        """
        Verifies desktop Clash & Dispatch Board enhancements:
        - Accurate per-delegate conflict detection in Delegate Agenda (no false clash when different delegates cover parallel streams)
        - Free Colleagues column in `team_clashes_table`
        - 1-click `_quick_appoint_colleague` resolving the collision
        """
        win = SessionPlanWindow(
            meeting_dir=self.temp_dir,
            mtg_info=self.mtg_info,
            agenda_map={},
            tracked_tokens=set(),
        )
        try:
            m1 = DelegationMember(name="Martin", color="#2563EB", is_self=True)
            m2 = DelegationMember(name="Anna", color="#059669", is_self=False)
            win.db.save_delegation_member(win.meeting_number, m1)
            win.db.save_delegation_member(win.meeting_number, m2)
            win.db.set_topic_owner(win.meeting_number, "20.6.1", "Martin")
            win.db.set_topic_owner(win.meeting_number, "20.6.2", "Martin")
            win.delegation_plan = win.db.get_delegation_plan(win.meeting_number)

            s1 = SessionSlot(
                day_index=0, day_name="Monday", quarter="Q1", sub_slot=1, stream_number=1,
                raw_text="AI 20.6.1 NAS", items=[AgendaSlotItem(ai_number="20.6.1", display_expr="AI 20.6.1", topic_label="NAS")],
            )
            s2 = SessionSlot(
                day_index=0, day_name="Monday", quarter="Q1", sub_slot=1, stream_number=2,
                raw_text="AI 20.6.2 SBA", items=[AgendaSlotItem(ai_number="20.6.2", display_expr="AI 20.6.2", topic_label="SBA")],
            )
            win.active_snapshot = SessionPlanSnapshot(
                meeting_number=win.meeting_number, tdoc_id="S2-2610195", slots=[s1, s2]
            )
            win._render_current_views()

            # 1. Clash detected for Martin; Anna is shown as Free Colleague in column 5
            self.assertEqual(win.team_clashes_table.rowCount(), 1)
            free_item = win.team_clashes_table.item(0, 5)
            self.assertIsNotNone(free_item)
            self.assertIn("Anna", free_item.text())
            self.assertEqual(win.btn_conflicts_quick.text(), "⚠️ Conflicts (1)")

            # 2. Quick-appoint Anna to cover Stream 2 -> resolves collision!
            win._quick_appoint_colleague(s2, "Martin", "Anna")
            self.assertEqual(win.team_clashes_table.rowCount(), 0)
            self.assertEqual(win.btn_conflicts_quick.text(), "⚠️ Conflicts (0)")

            # 3. Delegate Agenda in 'Whole Delegation (ALL)' scope must NOT report a false conflict
            # now that Martin is in Stream 1 and Anna is in Stream 2
            self.assertNotIn("Clash", win.btn_tracked.text())
        finally:
            win.close()


if __name__ == "__main__":
    unittest.main()
