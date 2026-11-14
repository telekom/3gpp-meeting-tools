# -*- coding: utf-8 -*-
"""Unit tests for SA2 session plan parsing robustness and edge cases."""

import unittest
import tempfile
import os
import struct
from pathlib import Path
from unittest.mock import MagicMock

from PyQt5.QtWidgets import QApplication, QWidget, QTableWidgetItem
from PyQt5.QtCore import QMimeData, QUrl, QByteArray, QPoint, Qt, QSettings
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

# Ensure QApplication exists for Qt tests
app = QApplication.instance()
if not app:
    app = QApplication([])

from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
from modules.meetings.core.session_plan.session_plan_models import SessionPlanSnapshot, SessionSlot
from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow
from modules.meetings.ui.tdocs_drop_overlay import DropOverlayWidget


class TestSessionPlanRobustness(unittest.TestCase):
    """Verifies parsing robustness against flexible SA2 session plan conventions."""

    def test_extract_room_override(self):
        """Tests physical room overrides detection and text cleanup."""
        text = "In Breakout 2\nKI#8 (Net Sharing) (20.6.8)"
        room, tokens = SessionPlanParser._extract_room_override_and_tokens(text)
        self.assertEqual(room, "Breakout 2")
        self.assertEqual(tokens, ["KI#8 (Net Sharing) (20.6.8)"])

        text2 = "In main room \nKI#19 (Network for AI) (20.6.19)"
        room2, tokens2 = SessionPlanParser._extract_room_override_and_tokens(text2)
        self.assertEqual(room2, "Main Room")
        self.assertEqual(tokens2, ["KI#19 (Network for AI) (20.6.19)"])

        text3 = "In breakout 1\nFS_enIMS_ARC (20.10.1)"
        room3, tokens3 = SessionPlanParser._extract_room_override_and_tokens(text3)
        self.assertEqual(room3, "Breakout 1")
        self.assertEqual(tokens3, ["FS_enIMS_ARC (20.10.1)"])

        text_none = "KI#5 (QoS) (20.6.5)"
        room_none, tokens_none = SessionPlanParser._extract_room_override_and_tokens(text_none)
        self.assertEqual(room_none, "")
        self.assertEqual(tokens_none, [text_none])

    def test_breakout_not_detected_as_break_period(self):
        """Ensures 'Breakout 1/2' in room headers or cells is not mistaken for a coffee break."""
        p1 = SessionPlanParser._detect_break_period("Breakout 1", "Main Room", "Palomino", "Q1")
        self.assertIsNone(p1)

        p2 = SessionPlanParser._detect_break_period("Coffee Break", "", "", "Q1")
        self.assertEqual(p2, "Coffee 1")

        p3 = SessionPlanParser._detect_break_period("Lunch Break 12:30 - 14:00", "", "", "Q2")
        self.assertEqual(p3, "Lunch")

    def test_extract_multi_line_agenda_items(self):
        """Tests that multiple distinct agenda lines in one slot map to their respective topics."""
        raw_text = (
            "In Breakout 2\n"
            "KI#10 (FWA) (20.6.10)\n"
            "KI#12 – KI#16"
        )
        room_override, items = SessionPlanParser._extract_agenda_items(raw_text)
        self.assertEqual(room_override, "Breakout 2")
        item_map = {it.ai_number: it.topic_label for it in items}

        self.assertIn("20.6.10", item_map)
        self.assertEqual(item_map["20.6.10"], "KI#10 (FWA)")

        for ai in ["20.6.12", "20.6.13", "20.6.14", "20.6.15", "20.6.16"]:
            self.assertIn(ai, item_map)
            self.assertEqual(item_map[ai], "KI#12 – KI#16")

    def test_single_digit_agenda_items(self):
        """Tests parenthesized single-digit AIs like Opening (1), Agenda (2), Reports (3)."""
        raw_text = "Opening (1), Agenda (2), Reports (3), Work plan (4), Outgoing LSs (5.1)"
        _, items = SessionPlanParser._extract_agenda_items(raw_text)
        item_map = {it.ai_number: it.topic_label for it in items}

        self.assertEqual(item_map.get("1"), "Opening")
        self.assertEqual(item_map.get("2"), "Agenda")
        self.assertEqual(item_map.get("3"), "Reports")
        self.assertEqual(item_map.get("4"), "Work plan")
        self.assertEqual(item_map.get("5.1"), "Outgoing LSs")

    def test_implicit_ki_number_mapping(self):
        """Tests resolving implicit KI#N (e.g. KI#21 without repeating (20.6.21)) using shared ki_map."""
        ki_map = {"21": "20.6.21"}
        raw_text = "KI#21 (Data Framework)"
        _, items = SessionPlanParser._extract_agenda_items(raw_text, ki_map=ki_map)
        self.assertTrue(len(items) >= 1)
        self.assertEqual(items[0].ai_number, "20.6.21")
        self.assertEqual(items[0].topic_label, "KI#21 (Data Framework)")

    def test_balance_parentheses(self):
        """Tests that parentheses in topic labels are not erroneously stripped or left dangling."""
        self.assertEqual(SessionPlanParser._balance_parentheses("(Net Sharing)"), "(Net Sharing)")
        self.assertEqual(SessionPlanParser._balance_parentheses("Net Sharing)"), "Net Sharing")
        self.assertEqual(SessionPlanParser._balance_parentheses("(Net Sharing"), "Net Sharing")

    def test_db_room_override_persistence(self):
        """Tests that room_override is saved and restored from SQLite database."""
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tf:
            temp_db = tf.name

        try:
            db = SessionPlanDatabase(temp_db)
            slot = SessionSlot(
                day_index=0,
                day_name="Monday",
                day_date="2026-11-09",
                quarter="Q1",
                stream_number=1,
                room_override="Breakout 1",
                raw_text="In breakout 1\nFS_enIMS_ARC (20.10.1)",
            )
            snap = SessionPlanSnapshot(
                meeting_number="178",
                slots=[slot],
            )
            db.save_snapshot(snap)

            loaded = db.get_active_snapshot("178")
            self.assertIsNotNone(loaded)
            self.assertEqual(len(loaded.slots), 1)
            self.assertEqual(loaded.slots[0].room_override, "Breakout 1")
        finally:
            if os.path.exists(temp_db):
                os.remove(temp_db)

    def test_drag_drop_validation_local_files_and_text(self):
        """Tests that local files, Outlook attachment formats, and table texts are validated."""
        with tempfile.TemporaryDirectory() as td:
            win = SessionPlanWindow(
                meeting_dir=Path(td),
                mtg_info={"meeting_number": "178", "location": "Calgary"},
            )
            try:
                # 1. Valid local file
                mime_file = QMimeData()
                mime_file.setUrls([QUrl.fromLocalFile(os.path.join(td, "plan.docx"))])
                self.assertTrue(win._is_valid_drop_data(mime_file))

                # 2. Invalid local file (.exe)
                mime_bad = QMimeData()
                mime_bad.setUrls([QUrl.fromLocalFile(os.path.join(td, "malware.exe"))])
                self.assertFalse(win._is_valid_drop_data(mime_bad))

                # 3. Virtual Outlook attachment (FileGroupDescriptorW)
                mime_outlook = QMimeData()
                mime_outlook.setData("application/x-qt-windows-mime;value=\"FileGroupDescriptorW\"", QByteArray(b"\x01\x00\x00\x00"))
                self.assertTrue(win._is_valid_drop_data(mime_outlook))

                # 4. Dragged email text table
                mime_text = QMimeData()
                mime_text.setText("| Stream 1 | Main Room | Monday Q1 | Opening |")
                self.assertTrue(win._is_valid_drop_data(mime_text))

                # 5. Irrelevant text
                mime_random = QMimeData()
                mime_random.setText("Hello world, this is a plain message.")
                self.assertFalse(win._is_valid_drop_data(mime_random))
            finally:
                win.close()

    def test_extract_outlook_virtual_attachment(self):
        """Tests parsing FileGroupDescriptorW and FileContents to extract a dropped attachment."""
        with tempfile.TemporaryDirectory() as td:
            win = SessionPlanWindow(
                meeting_dir=Path(td),
                mtg_info={"meeting_number": "178", "location": "Calgary"},
            )
            try:
                mime = QMimeData()
                # Synthesize FILEGROUPDESCRIPTORW with 1 file: "S2-2610195_Plan.docx"
                # Header: 4 bytes (num_files = 1)
                header = struct.pack("<I", 1)
                # FILEDESCRIPTORW record: 72 bytes pre-name, 520 bytes name
                filename = "S2-2610195_Plan.docx\0"
                encoded_name = filename.encode("utf-16le")
                encoded_name += b"\x00" * (520 - len(encoded_name))
                record = b"\x00" * 72 + encoded_name
                descriptor_bytes = header + record

                content = b"PK\x03\x04MockDocxZipContent"
                mime.setData("application/x-qt-windows-mime;value=\"FileGroupDescriptorW\"", QByteArray(descriptor_bytes))
                mime.setData("application/x-qt-windows-mime;value=\"FileContents\"", QByteArray(content))

                extracted = win._extract_outlook_attachments(mime)
                self.assertEqual(len(extracted), 1)
                self.assertEqual(extracted[0].name, "S2-2610195_Plan.docx")
                self.assertEqual(extracted[0].read_bytes(), content)

                payloads = win._extract_drop_payloads(mime)
                self.assertEqual(len(payloads), 1)
                self.assertEqual(payloads[0][0], "file")
                self.assertTrue(payloads[0][1].endswith("S2-2610195_Plan.docx"))
            finally:
                win.close()

    def test_drop_overlay_event_forwarding(self):
        """Tests that DropOverlayWidget forwards drag and drop events directly to its parent."""
        parent_widget = QWidget()
        parent_widget.dragEnterEvent = MagicMock()
        parent_widget.dragMoveEvent = MagicMock()
        parent_widget.dropEvent = MagicMock()

        overlay = DropOverlayWidget(parent_widget)
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile("test.docx")])

        enter_ev = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        overlay.dragEnterEvent(enter_ev)
        parent_widget.dragEnterEvent.assert_called_once_with(enter_ev)

        drop_ev = QDropEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        overlay.dropEvent(drop_ev)
        parent_widget.dropEvent.assert_called_once_with(drop_ev)

    def test_scheduled_only_toggle_and_table_readability(self):
        """Verifies the Scheduled Only toggle button, column sizing, and row filtering in Team Allocations."""
        QSettings("3GPP_Tools", "SessionPlanWindow").remove("team_scheduled_only")
        with tempfile.TemporaryDirectory() as td:
            win = SessionPlanWindow(
                meeting_dir=Path(td),
                mtg_info={"meeting_number": "178", "location": "Calgary"},
            )
            try:
                # 1. UI element and column header verification
                self.assertTrue(hasattr(win, "btn_toggle_scheduled_only"))
                self.assertTrue(win.btn_toggle_scheduled_only.isCheckable())
                self.assertFalse(win.btn_toggle_scheduled_only.isChecked())

                headers = [win.team_topics_table.horizontalHeaderItem(c).text() for c in range(win.team_topics_table.columnCount())]
                expected_headers = ["AI #", "Topic Description", "📅 Scheduled Slots", "👤 Topic Owner", "🛡️ Coverage Status"]
                self.assertEqual(headers, expected_headers)
                self.assertEqual(win.team_topics_table.columnWidth(0), 65)
                self.assertEqual(win.team_topics_table.columnWidth(2), 220)
                self.assertEqual(win.team_topics_table.columnWidth(3), 195)
                self.assertEqual(win.team_topics_table.columnWidth(4), 185)

                # 2. Populate synthetic rows in team_topics_table
                win.team_topics_table.setRowCount(3)

                # Row 0: Unscheduled AI
                item_ai0 = QTableWidgetItem("1.1")
                item_desc0 = QTableWidgetItem("Opening of meeting")
                item_slots0 = QTableWidgetItem("— (Not scheduled)")
                item_slots0.setData(Qt.UserRole, 0)
                win.team_topics_table.setItem(0, 0, item_ai0)
                win.team_topics_table.setItem(0, 1, item_desc0)
                win.team_topics_table.setItem(0, 2, item_slots0)

                # Row 1: Scheduled AI
                item_ai1 = QTableWidgetItem("20.6.5")
                item_desc1 = QTableWidgetItem("KI#5 (QoS)")
                item_slots1 = QTableWidgetItem("Mon Q5 (Str 1)")
                item_slots1.setData(Qt.UserRole, 1)
                win.team_topics_table.setItem(1, 0, item_ai1)
                win.team_topics_table.setItem(1, 1, item_desc1)
                win.team_topics_table.setItem(1, 2, item_slots1)

                # Row 2: Scheduled AI
                item_ai2 = QTableWidgetItem("20.6.6")
                item_desc2 = QTableWidgetItem("KI#6 (Policy)")
                item_slots2 = QTableWidgetItem("Tue Q3 (Str 1)")
                item_slots2.setData(Qt.UserRole, 1)
                win.team_topics_table.setItem(2, 0, item_ai2)
                win.team_topics_table.setItem(2, 1, item_desc2)
                win.team_topics_table.setItem(2, 2, item_slots2)

                # Initial state: unchecked -> all rows visible
                win._filter_topic_ownership_table()
                self.assertFalse(win.team_topics_table.isRowHidden(0))
                self.assertFalse(win.team_topics_table.isRowHidden(1))
                self.assertFalse(win.team_topics_table.isRowHidden(2))

                # 3. Toggle ON Scheduled Only
                win.btn_toggle_scheduled_only.setChecked(True)
                self.assertTrue(win.team_topics_table.isRowHidden(0))
                self.assertFalse(win.team_topics_table.isRowHidden(1))
                self.assertFalse(win.team_topics_table.isRowHidden(2))
                self.assertIn("✓ 📅 Scheduled Only (2/3)", win.btn_toggle_scheduled_only.text())

                # 4. Search filter interaction while Scheduled Only is active
                win.team_ai_search.setText("QoS")
                self.assertTrue(win.team_topics_table.isRowHidden(0))
                self.assertFalse(win.team_topics_table.isRowHidden(1))
                self.assertTrue(win.team_topics_table.isRowHidden(2))

                win.team_ai_search.clear()
                self.assertTrue(win.team_topics_table.isRowHidden(0))
                self.assertFalse(win.team_topics_table.isRowHidden(1))
                self.assertFalse(win.team_topics_table.isRowHidden(2))

                # 5. Toggle OFF Scheduled Only -> all visible again
                win.btn_toggle_scheduled_only.setChecked(False)
                self.assertFalse(win.team_topics_table.isRowHidden(0))
                self.assertFalse(win.team_topics_table.isRowHidden(1))
                self.assertFalse(win.team_topics_table.isRowHidden(2))
                self.assertIn("📅 Scheduled Only (2/3)", win.btn_toggle_scheduled_only.text())

                # 6. Verify QSettings persistence
                settings = QSettings("3GPP_Tools", "SessionPlanWindow")
                win.btn_toggle_scheduled_only.setChecked(True)
                self.assertTrue(settings.value("team_scheduled_only", False, type=bool))
                win.btn_toggle_scheduled_only.setChecked(False)
                self.assertFalse(settings.value("team_scheduled_only", False, type=bool))
            finally:
                win.close()

    def test_docx_track_changes_and_dynamic_quarter_times(self):
        """
        Tests that <w:ins> insertions are accepted, <w:del> deletions are ignored,
        vMerge into Break rows is deduplicated, and dynamic quarter time ranges (e.g. 1230 - 1400)
        compute accurate half-quarter midpoints (12:30-13:15 and 13:15-14:00).
        """
        import docx
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        with tempfile.TemporaryDirectory() as td:
            docx_path = Path(td) / "S2-261nnnn_r1.docx"
            doc = docx.Document()
            doc.add_paragraph("SA2 #178 meeting session plan")
            t = doc.add_table(rows=5, cols=13)

            # Row 0: Header
            t.rows[0].cells[3].text = "Monday"
            t.rows[0].cells[5].text = "Tuesday"

            # Row 1: Q0 Stream 1 with vertical merge across Row 2 (Break) on Monday
            t.rows[1].cells[0].text = "Q0"
            t.rows[1].cells[1].text = "0800 - 0850"
            t.rows[1].cells[2].text = "Stream 1"
            t.rows[1].cells[3].text = "Opening of meeting at 0900"
            t.rows[1].cells[3].merge(t.rows[1].cells[4])

            # Row 2: Break row where Monday cells are vertically merged with Row 1
            t.rows[2].cells[2].text = "Break"
            t.rows[1].cells[3].merge(t.rows[2].cells[4])

            # Row 3: Q2 Stream 1 with Track Changes on time (1100-1230 -> 1230-1400)
            t.rows[3].cells[0].text = "Q2"
            tc_time = t.rows[3].cells[1]._tc
            tc_time.clear_content()
            tc_time.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}>"
                    f"<w:del><w:r><w:delText>1100 </w:delText></w:r></w:del>"
                    f"<w:ins><w:r><w:t>1230</w:t></w:r></w:ins>"
                    f"<w:r><w:t> - </w:t></w:r>"
                    f"<w:del><w:r><w:delText>1230</w:delText></w:r></w:del>"
                    f"<w:ins><w:r><w:t>1400</w:t></w:r></w:ins>"
                    f"</w:p>"
                )
            )
            t.rows[3].cells[2].text = "Stream 1"
            t.rows[3].cells[3].text = "KI#1.1 (NAS) (20.6.1)"

            # Row 4: Q2 Stream 2 with Track Changes swap between Monday and Tuesday
            t.rows[4].cells[2].text = "Stream 2"
            # Mon TU1 (col 3): inserted KI#10 + KI#12-16, deleted Pre-Rel-19
            tc_mon_tu1 = t.rows[4].cells[3]._tc
            tc_mon_tu1.clear_content()
            tc_mon_tu1.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:ins><w:r><w:t>KI#10 (FWA) (20.6.10)</w:t></w:r></w:ins></w:p>"
                )
            )
            tc_mon_tu1.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:ins><w:r><w:t>KI#12 – KI#16 (20.6.12 - 16)</w:t></w:r></w:ins></w:p>"
                )
            )
            tc_mon_tu1.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:del><w:r><w:delText>Pre-Rel-19 maintenance (5.x, 6.x, 7.x)</w:delText></w:r></w:del></w:p>"
                )
            )

            # Mon TU2 (col 4): inserted SMS2EC-ARC (20.8.2)
            tc_mon_tu2 = t.rows[4].cells[4]._tc
            tc_mon_tu2.clear_content()
            tc_mon_tu2.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:ins><w:r><w:t>SMS2EC-ARC (20.8.2)</w:t></w:r></w:ins></w:p>"
                )
            )

            # Tue TU1 (col 5): inserted Pre-Rel-19 maintenance, deleted KI#10
            tc_tue_tu1 = t.rows[4].cells[5]._tc
            tc_tue_tu1.clear_content()
            tc_tue_tu1.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:ins><w:r><w:t>Pre-Rel-19 maintenance (5.x, 6.x, 7.x)</w:t></w:r></w:ins></w:p>"
                )
            )
            tc_tue_tu1.append(
                parse_xml(
                    f"<w:p {nsdecls('w')}><w:del><w:r><w:delText>KI#10 (FWA) (20.6.10)</w:delText></w:r></w:del></w:p>"
                )
            )

            # Tue TU2 (col 6): normal text 5GSAT_Ph4-ARC (20.1.2)
            t.rows[4].cells[6].text = "5GSAT_Ph4-ARC (20.1.2)"

            doc.save(str(docx_path))

            snap = SessionPlanParser.parse_docx(docx_path, meeting_number="178")

            # Verify Break 1 did not inherit 'Opening of meeting at 0900' from Q0 vertical merge
            break_slots = [s for s in snap.slots if "break" in s.quarter.lower()]
            self.assertEqual(len(break_slots), 0)

            # Verify Q2 Stream 2 slots on Monday and Tuesday
            q2_s2 = {(s.day_name, s.sub_slot): s for s in snap.slots if s.quarter == "Q2" and s.stream_number == 2}
            self.assertIn(("Monday", 1), q2_s2)
            self.assertIn(("Monday", 2), q2_s2)
            self.assertIn(("Tuesday", 1), q2_s2)
            self.assertIn(("Tuesday", 2), q2_s2)

            mon_tu1 = q2_s2[("Monday", 1)]
            self.assertIn("KI#10 (FWA)", mon_tu1.raw_text)
            self.assertNotIn("Pre-Rel-19", mon_tu1.raw_text)
            self.assertEqual((mon_tu1.start_time_local, mon_tu1.end_time_local), ("12:30", "13:15"))

            mon_tu2 = q2_s2[("Monday", 2)]
            self.assertEqual(mon_tu2.raw_text, "SMS2EC-ARC (20.8.2)")
            self.assertEqual((mon_tu2.start_time_local, mon_tu2.end_time_local), ("13:15", "14:00"))

            tue_tu1 = q2_s2[("Tuesday", 1)]
            self.assertIn("Pre-Rel-19 maintenance", tue_tu1.raw_text)
            self.assertNotIn("KI#10", tue_tu1.raw_text)
            self.assertEqual((tue_tu1.start_time_local, tue_tu1.end_time_local), ("12:30", "13:15"))

            tue_tu2 = q2_s2[("Tuesday", 2)]
            self.assertEqual(tue_tu2.raw_text, "5GSAT_Ph4-ARC (20.1.2)")
            self.assertEqual((tue_tu2.start_time_local, tue_tu2.end_time_local), ("13:15", "14:00"))

    def test_overview_grid_clears_stale_cell_widgets_on_snapshot_switch(self):
        """
        Verifies that switching from a snapshot with a populated slot to a snapshot where
        that slot is empty properly clears the old cellWidget from QTableWidget.
        """
        from modules.meetings.core.session_plan.session_plan_models import SessionPlanSnapshot, SessionSlot

        with tempfile.TemporaryDirectory() as td:
            win = SessionPlanWindow(
                meeting_dir=Path(td),
                mtg_info={"meeting_number": "178", "location": "Calgary"},
            )
            try:
                snap1 = SessionPlanSnapshot(
                    meeting_number="178",
                    tdoc_id="S2-2610001",
                    snapshot_label="v1",
                    slots=[
                        SessionSlot(
                            day_index=0,
                            day_name="Monday",
                            quarter="Q2",
                            sub_slot=1,
                            start_time_local="12:30",
                            end_time_local="13:15",
                            stream_number=2,
                            raw_text="Pre-Rel-19 maintenance (5.x, 6.x, 7.x)",
                        )
                    ],
                )
                win.active_snapshot = snap1
                win._render_overview_grid()

                q2_s2_row = next(
                    idx for idx, r in enumerate(win.grid_rows)
                    if r.period == "Q2" and r.stream_number == 2
                )
                self.assertIsNotNone(win.grid_table.cellWidget(q2_s2_row, 0))

                # Switch to snap2 where Monday Q2 Stream 2 is empty
                snap2 = SessionPlanSnapshot(
                    meeting_number="178",
                    tdoc_id="S2-2610002",
                    snapshot_label="v2",
                    slots=[],
                )
                win.active_snapshot = snap2
                win._render_overview_grid()

                self.assertIsNone(win.grid_table.cellWidget(q2_s2_row, 0))
            finally:
                win.close()

    def test_wildcard_dot_x_ai_parsing_and_delegation_matching(self):
        """
        Verifies '.x' / '.X' AI parsing (parenthesized, standalone, multi-item) and
        wildcard resolution against concrete agenda items & delegation ownership.
        """
        from types import SimpleNamespace
        from modules.meetings.core.session_plan.session_plan_models import (
            DelegationMember,
            DelegationPlan,
            SessionPlanSnapshot,
            SessionSlot,
            TopicOwnership,
            ai_number_in_text,
            ai_numbers_match,
            format_slot_display_text,
            is_slot_not_followed,
            resolve_slot_effective_delegates,
        )
        from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter

        # 1. Parser variants
        _, items_wp = SessionPlanParser._extract_agenda_items(
            "Work planning (30.x), etc\n(Rel-21 5G-A revisions)"
        )
        self.assertEqual([it.ai_number for it in items_wp], ["30.x"])
        self.assertIn("Work planning", items_wp[0].topic_label)

        _, items_standalone = SessionPlanParser._extract_agenda_items("Work planning 30.x")
        self.assertEqual([it.ai_number for it in items_standalone], ["30.x"])

        _, items_upper = SessionPlanParser._extract_agenda_items("Rel-19 revisions (20.6.X)")
        self.assertEqual([it.ai_number for it in items_upper], ["20.6.x"])

        _, items_multi = SessionPlanParser._extract_agenda_items("Opening (1), Work planning (30.x)")
        self.assertEqual([it.ai_number for it in items_multi], ["1", "30.x"])

        # 2. Wildcard match helpers
        self.assertTrue(ai_numbers_match("30.x", "30.1"))
        self.assertTrue(ai_numbers_match("30.x", "30"))
        self.assertTrue(ai_numbers_match("20.6.x", "20.6.18"))
        self.assertFalse(ai_numbers_match("30.x", "300.1"))
        self.assertFalse(ai_numbers_match("3.x", "30.1"))
        self.assertTrue(ai_number_in_text("30.1", "Work planning (30.x)"))
        self.assertFalse(ai_number_in_text("20.6.2", "KI#24 (20.6.24)"))

        # 3. Delegation & Not-Followed resolution on a (30.x) slot
        slot_30x = SessionSlot(
            day_index=4,
            day_name="Friday",
            quarter="Q3",
            sub_slot=0,
            stream_number=1,
            room_override="Main Room",
            raw_text="In main room\nWork planning (30.x), etc\n[12]",
            items=items_wp,
        )
        plan_assigned = DelegationPlan(
            meeting_number="178",
            members={"Josep": DelegationMember(name="Josep", color="#2563EB")},
            topic_owners={"30.1": TopicOwnership(meeting_number="178", ai_number="30.1", owner_name="Josep")},
        )
        dels, _, is_all = resolve_slot_effective_delegates(slot_30x, plan_assigned)
        self.assertEqual(dels, ["Josep"])
        self.assertFalse(is_all)
        self.assertFalse(is_slot_not_followed(slot_30x, plan_assigned))

        plan_nf = DelegationPlan(
            meeting_number="178",
            members={"Josep": DelegationMember(name="Josep", color="#2563EB")},
            topic_owners={"30.1": TopicOwnership(meeting_number="178", ai_number="30.1", owner_name="NOT_FOLLOWED")},
        )
        self.assertTrue(is_slot_not_followed(slot_30x, plan_nf))

        # 4. Shared card display text formatting
        clean_txt, empty_draft = format_slot_display_text(
            "In breakout 1\nFS_enIMS_ARC (20.10.1)\n[23]",
            room_override="Breakout 1",
        )
        self.assertEqual(clean_txt, "FS_enIMS_ARC (20.10.1) [23]")
        self.assertFalse(empty_draft)

        _, is_empty = format_slot_display_text("Drafting:")
        self.assertTrue(is_empty)

        # 5. Desktop unified card & HTML export float-right / decluttered output
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            agenda_map = {
                "30.1": SimpleNamespace(ai_number="30.1", description="Work planning Rel-21", acronym="WP")
            }
            win = SessionPlanWindow(
                meeting_dir=td_path,
                mtg_info={"meeting_number": "178", "location": "Calgary"},
                agenda_map=agenda_map,
            )
            try:
                win.db.save_delegation_member("178", DelegationMember(name="Josep", color="#2563EB"))
                win.db.set_topic_owner("178", "30.1", "Josep")
                win.delegation_plan = win.db.get_delegation_plan("178")
                slot_draft_empty = SessionSlot(
                    day_index=1, day_name="Tuesday", quarter="Q0", sub_slot=0, stream_number=2,
                    raw_text="Drafting:", is_drafting=True,
                )
                win.active_snapshot = SessionPlanSnapshot(
                    meeting_number="178",
                    tdoc_id="S2-2611637",
                    slots=[slot_30x, slot_draft_empty],
                )
                win._render_current_views()

                # Agenda Mode should mark 30.1 as Scheduled via (30.x)
                self.assertIn("1 Sched", win.btn_agenda.text())

                # Grid cell uses unified #slotCard frame with 📍 badge
                q3_s1_row = next(
                    idx for idx, r in enumerate(win.grid_rows)
                    if r.period == "Q3" and r.stream_number == 1
                )
                cell_w = win.grid_table.cellWidget(q3_s1_row, 8)
                self.assertIsNotNone(cell_w)
                self.assertIsNotNone(cell_w.findChild(type(win.room_strip), "slotCard"))

                out_html = td_path / "compact_export.html"
                SessionPlanExporter.export_mobile_html(
                    snapshot=win.active_snapshot,
                    output_path=out_html,
                    meeting_tz_offset_min=-420,
                    delegation_plan=win.delegation_plan,
                    agenda_map=agenda_map,
                )
                html_txt = out_html.read_text(encoding="utf-8")
                self.assertIn('class="cell-del-float"', html_txt)
                self.assertIn('class="cell-room-override">📍 Main Room</span>', html_txt)
                self.assertIn('is-empty-drafting', html_txt)
                self.assertNotIn('"ai_number": "30.1"', html_txt)
                self.assertNotIn('"ai_number": "30.x"', html_txt)
            finally:
                win.close()

    def test_wildcard_x_expansion_against_agenda_map(self):
        """
        Verifies that .x wildcards (e.g. 30.x) expand into concrete agenda_map child AIs
        (30.1, 30.2, 30.3) for tooltips, tracking, partial unassigned detection,
        Not-Followed semantics, and HTML export (including appointment override safety).
        """
        from types import SimpleNamespace
        from modules.meetings.core.session_plan.session_plan_models import (
            DelegationMember,
            DelegationPlan,
            SessionAppointment,
            SessionPlanSnapshot,
            SessionSlot,
            TopicOwnership,
            build_slot_ai_bullets_html,
            expand_ai_number,
            get_slot_unassigned_ais,
            is_slot_not_followed,
            resolve_slot_effective_delegates,
        )
        from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
        from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
        from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow

        agenda_map = {
            "30": SimpleNamespace(ai_number="30", description="Work plan, future meetings, etc."),
            "30.1": SimpleNamespace(ai_number="30.1", description="Work plan and schedule"),
            "30.2": SimpleNamespace(ai_number="30.2", description="Rel-21 5G-A revisions"),
            "30.3": SimpleNamespace(ai_number="30.3", description="Future meetings"),
        }

        # 1. Wildcard expansion prefers strict children over section header "30"
        self.assertEqual(expand_ai_number("30.x", agenda_map), ["30.1", "30.2", "30.3"])
        self.assertEqual(expand_ai_number("30.x", {"30": agenda_map["30"]}), ["30"])
        self.assertEqual(expand_ai_number("30.x", {}), ["30.x"])

        _, items_30x = SessionPlanParser._extract_agenda_items("Work planning (30.x), etc (Rel-21 5G-A revisions)")
        slot_30x = SessionSlot(
            day_index=2, day_name="Wednesday", quarter="Q4", sub_slot=1, stream_number=1,
            room_override="Main Room",
            raw_text="Work planning (30.x), etc (Rel-21 5G-A revisions)",
            items=items_30x,
        )

        # 2. Partial assignment: 30.1 -> Josep, 30.2 -> NOT_FOLLOWED, 30.3 -> Unassigned
        plan_partial = DelegationPlan(
            meeting_number="178",
            members={
                "Josep": DelegationMember(name="Josep", color="#2563EB"),
                "Dieter": DelegationMember(name="Dieter", color="#059669"),
            },
            topic_owners={
                "30.1": TopicOwnership(meeting_number="178", ai_number="30.1", owner_name="Josep"),
                "30.2": TopicOwnership(meeting_number="178", ai_number="30.2", owner_name="NOT_FOLLOWED"),
            },
        )
        dels, _, is_all = resolve_slot_effective_delegates(slot_30x, plan_partial, agenda_map=agenda_map)
        self.assertEqual(dels, ["Josep"])
        self.assertFalse(is_all)
        self.assertEqual(get_slot_unassigned_ais(slot_30x, plan_partial, agenda_map=agenda_map), ["30.3"])
        self.assertFalse(is_slot_not_followed(slot_30x, plan_partial, agenda_map=agenda_map))

        # Only Not Followed when EVERY expanded sub-AI is Not Followed
        plan_partial_nf = DelegationPlan(
            meeting_number="178",
            members={"Josep": DelegationMember(name="Josep", color="#2563EB")},
            topic_owners={
                "30.1": TopicOwnership(meeting_number="178", ai_number="30.1", owner_name="NOT_FOLLOWED"),
            },
        )
        self.assertFalse(is_slot_not_followed(slot_30x, plan_partial_nf, agenda_map=agenda_map))

        plan_all_nf = DelegationPlan(
            meeting_number="178",
            members={"Josep": DelegationMember(name="Josep", color="#2563EB")},
            topic_owners={
                "30.1": TopicOwnership(meeting_number="178", ai_number="30.1", owner_name="NOT_FOLLOWED"),
                "30.2": TopicOwnership(meeting_number="178", ai_number="30.2", owner_name="NOT_FOLLOWED"),
                "30.3": TopicOwnership(meeting_number="178", ai_number="30.3", owner_name="NOT_FOLLOWED"),
            },
        )
        self.assertTrue(is_slot_not_followed(slot_30x, plan_all_nf, agenda_map=agenda_map))

        # 3. Tooltip bullets show expanded sub-AIs and their per-AI statuses
        bullets_html = "<br>".join(build_slot_ai_bullets_html(slot_30x, agenda_map=agenda_map, plan=plan_partial))
        self.assertIn("AI 30.x:", bullets_html)
        self.assertIn("covers 3 agenda item(s)", bullets_html)
        self.assertIn("30.1:", bullets_html)
        self.assertIn("Work plan and schedule", bullets_html)
        self.assertIn("Josep", bullets_html)
        self.assertIn("30.2:", bullets_html)
        self.assertIn("Not Followed", bullets_html)
        self.assertIn("30.3:", bullets_html)
        self.assertIn("Unassigned", bullets_html)

        # 4. Desktop UI tracking + HTML export + appointment override safety
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            win = SessionPlanWindow(
                meeting_dir=td_path,
                mtg_info={"meeting_number": "178", "location": "Calgary"},
                agenda_map=agenda_map,
                tracked_tokens={"30.3"},
            )
            try:
                win.db.save_delegation_member("178", DelegationMember(name="Josep", color="#2563EB"))
                win.db.save_delegation_member("178", DelegationMember(name="Dieter", color="#059669"))
                win.db.set_topic_owner("178", "30.1", "Josep")
                win.db.set_topic_owner("178", "30.2", "NOT_FOLLOWED")
                win.delegation_plan = win.db.get_delegation_plan("178")
                win.active_snapshot = SessionPlanSnapshot(
                    meeting_number="178", tdoc_id="S2-2611637", slots=[slot_30x]
                )
                win._render_current_views()

                # Tracking concrete 30.3 must match the (30.x) slot
                self.assertTrue(win._is_slot_tracked(slot_30x))
                tooltip_html = win._build_rich_slot_tooltip(slot_30x, 2, "Q4", 1, 1)
                self.assertIn("30.1:", tooltip_html)
                self.assertIn("Work plan and schedule", tooltip_html)
                self.assertIn("30.3:", tooltip_html)

                out_html = td_path / "expanded_30x.html"
                SessionPlanExporter.export_mobile_html(
                    snapshot=win.active_snapshot,
                    output_path=out_html,
                    meeting_tz_offset_min=-420,
                    delegation_plan=win.delegation_plan,
                    agenda_map=agenda_map,
                )
                html_txt = out_html.read_text(encoding="utf-8")
                self.assertIn("Work plan and schedule", html_txt)
                self.assertIn("❓ AI 30.3", html_txt)
                self.assertNotIn("❓ AI 30.1", html_txt)

                # Appointment override on the slot must work cleanly in get_slot_unassigned_ais & HTML export
                win.db.set_session_appointment("178", slot_30x.slot_key, "Dieter", "Josep", "Cover 30.x")
                win.delegation_plan = win.db.get_delegation_plan("178")
                self.assertEqual(get_slot_unassigned_ais(slot_30x, win.delegation_plan, agenda_map=agenda_map), [])
                SessionPlanExporter.export_mobile_html(
                    snapshot=win.active_snapshot,
                    output_path=out_html,
                    meeting_tz_offset_min=-420,
                    delegation_plan=win.delegation_plan,
                    agenda_map=agenda_map,
                )
            finally:
                win.close()


if __name__ == "__main__":
    unittest.main()
