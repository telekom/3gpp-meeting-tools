"""
Regression tests for SessionPlanWindow responsive resizing, time-first vertical header,
explicit cell time badges, prominent dual-clock pill, and GUI <-> HTML export consistency.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QImage, QPainter
from PyQt5.QtWidgets import QApplication, QLabel

from core.ui.font_scale import apply_font_scale, install_font_scaling
from core.ui.theme import C
from core.ui.ui_components import GLOBAL_STYLE
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan.session_plan_models import (
    AgendaSlotItem,
    DelegationMember,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone
from modules.meetings.ui.session_plan.session_plan_window import (
    SessionPlanVerticalHeader,
    SessionPlanWindow,
    _format_compact_snapshot_label,
)


class TestSessionPlanUIAndHTMLConsistency(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)
        cls.app.setStyle("Fusion")
        cls.app.setStyleSheet(GLOBAL_STYLE)
        install_font_scaling(cls.app)
        apply_font_scale(1.0, app=cls.app, save=False)

    def tearDown(self) -> None:
        apply_font_scale(1.0, app=self.app, save=False)

    def test_compact_snapshot_label_formatter(self) -> None:
        full = "S2-2611637 (S2-2611637_SA2-178-Session Plan - r7.docx)"
        compact = _format_compact_snapshot_label(full, "2026-11-14T10:43:12Z")
        self.assertEqual(compact, "S2-2611637 · r7 (2026-11-14T10:43)")

    def test_session_plan_resizable_time_headers_and_html_parity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            td = Path(tmp)
            # Pick a meeting start date anchored to today so live-period highlighting is deterministically testable
            venue_offset_min = -420  # UTC-07:00
            venue_today = SessionPlanTimezone.get_venue_current_datetime(venue_offset_min).date()
            monday_date = venue_today - timedelta(days=venue_today.weekday())

            win = SessionPlanWindow(
                meeting_dir=td,
                mtg_info={
                    "meeting_number": "178",
                    "location": "Calgary",
                    "start_date": monday_date.strftime("%Y-%m-%d"),
                },
            )
            try:
                win.tz_offset_min = venue_offset_min
                win.db.save_delegation_member("178", DelegationMember("Josep", "#7C3AED"))
                win.db.save_delegation_member("178", DelegationMember("Dieter", "#2563EB"))
                win.db.save_delegation_member("178", DelegationMember("Piotr", "#059669"))
                win.db.set_topic_owner("178", "20.6.1", "Josep")
                win.delegation_plan = win.db.get_delegation_plan("178")

                snap = SessionPlanSnapshot(
                    meeting_number="178",
                    tdoc_id="S2-2611637",
                    snapshot_label="S2-2611637 (S2-2611637_SA2-178-Session Plan - r7.docx)",
                    streams={
                        1: MeetingStreamInfo(1, "Stream 1", "Main Room (300)", 300, "MCC", "SA2_Main"),
                        2: MeetingStreamInfo(2, "Stream 2", "Breakout 1 (100)", 100, "Mike", "SA2_Brk1"),
                        3: MeetingStreamInfo(3, "Stream 3", "Breakout 2 (120)", 120, "Wanqiang", "SA2_Brk2"),
                    },
                    slots=[
                        SessionSlot(
                            day_index=0,
                            day_name="Monday",
                            quarter="Q1",
                            sub_slot=1,
                            start_time_local="09:00",
                            end_time_local="09:45",
                            stream_number=1,
                            raw_text="KI#1.1 (NAS) (20.6.1)",
                            items=[AgendaSlotItem("20.6.1", "KI#1.1", "NAS")],
                        ),
                        SessionSlot(
                            day_index=0,
                            day_name="Monday",
                            quarter="Q1",
                            sub_slot=2,
                            start_time_local="09:45",
                            end_time_local="10:30",
                            stream_number=1,
                            raw_text="KI#2 (SBA) (20.6.2)",
                            items=[AgendaSlotItem("20.6.2", "KI#2", "SBA")],
                        ),
                        SessionSlot(
                            day_index=1,
                            day_name="Tuesday",
                            quarter="Q0",
                            sub_slot=0,
                            stream_number=2,
                            raw_text="Drafting:",
                            is_drafting=True,
                        ),
                        SessionSlot(
                            day_index=0,
                            day_name="Monday",
                            quarter="Q3",
                            sub_slot=0,
                            start_time_local="14:00",
                            end_time_local="15:30",
                            stream_number=1,
                            raw_text="KI#18 (AI) (20.6.18)",
                        ),
                        SessionSlot(
                            day_index=4,
                            day_name="Friday",
                            quarter="Q3",
                            sub_slot=0,
                            start_time_local="13:30",
                            end_time_local="16:00",
                            stream_number=1,
                            raw_text="Plenary session (1330 - 1600)",
                        ),
                    ],
                )
                win.db.save_snapshot(snap)
                win._load_data()
                win.resize(980, 760)
                win.show()
                self.app.processEvents()

                # 1. Window is freely resizable (minimum width <= 960 px, half of 1920 px screen)
                self.assertLessEqual(win.minimumSizeHint().width(), 960)
                self.assertLessEqual(win.width(), 1020)

                # 2. Concise navigation labels prevent truncation
                self.assertIn("Agenda (2 Sched / 2)", win.btn_agenda.text())
                self.assertIn("Delegate Agenda (1)", win.btn_tracked.text())
                self.assertIn("S2-2611637 · r7", win.snap_combo.currentText())

                # 3. Prominent dual-clock pill in header
                win._update_live_clock()
                clock_txt = win.clock_lbl.text()
                self.assertIn("Venue", clock_txt)
                self.assertIn("13px", clock_txt)

                # 4. Time-first custom vertical header paints cleanly
                vh = win.grid_table.verticalHeader()
                self.assertIsInstance(vh, SessionPlanVerticalHeader)
                q1_s1_row = next(
                    idx for idx, r in enumerate(win.grid_rows)
                    if r.period == "Q1" and r.stream_number == 1
                )
                self.assertTrue(win.grid_rows[q1_s1_row].display_label.startswith("Q1 · 09:00 – 10:30"))
                win.grid_rows[q1_s1_row].is_live_period = True

                img = QImage(220, 56, QImage.Format_ARGB32)
                img.fill(QColor(C.BG_CARD))
                painter = QPainter(img)
                vh.paintSection(painter, img.rect(), q1_s1_row)
                painter.end()

                # 5. Grid cell widgets show explicit time pills for half-quarters and custom times
                mon_tu1 = win.grid_table.cellWidget(q1_s1_row, 0)
                mon_tu2 = win.grid_table.cellWidget(q1_s1_row, 1)
                self.assertIsNotNone(mon_tu1)
                self.assertIsNotNone(mon_tu2)
                tu1_texts = [lbl.text() for lbl in mon_tu1.findChildren(QLabel)]
                tu2_texts = [lbl.text() for lbl in mon_tu2.findChildren(QLabel)]
                self.assertIn("🕒 09:00–09:45", tu1_texts)
                self.assertIn("🕒 09:45–10:30", tu2_texts)

                q3_s1_row = next(
                    idx for idx, r in enumerate(win.grid_rows)
                    if r.period == "Q3" and r.stream_number == 1
                )
                fri_q3 = win.grid_table.cellWidget(q3_s1_row, 8)
                self.assertIsNotNone(fri_q3)
                q3_texts = [lbl.text() for lbl in fri_q3.findChildren(QLabel)]
                self.assertIn("🕒 13:30–16:00", q3_texts)

                # Empty drafting placeholders remain uncluttered (no time pill)
                q0_s2_row = next(
                    idx for idx, r in enumerate(win.grid_rows)
                    if r.period == "Q0" and r.stream_number == 2
                )
                tue_draft = win.grid_table.cellWidget(q0_s2_row, 2)
                self.assertIsNotNone(tue_draft)
                draft_texts = [lbl.text() for lbl in tue_draft.findChildren(QLabel)]
                self.assertFalse(any("🕒" in t for t in draft_texts))

                # 6. HTML Export mirrors live clock pill, time-first row headers, and cell time pills
                out_html = td / "parity_plan.html"
                SessionPlanExporter.export_mobile_html(
                    snapshot=win.active_snapshot,
                    output_path=out_html,
                    meeting_tz_offset_min=venue_offset_min,
                    tz_name="UTC-07:00 (America/Edmonton)",
                    meeting_dates_str=monday_date.strftime("%Y-%m-%d"),
                    delegation_plan=win.delegation_plan,
                )
                html_txt = out_html.read_text(encoding="utf-8")
                self.assertIn('id="liveClockPill"', html_txt)
                self.assertIn('class="badge-live-clock"', html_txt)
                self.assertIn('class="hdr-time-primary">Q1 · 09:00 – 10:30</span>', html_txt)
                self.assertIn('class="hdr-now-tag">NOW</span>', html_txt)
                self.assertIn('class="cell-time-pill is-half">🕒 09:00–09:45</span>', html_txt)
                self.assertIn('class="cell-time-pill is-half">🕒 09:45–10:30</span>', html_txt)
                self.assertIn('class="cell-time-pill is-full">🕒 14:00–15:30</span>', html_txt)
                self.assertIn('class="cell-time-pill is-custom-time">🕒 13:30–16:00</span>', html_txt)

                # Verify JS updates liveClockPill cleanly in Node.js if available
                node_bin = shutil.which("node")
                if node_bin:
                    scripts = re.findall(r"<script>(.*?)</script>", html_txt, re.DOTALL)
                    full_js = "\n".join(scripts)
                    harness = f"""
const mockEl = () => ({{
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
const clockPill = mockEl();
global.document = {{
  documentElement: {{ style: {{ setProperty: () => {{}} }} }},
  getElementById: (id) => (id === 'liveClockPill' ? clockPill : mockEl()),
  querySelectorAll: () => [],
  querySelector: () => mockEl(),
  body: {{ appendChild: () => {{}}, removeChild: () => {{}} }}
}};
global.window = {{ innerWidth: 1280, innerHeight: 900, addEventListener: () => {{}} }};
global.setInterval = () => {{}};
global.setTimeout = () => {{}};

{full_js}
checkLiveVenueStatus();
if (!clockPill.innerHTML.includes('Venue:')) throw new Error('liveClockPill not updated');
"""
                    js_file = td / "verify_parity.js"
                    js_file.write_text(harness, encoding="utf-8")
                    res = subprocess.run([node_bin, str(js_file)], capture_output=True, text=True, encoding="utf-8", errors="replace")
                    self.assertEqual(res.returncode, 0, f"Node verification failed: {res.stderr}")
            finally:
                win.close()


if __name__ == "__main__":
    unittest.main()
