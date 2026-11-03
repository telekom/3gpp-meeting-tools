# --- File: src/modules/meetings/core/session_plan_parser.py ---
"""
High-performance, fault-tolerant parser for 3GPP SA2 Session Plans ("Draft time allocation").
Parses both OpenXML Word (.docx) and Markdown/Text tables into structured
SessionPlanSnapshot, SessionSlot, and AgendaSlotItem models.

Supports TU 1 / TU 2 half-quarter sub-slots, range expansion, typo resilience,
and complex real-world 3GPP SA2 drafting conventions.

========================================================================================
SA2 SESSION PLAN CONVENTIONS & PARSING QUIRKS REFERENCE GUIDE
========================================================================================

1. Physical Room Overrides ("In main room", "In Breakout 1", "In Breakout 2"):
   - Leadership occasionally shifts sessions out of their standard stream rooms (e.g.
     Stream 1 moving into Breakout 1, or Stream 2 moving into the Main Room).
   - These directives are written directly inside table cells (e.g. "In Breakout 2\\nKI#8...").
   - If not parsed separately, the room string becomes the Agenda Item's topic description,
     corrupting topic ownership tables, search filters, and agenda matrices.
   - Parser behavior: Regex `ROOM_OVERRIDE_RE` isolates and normalizes the physical room,
     storing it in `slot.room_override`, and purges it from the topic tokens.

2. Disambiguating "Breakout" Rooms from "Coffee Breaks":
   - Room header rows (e.g. "| Main Room | Palomino | Breakout 1 | ...") contain "Breakout 1".
   - A naive `"break" in cell.lower()` test falsely matches "Breakout 1", turning room header
     rows into fake coffee breaks across Monday-Wednesday and shifting all session timings.
   - Parser behavior: Regex word boundary `re.search(r"\\bbreak\\b", ...)` and quarter checks
     ensure "Breakout 1/2" is never mistaken for a coffee or meal break.

3. OpenXML Vertically Merged Cells (vMerge) Across Parallel Streams:
   - Plenary sessions (e.g. Monday Q1/Q2, Friday Q4/Q5) or single-stream sessions span across
     all 3 stream rows using OpenXML vertical cell merges (`<w:vMerge>`).
   - In python-docx, iterating across row cells on vertically merged tables repeats the
     master cell's XML pointer (`_tc`) and text across Streams 1, 2, and 3.
   - If unmanaged, this creates 3 identical duplicate slots in Streams 1, 2, and 3, causing
     the UI to signal false parallel-stream double-booking collisions for delegates.
   - Parser behavior: Tracks underlying XML cell references (`seen_tcs = set()`). Vertically
     merged continuations across parallel streams are deduplicated.

4. Markdown / HTML Table Colspan Expansion:
   - When delegates paste session plan tables from email reflectors or Markdown notes,
     full-quarter sessions (spanning TU1 and TU2) often use `<td colspan=2/>`, resulting in
     fewer than 10 cells in a row (e.g. 8 or 9 cells).
   - Naive slicing (`cells[:5]`) causes severe day-shifting (e.g. Tuesday TU2 shifts into
     Wednesday, dropping Thursday and Friday entirely).
   - Parser behavior: `_process_day_cells` expands `<td colspan=2/>` and 5-cell day rows
     into 10 normalized half-quarter slots, maintaining exact Monday-Friday alignment.

5. Multi-Topic Slots & Numeric Range Expansion:
   - Cells frequently bundle multiple topics on separate lines, e.g.:
     "In Breakout 2\\nKI#10 (FWA) (20.6.10)\\nKI#12 – KI#16 (20.6.12 - 16)"
   - Each line must map to its own specific Agenda Item and topic description.
   - `_expand_numeric_range` supports both shorthand ranges (`20.6.12 - 16`), full prefix
     ranges (`20.6.12 - 20.6.16`), and standalone KI ranges (`KI#12 – KI#16`).

6. Implicit Key Issue (KI#) Mapping:
   - Study items (e.g. Rel-19 FS_AIML_Ph2 under AI 20.6.x) define Key Issues (KI#1..KI#30).
   - Early in the week, chairs write explicit references like "KI#8 (Net Sharing) (20.6.8)".
   - Later in the week, chairs frequently omit the AI number, writing only "KI#8 (Net Sharing)"
     or "KI#21 (Data Framework)".
   - Parser behavior: Dynamically populates a meeting-wide `ki_map` cache during table traversal.
     Subsequent implicit KI# references automatically resolve to their full 20.6.x AI numbers.

7. Balanced Parentheses & Topic Label Hygiene:
   - Stripping extracted AI numbers (e.g. removing `(20.6.8)` from `KI#8 (Net Sharing) (20.6.8)`)
     must not strip genuine topic brackets (e.g. `(Net Sharing)` or `(AI for Network)`),
     nor leave dangling closing brackets from multi-AI lists like `(5.x, 6.x )`.
   - Parser behavior: `_balance_parentheses()` audits bracket parity, stripping unmatched
     outer delimiters while preserving valid technical qualifiers.

8. Timetable Administrative Note Filtering:
   - Chairs frequently insert logistical scheduling notes into session slots, e.g.:
     "Drafting 18:30-20:30", "Session ends at 17:30", "Revisions (tbd)", or standalone dashes.
   - Parser behavior: Regex `ADMIN_NOTE_RE` identifies logistical timetable annotations
     and suppresses them from polluting technical Agenda Item lists.

9. Single-Digit Plenary Agenda Items:
   - Plenary cells frequently list "Opening (1), Agenda (2), Reports (3), Work plan (4), Outgoing LSs (5.1)".
   - Parser behavior: `single_ai_matches` extracts each numbered plenary item as an independent
     agenda item entity with its corresponding title.

10. OpenXML Track Changes (<w:ins> / <w:del> / <w:moveTo> / <w:moveFrom>):
    - Mid-meeting session plan revisions (`-r1.docx`, `-r2.docx`) are distributed with Word
      Track Changes enabled as chairs swap or reschedule topics across days and quarters.
    - `python-docx`'s `cell.text` / `paragraph.text` only traverses direct `<w:p>/<w:r>/<w:t>`
      nodes and completely ignores `<w:ins>` and `<w:moveTo>` containers, causing newly inserted
      or moved sessions (and revised quarter timings) to vanish as empty strings.
    - Parser behavior: `_extract_xml_text` walks the underlying `<w:tc>` / `<w:p>` XML tree,
      accepting all `<w:t>` nodes inside normal runs, `<w:ins>`, `<w:moveTo>`, and hyperlinks,
      while strictly skipping `<w:del>` and `<w:moveFrom>` subtrees.

11. Dynamic Quarter Time Ranges & Half-Quarter Midpoint Calculation:
    - Chairs occasionally shift quarter boundaries (e.g. moving Q2 from `1100 - 1230` to
      `1230 - 1400`, or Q5 to `1745 - 1900`).
    - Parser behavior: `_resolve_quarter_times` honors the document's explicit time column
      and dynamically computes the exact midpoint for TU1 and TU2 half-quarters.
"""
import hashlib
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan.session_plan_models import (
    AgendaSlotItem,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
)

logger = logging.getLogger(__name__)

# Standard SA2 Period times (start, end, tu1_start, tu1_end, tu2_start, tu2_end)
PERIOD_ORDER = [
    "Q0",
    "Break 1",
    "Q1",
    "Coffee 1",
    "Q2",
    "Lunch",
    "Q3",
    "Coffee 2",
    "Q4",
    "Break 2",
    "Q5",
]

PERIOD_TIMES = {
    "Q0": ("08:00", "08:50", "08:00", "08:25", "08:25", "08:50"),
    "Break 1": ("08:50", "09:00", "08:50", "09:00", "08:50", "09:00"),
    "Q1": ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30"),
    "Coffee 1": ("10:30", "11:00", "10:30", "11:00", "10:30", "11:00"),
    "Q2": ("11:00", "12:30", "11:00", "11:45", "11:45", "12:30"),
    "Lunch": ("12:30", "14:00", "12:30", "13:15", "13:15", "14:00"),
    "Q3": ("14:00", "15:30", "14:00", "14:45", "14:45", "15:30"),
    "Coffee 2": ("15:30", "16:00", "15:30", "16:00", "15:30", "16:00"),
    "Q4": ("16:00", "17:30", "16:00", "16:45", "16:45", "17:30"),
    "Break 2": ("17:30", "18:00", "17:30", "18:00", "17:30", "18:00"),
    "Q5": ("18:00", "19:30", "18:00", "18:45", "18:45", "19:30"),
}
QUARTER_TIMES = PERIOD_TIMES

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

ROOM_OVERRIDE_RE = re.compile(
    r"^\s*(?:\(?\s*in\s+(?:the\s+)?)?(main room|breakout\s*1|breakout\s*2|bmo room \d+|palomino [a-z\-]+|percheron [a-z\-]+)\s*\)?",
    re.I
)

ADMIN_NOTE_RE = re.compile(
    r"^\s*(?:\d{1,2}:\d{2}\s+(?:list of|block approval)|agreed tdocs not available|notes are taken by|close of meeting by|when sessions run in parallel)\b",
    re.I
)


class SessionPlanParser:
    """Parses meeting session plan tables and extracts semantic schedule models."""

    @classmethod
    def parse_text_or_markdown(
        cls,
        text: str,
        meeting_number: str = "",
        tdoc_id: str = "",
        snapshot_label: str = "",
        cached_file_path: str = "",
    ) -> SessionPlanSnapshot:
        """Parses a markdown or plain text table representation of the session plan."""
        lines = [line.strip() for line in text.strip().splitlines() if line.strip()]

        # Extract TDoc number from header if not provided
        if not tdoc_id:
            m_tdoc = re.search(r"\b(S[1-5]-\d{7})\b", text, re.I)
            if m_tdoc:
                tdoc_id = m_tdoc.group(1).upper()
            else:
                tdoc_id = "SessionPlan"

        if not snapshot_label:
            snapshot_label = f"{tdoc_id} (Manual Import)"

        # 1. Parse Rooms & Streams from header
        streams = cls._parse_rooms_header(text)

        # 2. Parse table rows
        slots: List[SessionSlot] = []
        current_quarter = "Q1"
        current_times = ("09:00", "10:30")
        ki_map: Dict[str, str] = {}

        # Find markdown table rows
        table_rows = [l for l in lines if l.startswith("|") and l.endswith("|")]
        data_rows = []
        for tr in table_rows:
            # Strip outer pipes and split
            cells = [c.strip() for c in tr[1:-1].split("|")]
            # Skip separator rows like |---|---|
            if cells and all(re.match(r"^:?-+:?$", c) for c in cells if c):
                continue
            # Skip rooms header row: | Main Room | Palomino ... | Breakout 1 | ...
            if len(cells) >= 3 and re.search(r"^\s*(Main Room|Breakout\s*1)\b", cells[0], re.I):
                continue
            # Skip days header row: | | | | Monday ... | Tuesday ...
            if any("monday" in c.lower() for c in cells[:4]):
                continue
            data_rows.append(cells)

        for row in data_rows:
            if not row or len(row) < 3:
                continue

            first_cell = row[0]
            second_cell = row[1] if len(row) > 1 else ""
            third_cell = row[2] if len(row) > 2 else ""

            # Check for Quarter header in first cell
            m_q = re.search(r"\b(Q[0-5]|Lunch|Coffee|Break)\b", first_cell, re.I)
            if m_q:
                q_name = m_q.group(1).capitalize()
                if q_name in QUARTER_TIMES:
                    current_quarter = q_name
                    current_times = (QUARTER_TIMES[q_name][0], QUARTER_TIMES[q_name][1])

            # Check time range in second cell
            m_time = re.search(r"(\d{4})\s*-\s*(\d{4})", second_cell)
            if m_time:
                s_t = f"{m_time.group(1)[:2]}:{m_time.group(1)[2:]}"
                e_t = f"{m_time.group(2)[:2]}:{m_time.group(2)[2:]}"
                current_times = (s_t, e_t)

            # Check for Break / Coffee / Lunch rows
            break_period = cls._detect_break_period(first_cell, second_cell, third_cell, current_quarter)
            if break_period:
                cls._process_break_row(row, break_period, slots)
                continue

            # Check stream in 3rd cell
            m_str = re.search(r"Stream\s*(\d+)", third_cell, re.I)
            if not m_str:
                continue
            stream_num = int(m_str.group(1))

            # Process Day Cells
            day_cells = row[3:]
            cls._process_day_cells(
                day_cells=day_cells,
                quarter=current_quarter,
                stream_num=stream_num,
                quarter_times=current_times,
                slots=slots,
                ki_map=ki_map,
            )

        # Compute content hash
        content_hash = cls._compute_content_hash(slots)

        return SessionPlanSnapshot(
            meeting_number=meeting_number,
            tdoc_id=tdoc_id,
            snapshot_label=snapshot_label,
            content_hash=content_hash,
            cached_file_path=cached_file_path,
            is_active=True,
            streams=streams,
            slots=slots,
        )

    _W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    _IGNORED_XML_TAGS = {
        f"{_W_NS}del",
        f"{_W_NS}moveFrom",
        f"{_W_NS}rPr",
        f"{_W_NS}pPr",
        f"{_W_NS}tcPr",
        f"{_W_NS}tblPr",
        f"{_W_NS}trPr",
    }

    @classmethod
    def _extract_xml_text(cls, xml_elem) -> str:
        """
        Extracts visible/accepted text from an OpenXML element (<w:tc> or <w:p>),
        including text inside Track Changes insertions (<w:ins>, <w:moveTo>), hyperlinks,
        and structured document tags, while strictly ignoring deletions (<w:del>, <w:moveFrom>).
        """
        if xml_elem is None:
            return ""

        w_t = f"{cls._W_NS}t"
        w_br = f"{cls._W_NS}br"
        w_cr = f"{cls._W_NS}cr"
        w_tab = f"{cls._W_NS}tab"
        w_p = f"{cls._W_NS}p"

        def _walk_runs(node) -> List[str]:
            parts: List[str] = []
            for child in node:
                if child.tag in cls._IGNORED_XML_TAGS:
                    continue
                if child.tag == w_t:
                    if child.text:
                        parts.append(child.text)
                elif child.tag in (w_br, w_cr):
                    parts.append("\n")
                elif child.tag == w_tab:
                    parts.append("\t")
                else:
                    parts.extend(_walk_runs(child))
            return parts

        if xml_elem.tag == w_p:
            paragraphs = [xml_elem]
        else:
            paragraphs = xml_elem.findall(f".//{w_p}")

        p_lines: List[str] = []
        for p in paragraphs:
            line = "".join(_walk_runs(p)).strip()
            if line:
                p_lines.append(line)
        return "\n".join(p_lines)

    @classmethod
    def _resolve_quarter_times(
        cls, quarter: str, current_times: Tuple[str, str]
    ) -> Tuple[str, str, str, str, str, str]:
        """
        Computes (full_start, full_end, tu1_start, tu1_end, tu2_start, tu2_end)
        honoring document-specified current_times and splitting at the exact midpoint.
        """
        default_q = QUARTER_TIMES.get(
            quarter, ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30")
        )
        s_str, e_str = (
            current_times
            if (current_times and len(current_times) == 2 and current_times[0] and current_times[1])
            else (default_q[0], default_q[1])
        )
        if (s_str, e_str) == (default_q[0], default_q[1]):
            return default_q
        try:
            sh, sm = map(int, s_str.split(":"))
            eh, em = map(int, e_str.split(":"))
            s_min = sh * 60 + sm
            e_min = eh * 60 + em
            if e_min > s_min:
                mid_min = s_min + (e_min - s_min) // 2
                mid_str = f"{mid_min // 60:02d}:{mid_min % 60:02d}"
                return (s_str, e_str, s_str, mid_str, mid_str, e_str)
        except Exception:
            pass
        return (s_str, e_str, s_str, default_q[3], default_q[4], e_str)

    @classmethod
    def parse_docx(
        cls,
        docx_path: Path,
        meeting_number: str = "",
        tdoc_id: str = "",
        snapshot_label: str = "",
    ) -> SessionPlanSnapshot:
        """Parses an OpenXML .docx session plan table using python-docx without Word COM."""
        import docx

        docx_path = Path(docx_path)
        doc = docx.Document(str(docx_path))

        full_text = []
        for p in doc.paragraphs:
            p_txt = cls._extract_xml_text(p._p).strip()
            if p_txt:
                full_text.append(p_txt)

        header_text = "\n".join(full_text[:30])

        if not tdoc_id:
            m_tdoc = re.search(r"\b(S[1-5]-\d{7})\b", header_text, re.I)
            if m_tdoc:
                tdoc_id = m_tdoc.group(1).upper()
            else:
                tdoc_id = docx_path.stem

        if not snapshot_label:
            snapshot_label = f"{tdoc_id} ({docx_path.name})"

        streams = cls._parse_rooms_header(header_text)

        # Locate session plan table
        target_table = None
        for t in doc.tables:
            # Look for table with Q0/Q1 and Monday..Friday
            t_text = " ".join([cls._extract_xml_text(c._tc) for row in t.rows[:3] for c in row.cells]).lower()
            if "monday" in t_text or "stream 1" in t_text or "q1" in t_text:
                target_table = t
                break

        if not target_table and doc.tables:
            target_table = doc.tables[0]

        slots: List[SessionSlot] = []
        if target_table:
            cls._parse_docx_table(target_table, streams, slots)

        content_hash = cls._compute_content_hash(slots)

        return SessionPlanSnapshot(
            meeting_number=meeting_number,
            tdoc_id=tdoc_id,
            snapshot_label=snapshot_label,
            content_hash=content_hash,
            cached_file_path=str(docx_path),
            is_active=True,
            streams=streams,
            slots=slots,
        )

    @classmethod
    def _parse_docx_table(cls, table, streams: Dict[int, MeetingStreamInfo], slots: List[SessionSlot]):
        """Parses a docx Table object handling merged cells, Track Changes, and sub-slots."""
        current_quarter = "Q1"
        current_times = ("09:00", "10:30")
        seen_tcs = set()
        ki_map: Dict[str, str] = {}

        for row_idx, row in enumerate(table.rows):
            # Extract cell texts preserving grid positions and accepting Track Changes (<w:ins>)
            cells = [cls._extract_xml_text(c._tc).strip() for c in row.cells]
            if not cells or len(cells) < 3:
                continue

            first_cell = cells[0]
            second_cell = cells[1] if len(cells) > 1 else ""
            third_cell = cells[2] if len(cells) > 2 else ""

            m_q = re.search(r"\b(Q[0-5]|Lunch|Coffee|Break)\b", first_cell, re.I)
            if m_q:
                q_name = m_q.group(1).capitalize()
                if q_name in QUARTER_TIMES:
                    current_quarter = q_name
                    current_times = (QUARTER_TIMES[q_name][0], QUARTER_TIMES[q_name][1])

            m_time = re.search(r"(\d{4})\s*-\s*(\d{4})", second_cell)
            if m_time:
                s_t = f"{m_time.group(1)[:2]}:{m_time.group(1)[2:]}"
                e_t = f"{m_time.group(2)[:2]}:{m_time.group(2)[2:]}"
                current_times = (s_t, e_t)

            # Check for Break / Coffee / Lunch rows
            break_period = cls._detect_break_period(first_cell, second_cell, third_cell, current_quarter)
            if break_period:
                dedup_cells = list(cells[:3])
                row_tcs = set()
                for c_idx, c_obj in enumerate(row.cells[3:], start=3):
                    if c_obj._tc in seen_tcs:
                        dedup_cells.append("")
                    else:
                        dedup_cells.append(cells[c_idx])
                    row_tcs.add(c_obj._tc)
                seen_tcs.update(row_tcs)
                cls._process_break_row(dedup_cells, break_period, slots)
                continue

            m_str = re.search(r"Stream\s*(\d+)", third_cell, re.I)
            if not m_str:
                continue

            stream_num = int(m_str.group(1))

            if len(row.cells) >= 13:
                # OpenXML table with 10 sub-cells
                q_times = cls._resolve_quarter_times(current_quarter, current_times)
                for day_idx in range(5):
                    c_tu1_obj = row.cells[3 + day_idx * 2]
                    c_tu2_obj = row.cells[3 + day_idx * 2 + 1]

                    tu1_is_dup = c_tu1_obj._tc in seen_tcs
                    tu2_is_dup = c_tu2_obj._tc in seen_tcs

                    seen_tcs.add(c_tu1_obj._tc)
                    seen_tcs.add(c_tu2_obj._tc)

                    if tu1_is_dup and tu2_is_dup:
                        # Full vertical merge across streams
                        continue

                    t1_txt = "" if tu1_is_dup else cells[3 + day_idx * 2]
                    t2_txt = "" if tu2_is_dup else cells[3 + day_idx * 2 + 1]
                    if not t1_txt and not t2_txt:
                        continue

                    day_name = DAY_NAMES[day_idx]
                    if t1_txt == t2_txt and t1_txt:
                        slot = cls._create_slot(day_idx, day_name, current_quarter, 0, q_times[0], q_times[1], stream_num, t1_txt, ki_map)
                        slots.append(slot)
                    else:
                        if t1_txt:
                            slot = cls._create_slot(day_idx, day_name, current_quarter, 1, q_times[2], q_times[3], stream_num, t1_txt, ki_map)
                            slots.append(slot)
                        if t2_txt:
                            slot = cls._create_slot(day_idx, day_name, current_quarter, 2, q_times[4], q_times[5], stream_num, t2_txt, ki_map)
                            slots.append(slot)
            else:
                day_cells = cells[3:]
                cls._process_day_cells(day_cells, current_quarter, stream_num, current_times, slots, ki_map)

    @classmethod
    def _parse_rooms_header(cls, text: str) -> Dict[int, MeetingStreamInfo]:
        """Extracts Stream 1, 2, 3 room names, capacities, conveners, and TOHRU channels."""
        streams: Dict[int, MeetingStreamInfo] = {
            1: MeetingStreamInfo(1, "Stream 1", "Main Room", 300, "SA2 Chair / MCC", "SA2_Main"),
            2: MeetingStreamInfo(2, "Stream 2", "Breakout 1", 100, "Convener 1", "SA2_Brk1"),
            3: MeetingStreamInfo(3, "Stream 3", "Breakout 2", 120, "Convener 2", "SA2_Brk2"),
        }

        # Room regex: | Main Room | B2 (Floor 2) [300] | Breakout 1 | B3 (Floor 2) [100] | Breakout 2 | D11/D12 (Floor 3) [120] |
        m_main = re.search(r"Main Room\s*\|\s*([^|\[\]]+(?:\([^)]+\))?)(?:\[(\d+)\])?", text, re.I)
        if m_main:
            r_name = m_main.group(1).strip()
            cap = int(m_main.group(2)) if m_main.group(2) else 300
            streams[1].room_name = r_name
            streams[1].capacity = cap

        m_brk1 = re.search(r"Breakout 1\s*\|\s*([^|\[\]]+(?:\([^)]+\))?)(?:\[(\d+)\])?", text, re.I)
        if m_brk1:
            r_name = m_brk1.group(1).strip()
            cap = int(m_brk1.group(2)) if m_brk1.group(2) else 100
            streams[2].room_name = r_name
            streams[2].capacity = cap

        m_brk2 = re.search(r"Breakout 2\s*\|\s*([^|\[\]]+(?:\([^)]+\))?)(?:\[(\d+)\])?", text, re.I)
        if m_brk2:
            r_name = m_brk2.group(1).strip()
            cap = int(m_brk2.group(2)) if m_brk2.group(2) else 120
            streams[3].room_name = r_name
            streams[3].capacity = cap

        # Conveners: Convener 1: Mike, Convenor 2: Wanqiang
        m_conv = re.search(r"Conven[eo]r\s*1\s*:\s*([^,\n\|]+),\s*Conven[eo]r\s*2\s*:\s*([^,\n\|]+)", text, re.I)
        if m_conv:
            streams[2].convener = m_conv.group(1).strip()
            streams[3].convener = m_conv.group(2).strip()

        return streams

    @classmethod
    def _process_day_cells(
        cls,
        day_cells: List[str],
        quarter: str,
        stream_num: int,
        quarter_times: Tuple[str, str],
        slots: List[SessionSlot],
        ki_map: Optional[Dict[str, str]] = None,
    ):
        """
        Processes day cells for a single row.
        Handles both 5 cells (with colspan=2 or single cells) and 10 cells (2 sub-slots per day).
        """
        q_times = cls._resolve_quarter_times(quarter, quarter_times)
        full_s, full_e, tu1_s, tu1_e, tu2_s, tu2_e = q_times

        units = []
        for c in day_cells:
            m_span = re.search(r'colspan\s*=\s*["\']?(\d+)["\']?', c, re.I)
            span = int(m_span.group(1)) if m_span else 1
            clean_txt = re.sub(r"<[^>]+>", " ", c).strip()
            clean_txt = re.sub(r"colspan\s*=\s*[\d\"']+", "", clean_txt, flags=re.I).strip(" <>/")
            for _ in range(span):
                units.append((clean_txt, span > 1))

        if len(units) == 5:
            # 5 cells (1 per day)
            for day_idx in range(5):
                txt, is_colspan = units[day_idx]
                if not txt:
                    continue
                day_name = DAY_NAMES[day_idx]
                slot = cls._create_slot(
                    day_idx, day_name, quarter, 0, full_s, full_e, stream_num, txt, ki_map
                )
                slots.append(slot)
        else:
            while len(units) < 10:
                units.append(("", False))
            for day_idx in range(5):
                day_name = DAY_NAMES[day_idx]
                c_tu1, _ = units[day_idx * 2]
                c_tu2, _ = units[day_idx * 2 + 1]

                if c_tu1 == c_tu2 and c_tu1:
                    slot = cls._create_slot(
                        day_idx, day_name, quarter, 0, full_s, full_e, stream_num, c_tu1, ki_map
                    )
                    slots.append(slot)
                else:
                    if c_tu1:
                        slots.append(
                            cls._create_slot(
                                day_idx, day_name, quarter, 1, tu1_s, tu1_e, stream_num, c_tu1, ki_map
                            )
                        )
                    if c_tu2:
                        slots.append(
                            cls._create_slot(
                                day_idx, day_name, quarter, 2, tu2_s, tu2_e, stream_num, c_tu2, ki_map
                            )
                        )

    @classmethod
    def _detect_break_period(cls, c0: str, c1: str, c2: str, current_quarter: str) -> Optional[str]:
        combo = (c0 + " " + c1 + " " + c2).lower()
        if "lunch" in combo:
            return "Lunch"
        if "coffee" in combo:
            return "Coffee 1" if current_quarter in ["Q0", "Q1"] else "Coffee 2"
        if re.search(r"\bbreak\b", combo) and not re.search(r"Stream\s*\d", c2, re.I):
            return "Break 1" if current_quarter in ["Q0", "Q1"] else "Break 2"
        return None

    @classmethod
    def _process_break_row(cls, cells: List[str], period_name: str, slots: List[SessionSlot]):
        """Processes break and coffee rows across Monday..Friday (both 10-subcolumn and 5-column formats)."""
        day_cells = cells[3:]
        units = []
        for c in day_cells:
            m_span = re.search(r'colspan\s*=\s*["\']?(\d+)["\']?', c, re.I)
            span = int(m_span.group(1)) if m_span else 1
            clean_txt = re.sub(r"<[^>]+>", " ", c).strip()
            clean_txt = re.sub(r"colspan\s*=\s*[\d\"']+", "", clean_txt, flags=re.I).strip(" <>/")
            for _ in range(span):
                units.append(clean_txt)

        if len(units) >= 10:
            for day_idx in range(5):
                tu1 = units[day_idx * 2].strip()
                tu2 = units[day_idx * 2 + 1].strip()
                if tu1 or tu2:
                    if tu1 == tu2 or not tu2:
                        cls._create_break_slot(day_idx, period_name, 0, tu1 or tu2, slots)
                    elif not tu1:
                        cls._create_break_slot(day_idx, period_name, 2, tu2, slots)
                    else:
                        cls._create_break_slot(day_idx, period_name, 1, tu1, slots)
                        cls._create_break_slot(day_idx, period_name, 2, tu2, slots)
        else:
            for day_idx in range(min(5, len(units))):
                txt = units[day_idx].strip()
                if txt:
                    cls._create_break_slot(day_idx, period_name, 0, txt, slots)

    @classmethod
    def _create_break_slot(
        cls,
        day_idx: int,
        period_name: str,
        sub_slot: int,
        raw_text: str,
        slots: List[SessionSlot],
    ):
        """Creates a SessionSlot for drafting sessions or plenary schedules inside breaks."""
        txt_clean = re.sub(r"<[^>]+>", " ", raw_text).strip()
        if not txt_clean or txt_clean == "-":
            return

        times = PERIOD_TIMES.get(period_name, ("12:30", "14:00", "12:30", "13:15", "13:15", "14:00"))
        start_t, end_t = times[0], times[1]

        m_t = re.search(r"(\d{2}):(\d{2})", txt_clean)
        if m_t:
            start_t = f"{m_t.group(1)}:{m_t.group(2)}"

        stream_num = 1 if ("main room" in txt_clean.lower() or "stream 1" in txt_clean.lower()) else 0

        slot = cls._create_slot(
            day_idx=day_idx,
            day_name=DAY_NAMES[day_idx],
            quarter=period_name,
            sub_slot=sub_slot,
            start_t=start_t,
            end_t=end_t,
            stream_num=stream_num,
            raw_text=txt_clean,
        )
        slots.append(slot)

    @classmethod
    def _extract_room_override_and_tokens(cls, raw_text: str) -> Tuple[str, List[str]]:
        """Separates physical room overrides (Breakout 1/2, Main Room) from agenda topics."""
        t = raw_text.replace('\xa0', ' ')
        raw_tokens = re.split(r"\n+|\s{2,}|\s*(?<=\])\s+|\s*(?<=\))\s+(?=[A-Z0-9#])", t)

        room_override = ""
        clean_tokens = []

        for tok in raw_tokens:
            tok_s = tok.strip()
            if not tok_s:
                continue
            m_room = ROOM_OVERRIDE_RE.match(tok_s)
            if m_room:
                r_str = m_room.group(1).title()
                if "Main" in r_str:
                    room_override = "Main Room"
                elif "Breakout 1" in r_str or "Bmo" in r_str:
                    room_override = "Breakout 1"
                elif "Breakout 2" in r_str or "Percheron" in r_str:
                    room_override = "Breakout 2"
                else:
                    room_override = r_str
                rest = tok_s[m_room.end():].strip(" :) -–—")
                if rest:
                    clean_tokens.append(rest)
            else:
                clean_tokens.append(tok_s)

        # Join continuation lines: e.g. "Pre-Rel-19 maintenance" + "(5.x, 6.x, 7.x )"
        merged: List[str] = []
        for tok_s in clean_tokens:
            if ADMIN_NOTE_RE.search(tok_s):
                continue
            if merged and tok_s.startswith("(") and re.search(r"\d", tok_s) and not re.search(r"\(\d", merged[-1]):
                merged[-1] = f"{merged[-1]} {tok_s}"
            elif merged and (tok_s.startswith("(") or not re.search(r"\(|\b(?:KI#|FS_|AIML_|\d{1,2}\.\d+)", tok_s)):
                merged[-1] = f"{merged[-1]} {tok_s}"
            else:
                merged.append(tok_s)

        return room_override, merged

    @staticmethod
    def _balance_parentheses(s: str) -> str:
        """Strips leftover unmatched parentheses and surrounding punctuation without unclosing labels."""
        s = s.strip(" -–—,")
        while s.endswith(")") and s.count(")") > s.count("("):
            s = s[:-1].strip()
        while s.startswith("(") and s.count("(") > s.count(")"):
            s = s[1:].strip()
        return s

    @classmethod
    def _create_slot(
        cls,
        day_idx: int,
        day_name: str,
        quarter: str,
        sub_slot: int,
        start_t: str,
        end_t: str,
        stream_num: int,
        raw_text: str,
        ki_map: Optional[Dict[str, str]] = None,
    ) -> SessionSlot:
        """Parses cell text into AgendaSlotItems and flags metadata."""
        txt_lower = raw_text.lower()
        is_drafting = "drafting" in txt_lower or "soh" in txt_lower
        is_tentative = "(?)" in raw_text or "tentative" in txt_lower
        is_single_stream = "single stream" in txt_lower
        is_plenary = "plenary" in txt_lower or "block approval" in txt_lower

        room_override, items = cls._extract_agenda_items(raw_text, is_drafting, is_tentative, ki_map)

        return SessionSlot(
            day_index=day_idx,
            day_name=day_name,
            quarter=quarter,
            sub_slot=sub_slot,
            start_time_local=start_t,
            end_time_local=end_t,
            stream_number=stream_num,
            room_override=room_override,
            raw_text=raw_text,
            is_drafting=is_drafting,
            is_tentative=is_tentative,
            is_single_stream=is_single_stream,
            is_plenary=is_plenary,
            items=items,
        )

    @classmethod
    def _extract_agenda_items(
        cls,
        raw_text: str,
        default_drafting: bool = False,
        default_tentative: bool = False,
        ki_map: Optional[Dict[str, str]] = None,
    ) -> Tuple[str, List[AgendaSlotItem]]:
        """
        Extracts room override and individual Agenda Items, expanding ranges like:
        - KI#18 (AI for Network) (20.6.18) [5+15]
        - KI#12 – KI#16 (20.6.12 - 16) [5,0,2,0,1] -> expands to 20.6.12 .. 20.6.16
        - Pre-Rel-19 maintenance (5.x, 6.x, 7.x) [9] -> expands to 5.*, 6.*, 7.*
        - Handles typos: K#1.2 -> KI#1.2, unclosed UIA_ARC (19.8.2 [4]
        - Resolves implicit KI numbers: KI#21 -> 20.6.21
        """
        items: List[AgendaSlotItem] = []
        if not raw_text or raw_text.strip() in ("-", "—", "Revisions", "Revisions (tbd)"):
            return "", items

        if ki_map is None:
            ki_map = {}

        room_override, tokens = cls._extract_room_override_and_tokens(raw_text)

        for tok in tokens:
            tok_clean = tok.strip()
            if not tok_clean or tok_clean in ("-", "—", "Revisions", "Revisions (tbd)"):
                continue
            if ADMIN_NOTE_RE.search(tok_clean):
                continue

            # Single AIs like "Opening (1), Agenda (2), Reports (3), Outgoing LSs (5.1)"
            single_ai_matches = list(re.finditer(r"([A-Za-z\s]+)\((\d{1,2}(?:\.\d+)?)\)", tok_clean))
            if len(single_ai_matches) >= 2:
                for sm in single_ai_matches:
                    t_lbl = sm.group(1).strip(" ,")
                    a_num = sm.group(2).strip()
                    items.append(
                        AgendaSlotItem(
                            ai_number=a_num,
                            display_expr=f"{t_lbl} ({a_num})",
                            topic_label=t_lbl,
                            is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                            is_tentative=default_tentative or ("(?)" in tok_clean),
                        )
                    )
                continue

            # 1. Extract TDoc Count in brackets [5+15] or [7]
            tdoc_count = ""
            m_count = re.search(r"\[([\d\+\,\s]+)\]", tok_clean)
            if m_count:
                tdoc_count = m_count.group(1).strip()

            # 2. Extract AI Number: (20.6.18), (20.6.12 - 16), (19.8.2, (5.x, 6.x, 7.x )
            m_ai = re.search(r"\(\s*(\d{1,2}(?:\.\d+|\.x|\s*-\s*\d+)+(?:,\s*(?:\d{1,2}(?:\.\d+|\.x)+))*\s*)\)", tok_clean)
            ai_str = m_ai.group(1).strip() if m_ai else ""

            if not ai_str:
                # Look for standalone AI lists: e.g. "20.48, 20.49, 20.50" or "4.1"
                m_standalone = re.search(r"\b(\d{1,2}\.\d+(?:[a-z])?(?:,\s*\d{1,2}\.\d+(?:[a-z])?)*)\b", tok_clean)
                if m_standalone:
                    ai_str = m_standalone.group(1).strip()

            # 3. Check KI# mapping if ai_str is still empty: e.g. "KI#21 (Data Framework)" or "KI#12 – KI#16"
            m_ki_range = re.search(r"\bKI\s*#?\s*(\d+)\s*(?:-|–|—)\s*(?:KI\s*#?)?(\d+)\b", tok_clean, re.I)
            if not ai_str and m_ki_range:
                start_k, end_k = int(m_ki_range.group(1)), int(m_ki_range.group(2))
                ai_str = f"20.6.{start_k} - 20.6.{end_k}"
            else:
                m_ki = re.search(r"\bKI\s*#?\s*(\d+(?:\.\d+)?)\b", tok_clean, re.I)
                ki_num = m_ki.group(1) if m_ki else ""
                if not ai_str and ki_num:
                    if ki_num in ki_map:
                        ai_str = ki_map[ki_num]
                    elif "." in ki_num:
                        ai_str = f"20.6.{ki_num.split('.')[0]}"
                    else:
                        ai_str = f"20.6.{ki_num}"
                elif ki_num and ai_str:
                    ki_map[ki_num] = ai_str

            # 4. Clean topic label
            clean_label = re.sub(r"\bK#(\d+)", r"KI#\1", tok_clean)
            clean_label = re.sub(r"\[.*?\]", "", clean_label)
            if ai_str:
                clean_label = re.sub(r"\(\s*" + re.escape(ai_str) + r"\s*\)", "", clean_label)
            clean_label = re.sub(r"\s+", " ", clean_label)
            clean_label = cls._balance_parentheses(clean_label)

            # Expand ranges if present
            if ai_str and ("-" in ai_str or "–" in ai_str):
                expanded_ais = cls._expand_numeric_range(ai_str)
                counts = [c.strip() for c in tdoc_count.split(",")] if "," in tdoc_count else []
                for idx, single_ai in enumerate(expanded_ais):
                    c_val = counts[idx] if idx < len(counts) else tdoc_count
                    items.append(
                        AgendaSlotItem(
                            ai_number=single_ai,
                            display_expr=tok_clean,
                            topic_label=clean_label or tok_clean,
                            tdoc_count_str=c_val,
                            is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                            is_tentative=default_tentative or ("(?)" in tok_clean),
                        )
                    )
            elif ai_str and "," in ai_str:
                for single_ai in [a.strip() for a in ai_str.split(",") if a.strip()]:
                    items.append(
                        AgendaSlotItem(
                            ai_number=single_ai,
                            display_expr=tok_clean,
                            topic_label=clean_label or single_ai,
                            tdoc_count_str=tdoc_count,
                            is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                            is_tentative=default_tentative or ("(?)" in tok_clean),
                        )
                    )
            elif ai_str:
                items.append(
                    AgendaSlotItem(
                        ai_number=ai_str,
                        display_expr=tok_clean,
                        topic_label=clean_label or tok_clean,
                        tdoc_count_str=tdoc_count,
                        is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                        is_tentative=default_tentative or ("(?)" in tok_clean),
                    )
                )
            else:
                if clean_label:
                    items.append(
                        AgendaSlotItem(
                            ai_number="",
                            display_expr=tok_clean,
                            topic_label=clean_label,
                            tdoc_count_str=tdoc_count,
                            is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                            is_tentative=default_tentative or ("(?)" in tok_clean),
                        )
                    )

        return room_override, items

    @staticmethod
    def _expand_numeric_range(range_str: str) -> List[str]:
        """Expands '20.6.12 - 16' or '20.6.12 - 20.6.16' to ['20.6.12', '20.6.13', '20.6.14', '20.6.15', '20.6.16']."""
        clean = re.sub(r"[–—]", "-", range_str)
        m = re.match(r"^(\d+(?:\.\d+)*)\s*-\s*(?:.*?\.)?(\d+)$", clean.strip())
        if not m:
            return [range_str]
        prefix_part, end_num_str = m.group(1), m.group(2)
        parts = prefix_part.split(".")
        if len(parts) < 2:
            return [range_str]

        start_num = int(parts[-1])
        end_num = int(end_num_str)
        base = ".".join(parts[:-1])

        if start_num <= end_num and end_num - start_num <= 30:
            return [f"{base}.{i}" for i in range(start_num, end_num + 1)]
        return [range_str]

    @staticmethod
    def _compute_content_hash(slots: List[SessionSlot]) -> str:
        """Calculates a deterministic SHA-256 hash of slot schedule contents."""
        hasher = hashlib.sha256()
        for s in sorted(slots, key=lambda x: (x.day_index, x.quarter, x.sub_slot, x.stream_number)):
            hasher.update(f"{s.day_index}|{s.quarter}|{s.sub_slot}|{s.stream_number}|{s.raw_text.strip()}".encode("utf-8"))
        return hasher.hexdigest()

    @classmethod
    def stage_session_plan_file(
        cls,
        source_path: Path,
        meeting_dir: Path,
        tdoc_id: str = "SessionPlan",
    ) -> Path:
        """
        Safely copies the imported session plan into <meeting_dir>/Agenda/SessionPlans/.
        If a file with the same name already exists, appends a collision-free timestamp/snap suffix.
        """
        source_path = Path(source_path)
        meeting_dir = Path(meeting_dir)
        target_dir = meeting_dir / "Agenda" / "SessionPlans"
        target_dir.mkdir(parents=True, exist_ok=True)

        target_file = target_dir / source_path.name

        # If identical file exists on disk, keep it
        if target_file.exists():
            # Check if contents differ
            src_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
            dst_hash = hashlib.sha256(target_file.read_bytes()).hexdigest()
            if src_hash == dst_hash:
                return target_file

            # Differentiate revision name
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            target_file = target_dir / f"{source_path.stem}_{ts}{source_path.suffix}"

        shutil.copy2(source_path, target_file)
        logger.info(f"[SessionPlanParser] Staged session plan file to: {target_file}")
        return target_file
