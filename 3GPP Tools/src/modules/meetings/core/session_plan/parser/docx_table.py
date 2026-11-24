# --- File: src/modules/meetings/core/session_plan/parser/docx_table.py ---
"""OpenXML (.docx) table traversal, Track Changes handling, and grid slot builders."""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan.session_plan_models import (
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
)

from .constants import (
    DAY_NAMES,
    PERIOD_TIMES,
    QUARTER_TIMES,
    _IGNORED_XML_TAGS,
    _W_NS,
)


class DocxTableParserMixin:
    """Mixin handling OpenXML cell extraction, quarter time resolution, and slot construction."""

    _W_NS = _W_NS
    _IGNORED_XML_TAGS = _IGNORED_XML_TAGS

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

        for _row_idx, row in enumerate(table.rows):
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
                        slot = cls._create_slot(
                            day_idx, day_name, current_quarter, 0, q_times[0], q_times[1], stream_num, t1_txt, ki_map
                        )
                        slots.append(slot)
                    else:
                        if t1_txt:
                            slot = cls._create_slot(
                                day_idx, day_name, current_quarter, 1, q_times[2], q_times[3], stream_num, t1_txt, ki_map
                            )
                            slots.append(slot)
                        if t2_txt:
                            slot = cls._create_slot(
                                day_idx, day_name, current_quarter, 2, q_times[4], q_times[5], stream_num, t2_txt, ki_map
                            )
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
                txt, _is_colspan = units[day_idx]
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
