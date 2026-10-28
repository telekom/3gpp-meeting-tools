# --- File: src/modules/meetings/core/session_plan_parser.py ---
"""
High-performance parser for 3GPP SA2 Session Plans ("Draft time allocation").
Parses both OpenXML Word (.docx) and Markdown/Text tables into structured
SessionPlanSnapshot, SessionSlot, and AgendaSlotItem models.
Supports TU 1 / TU 2 half-quarter sub-slots, range expansion, and typo resilience.
"""
import hashlib
import logging
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan_models import (
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

        # Find markdown table rows
        table_rows = [l for l in lines if l.startswith("|") and l.endswith("|")]
        data_rows = []
        for tr in table_rows:
            # Strip outer pipes and split
            cells = [c.strip() for c in tr[1:-1].split("|")]
            # Skip separator rows like |---|---|
            if cells and all(re.match(r"^:?-+:?$", c) for c in cells if c):
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
            if p.text.strip():
                full_text.append(p.text.strip())

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
            t_text = " ".join([c.text for row in t.rows[:3] for c in row.cells]).lower()
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
        """Parses a docx Table object handling merged cells and sub-slots."""
        current_quarter = "Q1"
        current_times = ("09:00", "10:30")

        for row_idx, row in enumerate(table.rows):
            # Extract cell texts preserving grid positions
            cells = [c.text.strip() for c in row.cells]
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
                cls._process_break_row(cells, break_period, slots)
                continue

            m_str = re.search(r"Stream\s*(\d+)", third_cell, re.I)
            if not m_str:
                continue

            stream_num = int(m_str.group(1))
            day_cells = cells[3:]
            cls._process_day_cells(day_cells, current_quarter, stream_num, current_times, slots)

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
    ):
        """
        Processes day cells for a single row.
        Handles both 5 cells (with colspan=2 or single cells) and 10 cells (2 sub-slots per day).
        """
        q_times = QUARTER_TIMES.get(quarter, (quarter_times[0], quarter_times[1], quarter_times[0], "", "", quarter_times[1]))
        full_s, full_e, tu1_s, tu1_e, tu2_s, tu2_e = q_times

        # Clean HTML td tags
        cleaned_cells = []
        for c in day_cells:
            # Check for colspan=2 marker
            is_col2 = "colspan=2" in c.lower()
            clean_txt = re.sub(r"<[^>]+>", " ", c).strip()
            cleaned_cells.append((clean_txt, is_col2))

        # Check if row is 10 sub-cells (2 per day) or 5 cells
        if len(cleaned_cells) >= 10:
            for day_idx in range(5):
                day_name = DAY_NAMES[day_idx]
                c_tu1, _ = cleaned_cells[day_idx * 2]
                c_tu2, _ = cleaned_cells[day_idx * 2 + 1]

                if c_tu1 == c_tu2 and c_tu1:
                    # Full Quarter
                    slot = cls._create_slot(
                        day_idx, day_name, quarter, 0, full_s, full_e, stream_num, c_tu1
                    )
                    slots.append(slot)
                else:
                    if c_tu1:
                        slots.append(
                            cls._create_slot(
                                day_idx, day_name, quarter, 1, tu1_s, tu1_e, stream_num, c_tu1
                            )
                        )
                    if c_tu2:
                        slots.append(
                            cls._create_slot(
                                day_idx, day_name, quarter, 2, tu2_s, tu2_e, stream_num, c_tu2
                            )
                        )
        else:
            # 5 cells (1 per day)
            for day_idx, (txt, is_colspan) in enumerate(cleaned_cells[:5]):
                if not txt:
                    continue
                day_name = DAY_NAMES[day_idx]
                sub_slot = 0 if is_colspan else 0
                slot = cls._create_slot(
                    day_idx, day_name, quarter, sub_slot, full_s, full_e, stream_num, txt
                )
                slots.append(slot)

    @classmethod
    def _detect_break_period(cls, c0: str, c1: str, c2: str, current_quarter: str) -> Optional[str]:
        combo = (c0 + " " + c1 + " " + c2).lower()
        if "lunch" in combo:
            return "Lunch"
        if "coffee" in combo:
            return "Coffee 1" if current_quarter in ["Q0", "Q1"] else "Coffee 2"
        if "break" in combo and not re.search(r"Stream\s*\d", c2, re.I):
            return "Break 1" if current_quarter in ["Q0", "Q1"] else "Break 2"
        return None

    @classmethod
    def _process_break_row(cls, cells: List[str], period_name: str, slots: List[SessionSlot]):
        """Processes break and coffee rows across Monday..Friday (both 10-subcolumn and 5-column formats)."""
        day_cells = cells[3:]
        if len(day_cells) >= 10:
            for day_idx in range(5):
                tu1 = day_cells[day_idx * 2].strip()
                tu2 = day_cells[day_idx * 2 + 1].strip()
                if tu1 or tu2:
                    if tu1 == tu2 or not tu2:
                        cls._create_break_slot(day_idx, period_name, 0, tu1 or tu2, slots)
                    elif not tu1:
                        cls._create_break_slot(day_idx, period_name, 2, tu2, slots)
                    else:
                        cls._create_break_slot(day_idx, period_name, 1, tu1, slots)
                        cls._create_break_slot(day_idx, period_name, 2, tu2, slots)
        else:
            for day_idx in range(min(5, len(day_cells))):
                txt = day_cells[day_idx].strip()
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
    ) -> SessionSlot:
        """Parses cell text into AgendaSlotItems and flags metadata."""
        txt_lower = raw_text.lower()
        is_drafting = "drafting" in txt_lower or "soh" in txt_lower
        is_tentative = "(?)" in raw_text or "tentative" in txt_lower
        is_single_stream = "single stream" in txt_lower
        is_plenary = "plenary" in txt_lower or "block approval" in txt_lower

        items = cls._extract_agenda_items(raw_text, is_drafting, is_tentative)

        return SessionSlot(
            day_index=day_idx,
            day_name=day_name,
            quarter=quarter,
            sub_slot=sub_slot,
            start_time_local=start_t,
            end_time_local=end_t,
            stream_number=stream_num,
            raw_text=raw_text,
            is_drafting=is_drafting,
            is_tentative=is_tentative,
            is_single_stream=is_single_stream,
            is_plenary=is_plenary,
            items=items,
        )

    @classmethod
    def _extract_agenda_items(
        cls, raw_text: str, default_drafting: bool = False, default_tentative: bool = False
    ) -> List[AgendaSlotItem]:
        """
        Extracts individual Agenda Items and expands ranges like:
        - KI#18 (AI for Network) (20.6.18) [5+15]
        - KI#12 – KI#16 (20.6.12 - 16) [5,0,2,0,1] -> expands to 20.6.12 .. 20.6.16
        - Pre-Rel-19 maintenance (5.x, 6.x, 7.x) [9] -> expands to 5.*, 6.*, 7.*
        - Handles typos: K#1.2 -> KI#1.2, unclosed UIA_ARC (19.8.2 [4]
        """
        items: List[AgendaSlotItem] = []
        if not raw_text or raw_text.strip() in ("-", "—", "Revisions"):
            return items

        # Split cell on double spaces or newlines where multiple items reside
        tokens = [t.strip() for t in re.split(r"(?<=\])\s+|(?<=\))\s+(?=[A-Z0-9#])|\n+", raw_text) if t.strip()]

        for tok in tokens:
            if not tok or tok in ("-", "—"):
                continue

            # 1. Extract TDoc Count in brackets [5+15] or [7]
            tdoc_count = ""
            m_count = re.search(r"\[([\d\+\,\s]+)\]", tok)
            if m_count:
                tdoc_count = m_count.group(1).strip()

            # 2. Extract AI Number: (20.6.18), (20.6.12 - 16), (19.8.2, (5.x, 6.x, 7.x)
            m_ai = re.search(r"\((\d{1,2}(?:\.\d+|\.x|\s*-\s*\d+)+(?:,\s*(?:\d{1,2}(?:\.\d+|\.x)+))*)\)?", tok)
            ai_str = m_ai.group(1).strip() if m_ai else ""

            if not ai_str:
                # Look for standalone AI lists: e.g. "20.48, 20.49, 20.50" or "4.1"
                m_standalone = re.search(r"\b(\d{1,2}\.\d+(?:[a-z])?(?:,\s*\d{1,2}\.\d+(?:[a-z])?)*)\b", tok)
                if m_standalone:
                    ai_str = m_standalone.group(1).strip()

            # 3. Clean topic label
            # Normalize typos: 'K#1.2' -> 'KI#1.2'
            clean_label = re.sub(r"\bK#(\d+)", r"KI#\1", tok)
            # Remove brackets and parenthesized AI from label
            clean_label = re.sub(r"\[.*?\]", "", clean_label)
            if ai_str:
                clean_label = re.sub(r"\(" + re.escape(ai_str) + r"\)?", "", clean_label)
            clean_label = re.sub(r"\s+", " ", clean_label).strip()

            # Expand ranges if present
            if ai_str and ("-" in ai_str or "–" in ai_str):
                # E.g. "20.6.12 - 16"
                expanded_ais = cls._expand_numeric_range(ai_str)
                counts = [c.strip() for c in tdoc_count.split(",")] if "," in tdoc_count else []
                for idx, single_ai in enumerate(expanded_ais):
                    c_val = counts[idx] if idx < len(counts) else tdoc_count
                    items.append(
                        AgendaSlotItem(
                            ai_number=single_ai,
                            display_expr=tok,
                            topic_label=clean_label,
                            tdoc_count_str=c_val,
                            is_drafting=default_drafting or ("drafting" in tok.lower()),
                            is_tentative=default_tentative or ("(?)" in tok),
                        )
                    )
            elif ai_str and "," in ai_str:
                # E.g. "20.48, 20.49, 20.50"
                for single_ai in [a.strip() for a in ai_str.split(",") if a.strip()]:
                    items.append(
                        AgendaSlotItem(
                            ai_number=single_ai,
                            display_expr=tok,
                            topic_label=clean_label,
                            tdoc_count_str=tdoc_count,
                            is_drafting=default_drafting or ("drafting" in tok.lower()),
                            is_tentative=default_tentative or ("(?)" in tok),
                        )
                    )
            elif ai_str:
                items.append(
                    AgendaSlotItem(
                        ai_number=ai_str,
                        display_expr=tok,
                        topic_label=clean_label,
                        tdoc_count_str=tdoc_count,
                        is_drafting=default_drafting or ("drafting" in tok.lower()),
                        is_tentative=default_tentative or ("(?)" in tok),
                    )
                )
            else:
                # Drafting or general discussion without an explicit AI number
                items.append(
                    AgendaSlotItem(
                        ai_number="",
                        display_expr=tok,
                        topic_label=clean_label or tok,
                        tdoc_count_str=tdoc_count,
                        is_drafting=default_drafting or ("drafting" in tok.lower()),
                        is_tentative=default_tentative or ("(?)" in tok),
                    )
                )

        return items

    @staticmethod
    def _expand_numeric_range(range_str: str) -> List[str]:
        """Expands '20.6.12 - 16' to ['20.6.12', '20.6.13', '20.6.14', '20.6.15', '20.6.16']."""
        clean = re.sub(r"[–—]", "-", range_str)
        m = re.match(r"^(\d+(?:\.\d+)*)\s*-\s*(\d+)$", clean.strip())
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
