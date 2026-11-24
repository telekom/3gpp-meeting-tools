# --- File: src/modules/meetings/core/session_plan/parser/markdown_parser.py ---
"""Markdown/plain-text table parsing, content hashing, and file staging."""

import hashlib
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from modules.meetings.core.session_plan.session_plan_models import (
    SessionPlanSnapshot,
    SessionSlot,
)

from .constants import QUARTER_TIMES

logger = logging.getLogger(__name__)


class MarkdownParserMixin:
    """Mixin handling Markdown/text session plan parsing, hashing, and file staging."""

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
