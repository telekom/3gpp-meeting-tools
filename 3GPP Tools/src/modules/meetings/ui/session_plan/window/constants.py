# --- File: src/modules/meetings/ui/session_plan/window/constants.py ---
"""
Shared constants and small helpers for the Session Plan window.
"""
import logging
import re
from typing import Tuple
from dataclasses import dataclass
from modules.meetings.core.session_plan.session_plan_parser import PERIOD_ORDER, PERIOD_TIMES


logger = logging.getLogger("modules.meetings.ui.session_plan.session_plan_window")

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


QUARTER_ORDER = PERIOD_ORDER


QUARTER_TIMES = PERIOD_TIMES


@dataclass
class GridRowDescriptor:
    period: str                  # "Q0", "Break 1", "Q1", ...
    stream_number: int           # 1, 2, 3 (or 0 for break)
    row_type: str                # "stream" or "break"
    display_label: str           # Vertical header string
    time_range: Tuple[str, str]  # ("09:00", "10:30")
    short_room: str = ""
    local_time_range: Tuple[str, str] = ("", "")
    tu_split_label: str = ""
    is_live_period: bool = False


def _format_compact_snapshot_label(snapshot_label: str, imported_at: str) -> str:
    """Format long snapshot labels concisely for the header combo while preserving revision info."""
    raw = (snapshot_label or "").strip()
    ts = (imported_at or "")[:16]
    m = re.match(r"^([A-Za-z0-9-]+)\s*\((.+)\)$", raw)
    if m:
        tdoc = m.group(1).strip()
        inner = m.group(2).strip()
        inner = re.sub(r"\.(docx?|docm|zip|htm|html)$", "", inner, flags=re.I).strip()
        # Strip repeated "<tdoc>_SA2-NNN-Session Plan - " prefix
        inner = re.sub(rf"^{re.escape(tdoc)}[_\s-]*", "", inner, flags=re.I)
        inner = re.sub(r"^(?:SA\d+-?\d+[-\s]*)?Session\s*Plan\s*[-\s]*", "", inner, flags=re.I).strip(" -_")
        short = f"{tdoc} · {inner}" if inner else tdoc
    else:
        short = raw or "Snapshot"
    if len(short) > 36:
        short = short[:33].rstrip() + "..."
    return f"{short} ({ts})" if ts else short
