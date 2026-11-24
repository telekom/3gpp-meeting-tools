# --- File: src/modules/meetings/core/session_plan/parser/constants.py ---
"""Constants, period schedules, and regex patterns for SA2 Session Plan parsing."""

import re

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
    re.I,
)

ADMIN_NOTE_RE = re.compile(
    r"^\s*(?:\d{1,2}:\d{2}\s+(?:list of|block approval)|agreed tdocs not available|notes are taken by|close of meeting by|when sessions run in parallel)\b",
    re.I,
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
