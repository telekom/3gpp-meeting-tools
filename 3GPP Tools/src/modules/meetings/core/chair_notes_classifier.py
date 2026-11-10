# --- File: src/modules/meetings/core/chair_notes_classifier.py ---
"""
Classification and recency scoring engine for 3GPP Chairman's Notes and track agendas.

Supports multi-track meetings (such as SA2 per-chair track files: Andy, Hector,
Ruicheng, Jian), breakout sessions, WID/agenda prefixes (RAN3), and consolidated
closing reports across all 3GPP Working Groups.
"""
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# Reserved tokens that denote document status rather than a specific Vice Chair or track
_RESERVED_STATUS_TOKENS = {
    "final",
    "eom",
    "draft",
    "clean",
    "cl",
    "rm",
    "tracked",
    "track",
    "session",
    "breakout",
    "changes",
    "v",
    "v0",
    "v1",
    "v2",
    "v3",
    "v4",
    "v01",
    "v02",
    "v03",
    "v04",
    "rev",
    "r",
    "end",
    "closing",
    "agenda",
    "minutes",
    "summary",
    "report",
}


def classify_chair_track(filename: str) -> Tuple[str, str]:
    """
    Identifies the logical track, chair author, or topic group for a Chairman's Notes file.

    Returns:
        (display_label: str, track_id: str)
        e.g. ("Chair: Andy", "chair_andy") or ("Session A", "session_a") or ("Consolidated / Main", "main")
    """
    if not filename:
        return "Consolidated / Main", "main"

    stem = Path(filename).stem
    low = stem.lower()

    # 1. Breakout / Session / Track: sessionA, breakout_1, track_2
    m_sess = re.search(r'(?:session|breakout|track)[_\-\s]*([a-zA-Z0-9]+)', stem, re.IGNORECASE)
    if m_sess:
        sess_name = m_sess.group(1).upper()
        return f"Session {sess_name}", f"session_{sess_name.lower()}"

    # 2. Dedicated WG agenda / notes prefix: RAN3_133bis_agenda
    m_wg_ag = re.search(r'^([A-Z0-9]{2,6}[_\-][a-zA-Z0-9]+[_\-]agenda)', stem, re.IGNORECASE)
    if m_wg_ag:
        pfx = m_wg_ag.group(1).upper()
        return f"Agenda: {pfx}", f"ag_{pfx.lower()}"

    # 3. SA2 / Working Group per-chair files: ChairNotes_Andy, Chair_Notes_Hector, ChairNotes_Ruicheng
    m_chair = re.search(r'chair[_\-\s]?notes?[_\-\s]+([a-zA-Z]+)', stem, re.IGNORECASE)
    if m_chair:
        cand = m_chair.group(1).strip()
        cand_low = cand.lower()
        if (
            cand_low not in _RESERVED_STATUS_TOKENS
            and not cand_low.startswith(("session", "breakout", "track"))
            and not re.match(r'^v\d+', cand_low)
        ):
            return f"Chair: {cand.capitalize()}", f"chair_{cand_low}"

    # 4. TDoc ID in filename: S2-2612345
    m_tdoc = re.search(r'([A-Z][A-Z0-9]{1,4}-\d{4,8})', stem, re.IGNORECASE)
    if m_tdoc:
        td = m_tdoc.group(1).upper()
        return f"TDoc {td}", f"tdoc_{td.lower()}"

    # 5. Fallback: Consolidated / Main Track
    return "Consolidated / Main", "main"


def compute_recency_score(filename: str) -> int:
    """
    Computes a numerical recency and completeness score for ranking multiple versions
    belonging to the same chair or track.

    Higher scores indicate newer, more authoritative, or clean final versions.
    """
    if not filename:
        return 0

    stem = Path(filename).stem
    low = stem.lower()
    score = 0

    # 1. Final / EOM / End of meeting marker (authoritative closing documents)
    if any(k in low for k in ("final", "eom", "closing", "end_of_meeting", "end-of-meeting")):
        score += 100_000_000_000

    # 2. Prefer clean documents over revision-marked/tracked changes
    if any(k in low for k in ("clean", "_cl", "-cl")) and "track" not in low:
        score += 10_000_000
    elif "track" in low or "rm" in low:
        score -= 1_000_000

    # 3. Date / ISO timestamp: 20261016_1400 or 20261016
    m_iso = re.search(r'(20\d{2})(\d{2})(\d{2})(?:[_\-](\d{2,4}))?', stem)
    if m_iso:
        y, m, d, hm = m_iso.groups()
        time_int = int(y) * 100000000 + int(m) * 1000000 + int(d) * 10000 + int(hm or 0)
        score += time_int

    # 4. Date dash pattern: MM-DD-YY or DD-MM-YY (e.g. 10-16-26)
    m_dash = re.search(r'(\d{2})-(\d{2})-(\d{2,4})', stem)
    if m_dash:
        p1, p2, p3 = int(m_dash.group(1)), int(m_dash.group(2)), int(m_dash.group(3))
        year = p3 + 2000 if p3 < 100 else p3
        day, month = (p1, p2) if p1 > 12 else (p2, p1)
        score += year * 100000000 + month * 1000000 + day * 10000

    # 5. Version number (e.g. v01, v02, v3, rev2, r02)
    m_ver = re.search(r'(?:[_\-\s]v|rev|r)(\d{1,2})(?:$|[_\-\s\.])', low)
    if m_ver:
        score += int(m_ver.group(1)) * 10_000

    return score


def analyze_candidate_files(
    file_tuples: List[Tuple[str, str]], local_dir: Optional[Path] = None
) -> List[Dict[str, Any]]:
    """
    Analyzes a list of candidate `(filename, url)` tuples from the server:
    - Identifies chair / track for each file
    - Calculates recency score
    - Detects whether each file is already downloaded to `local_dir`
    - Pre-selects (`is_latest_for_chair = True`) the newest file in each chair/track group

    Returns a list of dict records.
    """
    records: List[Dict[str, Any]] = []
    groups: Dict[str, List[int]] = {}

    for idx, (fname, url) in enumerate(file_tuples):
        t_label, t_id = classify_chair_track(fname)
        score = compute_recency_score(fname)
        is_dl = (local_dir / fname).exists() if local_dir else False
        local_size = (local_dir / fname).stat().st_size if is_dl else 0

        record: Dict[str, Any] = {
            "filename": fname,
            "url": url,
            "track_id": t_id,
            "track_label": t_label,
            "recency_score": score,
            "is_latest_for_chair": False,
            "is_already_downloaded": is_dl,
            "local_size": local_size,
            "original_index": idx,
        }
        records.append(record)
        groups.setdefault(t_id, []).append(idx)

    # Pre-pick the highest recency file in each track group
    for t_id, indices in groups.items():
        best_idx = max(indices, key=lambda i: records[i]["recency_score"])
        records[best_idx]["is_latest_for_chair"] = True

    return records
