"""Chairman's Notes filename filtering, HTML directory parsing, and version selection."""

import os
import re
import urllib.parse
from pathlib import Path
from typing import List, Tuple

from bs4 import BeautifulSoup

# Standard 3GPP pure TDoc ZIP pattern (e.g. R2-2601234.zip, S2-2605678r01.zip)
_PURE_TDOC_ZIP_RE = re.compile(
    r"^[A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?\.zip$",
    re.IGNORECASE,
)

# Extensions eligible for Chairman's Notes / Agenda notes
_VALID_NOTES_EXTENSIONS = (".docx", ".doc", ".zip", ".htm", ".html")

SUBPATHS = [
    "Inbox/Chairs_Notes",
    "INBOX/Chairs_Notes",
    "Inbox/Chair_Notes",
    "INBOX/Chair_Notes",
    "Inbox/Chairs_notes",
    "INBOX/Chairs_notes",
    "Inbox/Chair_notes",
    "INBOX/Chair_notes",
    "Chairs_Notes",
    "Chair_Notes",
]


def looks_like_directory_listing(html_text: str) -> bool:
    """Return True if HTTP HTML resembles a 3GPP directory listing containing notes."""
    if not html_text:
        return False
    low = html_text.lower()
    return any(
        token in low
        for token in (
            "chairnotes",
            "chair_notes",
            "chairs_notes",
            "directory listing",
            "parent directory",
            "downloadinput",
            "sort by",
            "<table",
            ".docx",
            ".doc",
            ".zip",
        )
    )


def is_valid_chair_notes_filename(filename: str) -> bool:
    """Determine whether a filename inside Chair_Notes is a valid Chairman's Notes or Agenda EOM file."""
    if not filename:
        return False
    name = filename.strip()
    if name.startswith("~$") or name.startswith("."):
        return False
    up = name.upper()
    if up in ("OLDER", "PARENT DIRECTORY", "ARCHIVES", "ARCHIVE", "OLD", "BACKUP"):
        return False

    low = name.lower()
    if not low.endswith(_VALID_NOTES_EXTENSIONS):
        return False

    if _PURE_TDOC_ZIP_RE.match(name):
        return False

    if (
        "chairnotes" in low
        or "chair_notes" in low
        or "chairs_notes" in low
        or "chair notes" in low
        or "chairs notes" in low
    ):
        return True

    keywords = (
        "chair",
        "note",
        "agenda",
        "session",
        "report",
        "eom",
        "final",
        "summary",
        "status",
        "minutes",
        "agreement",
        "conclusion",
    )
    return any(kw in low for kw in keywords)


def parse_file_list(html: str, base_url: str) -> List[Tuple[str, str]]:
    """Extract and validate ChairNotes filenames and URLs from HTML directory listings."""
    soup = BeautifulSoup(html, "html.parser")
    file_map = {}

    for inp in soup.find_all("input", class_="downloadInput"):
        val = inp.get("value", "").strip()
        if val and is_valid_chair_notes_filename(val):
            file_map[val] = f"{base_url.rstrip('/')}/{urllib.parse.quote(val)}"

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith(("?", "#", "javascript:", "mailto:")):
            continue
        parsed_name = urllib.parse.unquote(os.path.basename(href.rstrip("/")))
        link_text = a.get_text().strip()
        target_name = parsed_name if parsed_name and "." in parsed_name else link_text

        if target_name and is_valid_chair_notes_filename(target_name):
            full_url = urllib.parse.urljoin(base_url.rstrip("/") + "/", href)
            if target_name not in file_map:
                file_map[target_name] = full_url

    return list(file_map.items())


def prefer_clean_variants(file_tuples: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """Keep clean variants when both clean and tracked-changes/rm files exist for a document."""
    by_norm = {}
    for fname, url in file_tuples:
        stem = Path(fname).stem.lower()
        norm = re.sub(
            r"[_\-\s]*(?:tracked[_\-\s]*changes|track[_\-\s]*changes|rev[_\-\s]*marks|clean|cl|rm)$",
            "",
            stem,
            flags=re.IGNORECASE,
        )
        by_norm.setdefault(norm, []).append((fname, url))

    result = []
    for items in by_norm.values():
        if len(items) == 1:
            result.append(items[0])
            continue
        clean_items = [
            (f, u)
            for f, u in items
            if re.search(r"(?:^|[_\-\s])(clean|cl)(?:$|[_\-\s\.])", Path(f).stem.lower())
            and "track" not in f.lower()
        ]
        non_tracked = [
            (f, u)
            for f, u in items
            if not re.search(r"(?:tracked|track[_\-\s]*changes|[_\-]rm$)", Path(f).stem.lower())
        ]
        if clean_items:
            result.extend(clean_items)
        elif non_tracked:
            result.extend(non_tracked)
        else:
            result.append(items[-1])
    return result


def extract_timestamp_sort_key(filename: str) -> Tuple:
    """Extract a chronological sort key from a filename if a date/time is present."""
    stem = Path(filename).stem
    m_iso = re.search(r"(20\d{2})(\d{2})(\d{2})(?:[_\-](\d{2,4}))?", stem)
    if m_iso:
        y, m, d, hm = m_iso.groups()
        return (int(y), int(m), int(d), int(hm or 0), filename.lower())

    m_dash = re.search(r"(\d{2})-(\d{2})-(\d{2,4})", stem)
    if m_dash:
        p1, p2, p3 = int(m_dash.group(1)), int(m_dash.group(2)), int(m_dash.group(3))
        year = p3 + 2000 if p3 < 100 else p3
        if p1 > 12:
            day, month = p1, p2
        else:
            month, day = p1, p2
        return (year, month, day, 0, filename.lower())

    return (0, 0, 0, 0, filename.lower())


def pick_latest_file(items: List[Tuple[str, str]]) -> Tuple[str, str]:
    """Return the most recent `(filename, url)` entry by filename timestamp."""
    return max(items, key=lambda item: extract_timestamp_sort_key(item[0]))


def group_and_pick_latest_by_track(items: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """Group files by logical track/author and return the latest file per track."""
    if not items:
        return []
    tracks = {}
    for fname, url in items:
        stem = Path(fname).stem.lower()
        track_key = re.sub(r"20\d{6}(?:[_\-]\d{2,4})?", "", stem)
        track_key = re.sub(r"\d{2}-\d{2}-\d{2,4}", "", track_key)
        track_key = re.sub(r"[_\-\s]+(?:final|eom|clean|cl|rm|v\d+|\d{4})(?=$|[_\-\s])", "", track_key)
        track_key = re.sub(r"[_\-\s]+", "_", track_key).strip("_") or "main"
        tracks.setdefault(track_key, []).append((fname, url))

    return [pick_latest_file(group) for group in tracks.values()]


def select_best_chair_notes_files(
    file_tuples: List[Tuple[str, str]],
    only_final_or_latest: bool = True,
) -> List[Tuple[str, str]]:
    """Select authoritative Final / EOM Chairman's Notes or latest track snapshots."""
    if not only_final_or_latest or len(file_tuples) <= 1:
        return list(file_tuples)

    deduped = prefer_clean_variants(file_tuples)
    combined_finals = []
    session_finals = []

    for fname, url in deduped:
        low = fname.lower()
        stem = Path(fname).stem.lower()
        is_final_or_eom = bool(
            re.search(r"(?:^|[_\-\s\.])(final|eom|end[_\-\s]?of[_\-\s]?meeting|closing)(?:$|[_\-\s\.])", stem)
            or stem.endswith(("final", "eom", "_cl", "_clean"))
            and ("final" in stem or "eom" in stem)
        )
        is_session_or_track = bool(
            "session" in low
            or re.search(r"\([a-z\s]+\)", low)
            or re.match(r"^chair_?notes_[a-z]+", low)
        )

        if is_final_or_eom:
            if is_session_or_track:
                session_finals.append((fname, url))
            else:
                combined_finals.append((fname, url))

    if combined_finals:
        best_combined = [pick_latest_file(combined_finals)]
        best_sessions = group_and_pick_latest_by_track(session_finals)
        return best_combined + best_sessions

    if session_finals:
        return group_and_pick_latest_by_track(session_finals)

    return group_and_pick_latest_by_track(deduped)
