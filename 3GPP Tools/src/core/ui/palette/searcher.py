"""Live SQLite + cached metadata searcher for CommandPaletteDialog."""

from __future__ import annotations

import logging
import re
import sqlite3
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from core.ui.palette.models import CommandItem, _TDOC_PATTERN

logger = logging.getLogger(__name__)


class DynamicPaletteSearcher:
    """Fast local SQLite + cached metadata provider for CommandPaletteDialog."""

    def __init__(
        self,
        *,
        db_path: Optional[Path] = None,
        protocol_db_path: Optional[Path] = None,
        on_open_spec: Optional[Callable[[str], None]] = None,
        on_open_meeting: Optional[Callable[[str], None]] = None,
        on_open_work_item: Optional[Callable[[str], None]] = None,
        on_open_protocol: Optional[Callable[[str, str], None]] = None,
    ) -> None:
        self.db_path = Path(db_path) if db_path else None
        self.protocol_db_path = Path(protocol_db_path) if protocol_db_path else None
        self.on_open_spec = on_open_spec
        self.on_open_meeting = on_open_meeting
        self.on_open_work_item = on_open_work_item
        self.on_open_protocol = on_open_protocol

        self._spec_titles_cache: Optional[List[Tuple[str, str, str]]] = None
        self._protocol_cache: Optional[List[Tuple[str, str, str]]] = None

    def invalidate_cache(self) -> None:
        self._spec_titles_cache = None
        self._protocol_cache = None

    def _ensure_spec_cache(self) -> List[Tuple[str, str, str]]:
        if self._spec_titles_cache is not None:
            return self._spec_titles_cache
        items: List[Tuple[str, str, str]] = []
        if self.db_path and self.db_path.exists():
            try:
                from modules.specifications.core.database import SpecsDatabase

                db = SpecsDatabase(self.db_path)
                titles = db.get_all_spec_titles() or {}
                # Fetch optional WG metadata cleanly in one lightweight read
                wg_map: Dict[str, str] = {}
                try:
                    with db._get_connection() as conn:
                        cur = conn.cursor()
                        cur.execute(
                            """
                            SELECT sp.number, COALESCE(wg.name, '')
                            FROM specifications sp
                            LEFT JOIN working_groups wg ON sp.primary_group_id = wg.id
                            """
                        )
                        for row in cur.fetchall():
                            s_num = str(row[0] or "").strip()
                            wg = str(row[1] or "").strip()
                            if s_num and wg:
                                wg_map[s_num] = wg
                except Exception:
                    pass
                for spec_num, title in titles.items():
                    items.append((str(spec_num), str(title or ""), wg_map.get(str(spec_num), "")))
            except Exception as exc:
                logger.debug("Palette failed loading spec title cache: %s", exc)
        self._spec_titles_cache = items
        return items

    def _ensure_protocol_cache(self) -> List[Tuple[str, str, str]]:
        if self._protocol_cache is not None:
            return self._protocol_cache
        items: List[Tuple[str, str, str]] = []
        if self.protocol_db_path and self.protocol_db_path.exists():
            try:
                conn = sqlite3.connect(str(self.protocol_db_path), timeout=1.5)
                try:
                    cur = conn.cursor()
                    cur.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name='nas_messages'"
                    )
                    if cur.fetchone():
                        cur.execute(
                            """
                            SELECT DISTINCT m.message_name, COALESCE(v.spec_type, 'NAS'), COALESCE(v.spec_number, '')
                            FROM nas_messages m
                            JOIN spec_versions v ON m.version_id = v.id
                            ORDER BY m.message_name ASC
                            """
                        )
                        for row in cur.fetchall():
                            name = str(row[0] or "").strip()
                            if name:
                                items.append((name, str(row[1] or "NAS"), str(row[2] or "")))
                finally:
                    conn.close()
            except Exception as exc:
                logger.debug("Palette failed loading protocol cache: %s", exc)
        self._protocol_cache = items
        return items

    def search(self, query: str) -> List[CommandItem]:
        q = (query or "").strip()
        if len(q) < 2:
            return []

        results: List[CommandItem] = []
        q_low = q.lower()
        clean_spec_q = re.sub(r"^(?:ts|tr)\s*", "", q_low).strip()

        # 1. Cached Specifications (up to 8 matches)
        spec_matches: List[Tuple[int, str, str, str]] = []
        for spec_num, title, wg in self._ensure_spec_cache():
            s_low = spec_num.lower()
            t_low = title.lower()
            if clean_spec_q and clean_spec_q in s_low:
                rank = 0 if s_low.startswith(clean_spec_q) else 1
            elif q_low in t_low or (wg and q_low in wg.lower()):
                rank = 2
            else:
                continue
            spec_matches.append((rank, spec_num, title, wg))

        spec_matches.sort(key=lambda x: (x[0], x[1]))
        for _rk, spec_num, title, wg in spec_matches[:8]:
            results.append(
                CommandItem(
                    id=f"spec:{spec_num}",
                    category="Specifications",
                    title=f"TS {spec_num}",
                    subtitle=title,
                    badge=wg,
                    icon_name="book-open",
                    callback=(lambda s=spec_num: self.on_open_spec(s)) if self.on_open_spec else None,
                    keywords=f"{spec_num} {title} {wg}",
                )
            )

        # 2. Live SQLite Meetings & TDoc jump (up to 6 matches)
        if self.db_path and self.db_path.exists():
            try:
                from modules.meetings.core.meetings_db import MeetingsDatabase

                mdb = MeetingsDatabase(self.db_path)
                if _TDOC_PATTERN.match(q):
                    tdoc_up = q.upper()
                    owning = mdb.find_meeting_by_tdoc(tdoc_up)
                    sub = (
                        f"{owning.get('wg_name', '')} · {owning.get('meeting_number', '')} ({owning.get('location', '')})"
                        if owning
                        else "Jump to TDoc in Meetings tab"
                    )
                    results.append(
                        CommandItem(
                            id=f"tdoc:{tdoc_up}",
                            category="Meetings & TDocs",
                            title=f"TDoc {tdoc_up}",
                            subtitle=sub,
                            badge="TDoc",
                            icon_name="file-text",
                            callback=(lambda t=tdoc_up: self.on_open_meeting(t)) if self.on_open_meeting else None,
                            keywords=tdoc_up,
                        )
                    )

                meetings = mdb.search_meetings(search_term=q) or []
                for m in meetings[:6]:
                    m_num = str(m.get("meeting_number") or m.get("name") or "").strip()
                    wg = str(m.get("wg_name") or "").strip()
                    loc = str(m.get("location") or "").strip()
                    s_date = str(m.get("start_date") or "").strip()
                    label = f"{wg} {m_num}".strip() if wg and wg not in m_num else m_num
                    if not label:
                        continue
                    sub_parts = [p for p in (loc, s_date) if p]
                    results.append(
                        CommandItem(
                            id=f"mtg:{label}",
                            category="Meetings & TDocs",
                            title=label,
                            subtitle=" · ".join(sub_parts),
                            badge=wg,
                            icon_name="calendar",
                            callback=(lambda key=m_num or label: self.on_open_meeting(key))
                            if self.on_open_meeting
                            else None,
                            keywords=f"{label} {loc} {s_date}",
                        )
                    )
            except Exception as exc:
                logger.debug("Palette failed querying meetings: %s", exc)

            # 3. Live SQLite Work Items (up to 6 matches)
            try:
                from modules.work_items.core.wi_database import WorkItemsDatabase

                wdb = WorkItemsDatabase(self.db_path)
                wis = wdb.search_work_item_suggestions(q, limit=6) or []
                for wi in wis[:6]:
                    acronym = str(wi.get("acronym") or "").strip()
                    code = str(wi.get("code") or "").strip()
                    name = str(wi.get("name") or "").strip()
                    results.append(
                        CommandItem(
                            id=f"wi:{code or acronym}",
                            category="Work Items",
                            title=f"{acronym} ({code})" if code else acronym,
                            subtitle=name,
                            badge="WI",
                            icon_name="clipboard-list",
                            callback=(lambda key=acronym or code: self.on_open_work_item(key))
                            if self.on_open_work_item
                            else None,
                            keywords=f"{acronym} {code} {name}",
                        )
                    )
            except Exception as exc:
                logger.debug("Palette failed querying work items: %s", exc)

        # 4. Cached Protocol Messages (up to 6 matches)
        proto_hits = [
            row
            for row in self._ensure_protocol_cache()
            if q_low in row[0].lower() or q_low in row[1].lower()
        ]
        for msg_name, ptype, spec_num in proto_hits[:6]:
            sub = f"{ptype} · TS {spec_num}" if spec_num else ptype
            results.append(
                CommandItem(
                    id=f"proto:{ptype}:{msg_name}",
                    category="Protocols",
                    title=msg_name,
                    subtitle=sub,
                    badge=ptype,
                    icon_name="network",
                    callback=(lambda p=ptype, m=msg_name: self.on_open_protocol(p, m))
                    if self.on_open_protocol
                    else None,
                    keywords=f"{msg_name} {ptype} {spec_num}",
                )
            )

        return results
