# --- File: src/modules/meetings/core/tdocs_db.py ---
import contextlib
import sqlite3
from pathlib import Path
from typing import List, Optional


def _normalize_tags_str(tags: str) -> str:
    """Deduplicates and normalizes comma-separated tags preserving order."""
    seen = []
    seen_lower = set()
    for part in str(tags or "").replace(";", ",").split(","):
        clean = part.strip()
        if clean and clean.lower() not in seen_lower:
            seen_lower.add(clean.lower())
            seen.append(clean)
    return ", ".join(seen)


class TDocsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        # Create the Agenda folder if it doesn't exist
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS user_tdocs (
                    tdoc_id TEXT PRIMARY KEY,
                    status TEXT DEFAULT '⚪ Neutral',
                    notes TEXT DEFAULT '',
                    tags TEXT DEFAULT ''
                )
            """)
            cols = [info[1] for info in conn.execute("PRAGMA table_info(user_tdocs)").fetchall()]
            if "tags" not in cols:
                conn.execute("ALTER TABLE user_tdocs ADD COLUMN tags TEXT DEFAULT ''")
            if "is_objected" not in cols:
                conn.execute("ALTER TABLE user_tdocs ADD COLUMN is_objected INTEGER DEFAULT 0")
            if "objected_at_rev" not in cols:
                conn.execute("ALTER TABLE user_tdocs ADD COLUMN objected_at_rev TEXT DEFAULT ''")

    def get_all(self) -> dict:
        """Returns a dictionary of {tdoc_id: {metadata}}"""
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT * FROM user_tdocs").fetchall()
            return {r['tdoc_id']: dict(r) for r in rows}

    def get_all_unique_tags(self) -> List[str]:
        """Returns a sorted list of unique non-empty tags across all annotated items."""
        all_rows = self.get_all()
        tags_map = {}
        for row in all_rows.values():
            raw = str(row.get("tags") or "")
            for part in raw.replace(";", ",").split(","):
                clean = part.strip()
                if clean and clean.lower() not in tags_map:
                    tags_map[clean.lower()] = clean
        return sorted(tags_map.values(), key=str.lower)

    def get_all_user_data(self) -> dict:
        """Returns a dictionary of {tdoc_id: {'tdoc_id': ..., 'status': ..., 'notes': ..., 'tags': ..., 'my_status': ..., 'my_notes': ..., 'my_tags': ...}}"""
        raw = self.get_all()
        enriched = {}
        for tid, row in raw.items():
            d = dict(row)
            d.setdefault("tags", "")
            d.setdefault("my_status", d.get("status", ""))
            d.setdefault("my_notes", d.get("notes", ""))
            d.setdefault("my_tags", d.get("tags", ""))
            d["is_objected"] = bool(int(d.get("is_objected", 0) or 0))
            d["objected_at_rev"] = str(d.get("objected_at_rev") or "").strip()
            enriched[tid] = d
        return enriched

    def set_tdoc_objected(self, tdoc_id: str, is_objected: bool, objected_at_rev: str = ""):
        """Marks or unmarks a TDoc as Objected/Parked (with the revision at objection time)."""
        clean_id = str(tdoc_id or "").strip().upper()
        if not clean_id:
            return
        with self._get_connection() as conn:
            conn.execute(
                """
                INSERT INTO user_tdocs (tdoc_id, status, notes, tags, is_objected, objected_at_rev)
                VALUES (?, '⚪ Neutral', '', '', ?, ?)
                ON CONFLICT(tdoc_id) DO UPDATE SET
                    is_objected = excluded.is_objected,
                    objected_at_rev = excluded.objected_at_rev
                """,
                (clean_id, 1 if is_objected else 0, str(objected_at_rev or "").strip() if is_objected else ""),
            )

    def upsert(self, tdoc_id: str, status: str, notes: str, tags: Optional[str] = None):
        with self._get_connection() as conn:
            if tags is None:
                row = conn.execute("SELECT tags FROM user_tdocs WHERE tdoc_id = ?", (tdoc_id,)).fetchone()
                final_tags = row[0] if (row and row[0]) else ""
            else:
                final_tags = _normalize_tags_str(tags)
            conn.execute("""
                INSERT INTO user_tdocs (tdoc_id, status, notes, tags)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(tdoc_id) DO UPDATE SET
                    status=excluded.status,
                    notes=excluded.notes,
                    tags=excluded.tags
            """, (tdoc_id, status, notes, final_tags))

    def update_tdoc_user_data(self, tdoc_id: str, status: str = "", notes: str = "", tags: Optional[str] = None):
        """Convenience wrapper around upsert() that preserves existing status, notes, or tags when omitted."""
        existing = self.get_all().get(tdoc_id, {})
        final_status = status if status else existing.get("status", "⚪ Neutral")
        final_notes = notes if notes else existing.get("notes", "")
        final_tags = tags if tags is not None else existing.get("tags", "")
        self.upsert(tdoc_id, final_status, final_notes, final_tags)
