# --- File: src/modules/meetings/core/tdocs_db.py ---
import contextlib
import sqlite3
from pathlib import Path

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
                    notes TEXT DEFAULT ''
                )
            """)

    def get_all(self) -> dict:
        """Returns a dictionary of {tdoc_id: {metadata}}"""
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT * FROM user_tdocs").fetchall()
            return {r['tdoc_id']: dict(r) for r in rows}

    def get_all_user_data(self) -> dict:
        """Returns a dictionary of {tdoc_id: {'tdoc_id': ..., 'status': ..., 'notes': ..., 'my_status': ..., 'my_notes': ...}}"""
        raw = self.get_all()
        enriched = {}
        for tid, row in raw.items():
            d = dict(row)
            d.setdefault("my_status", d.get("status", ""))
            d.setdefault("my_notes", d.get("notes", ""))
            enriched[tid] = d
        return enriched

    def upsert(self, tdoc_id: str, status: str, notes: str):
        with self._get_connection() as conn:
            conn.execute("""
                INSERT INTO user_tdocs (tdoc_id, status, notes)
                VALUES (?, ?, ?)
                ON CONFLICT(tdoc_id) DO UPDATE SET
                    status=excluded.status,
                    notes=excluded.notes
            """, (tdoc_id, status, notes))

    def update_tdoc_user_data(self, tdoc_id: str, status: str = "", notes: str = ""):
        """Convenience wrapper around upsert() that preserves existing status or notes when omitted."""
        existing = self.get_all().get(tdoc_id, {})
        final_status = status if status else existing.get("status", "⚪ Neutral")
        final_notes = notes if notes else existing.get("notes", "")
        self.upsert(tdoc_id, final_status, final_notes)
