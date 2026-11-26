"""SQLite persistence (`email_summaries.db`) for multi-scope incremental AI summaries."""

from __future__ import annotations

import contextlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

from .models import EmailSummaryRecord, SummaryFreshness, VALID_SCOPE_TYPES


def utc_now_iso() -> str:
    """Return current UTC timestamp in ISO-8601 format (`YYYY-MM-DDTHH:MM:SSZ`)."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _extract_msg_id(msg: Dict[str, Any]) -> str:
    return str(msg.get("id") or msg.get("email_id") or "").strip()


class EmailSummaryStore:
    """Per-meeting SQLite store (`<meeting_dir>/Agenda/email_summaries.db`) with history undo."""

    MAX_HISTORY_PER_SCOPE = 5

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self) -> None:
        with self._get_connection() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS email_summaries (
                    scope_type TEXT NOT NULL,
                    scope_key TEXT NOT NULL,
                    title TEXT DEFAULT '',
                    summary_md TEXT NOT NULL,
                    delta_md TEXT DEFAULT '',
                    structured_json TEXT DEFAULT '{}',
                    deterministic_facts_json TEXT DEFAULT '{}',
                    cutoff_utc TEXT DEFAULT '',
                    covered_message_ids TEXT DEFAULT '[]',
                    message_count INTEGER DEFAULT 0,
                    incremental_hops INTEGER DEFAULT 0,
                    provider TEXT DEFAULT '',
                    model TEXT DEFAULT '',
                    reasoning_effort TEXT DEFAULT '',
                    focus_prompt TEXT DEFAULT '',
                    created_at TEXT DEFAULT '',
                    updated_at TEXT DEFAULT '',
                    PRIMARY KEY (scope_type, scope_key)
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS email_summary_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scope_type TEXT NOT NULL,
                    scope_key TEXT NOT NULL,
                    title TEXT DEFAULT '',
                    summary_md TEXT NOT NULL,
                    delta_md TEXT DEFAULT '',
                    structured_json TEXT DEFAULT '{}',
                    deterministic_facts_json TEXT DEFAULT '{}',
                    cutoff_utc TEXT DEFAULT '',
                    covered_message_ids TEXT DEFAULT '[]',
                    message_count INTEGER DEFAULT 0,
                    incremental_hops INTEGER DEFAULT 0,
                    provider TEXT DEFAULT '',
                    model TEXT DEFAULT '',
                    reasoning_effort TEXT DEFAULT '',
                    focus_prompt TEXT DEFAULT '',
                    saved_at TEXT DEFAULT ''
                )
                """
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_es_type ON email_summaries(scope_type, updated_at DESC)"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_esh_scope ON email_summary_history(scope_type, scope_key, id DESC)"
            )
            conn.commit()

    @staticmethod
    def _parse_json_obj(raw: Any, default: Any) -> Any:
        if isinstance(raw, type(default)):
            return raw
        try:
            parsed = json.loads(str(raw or ""))
            return parsed if isinstance(parsed, type(default)) else default
        except Exception:
            return default

    @classmethod
    def _row_to_record(cls, row: sqlite3.Row) -> EmailSummaryRecord:
        d = dict(row)
        return EmailSummaryRecord(
            scope_type=str(d.get("scope_type") or ""),
            scope_key=str(d.get("scope_key") or ""),
            title=str(d.get("title") or ""),
            summary_md=str(d.get("summary_md") or ""),
            delta_md=str(d.get("delta_md") or ""),
            structured_meta=cls._parse_json_obj(d.get("structured_json"), {}),
            deterministic_facts=cls._parse_json_obj(d.get("deterministic_facts_json"), {}),
            cutoff_utc=str(d.get("cutoff_utc") or ""),
            covered_message_ids=[
                str(x) for x in cls._parse_json_obj(d.get("covered_message_ids"), []) if str(x).strip()
            ],
            message_count=int(d.get("message_count") or 0),
            incremental_hops=int(d.get("incremental_hops") or 0),
            provider=str(d.get("provider") or ""),
            model=str(d.get("model") or ""),
            reasoning_effort=str(d.get("reasoning_effort") or ""),
            focus_prompt=str(d.get("focus_prompt") or ""),
            created_at=str(d.get("created_at") or ""),
            updated_at=str(d.get("updated_at") or ""),
        )

    @staticmethod
    def _norm_key(scope_type: str, scope_key: str) -> tuple[str, str]:
        st = (scope_type or "").strip().lower()
        sk = str(scope_key or "").strip()
        if st == "tdoc":
            sk = sk.upper()
        return st, sk

    def get_summary(self, scope_type: str, scope_key: str) -> Optional[EmailSummaryRecord]:
        st, sk = self._norm_key(scope_type, scope_key)
        if not st or not sk:
            return None
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM email_summaries WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            ).fetchone()
            return self._row_to_record(row) if row else None

    def list_summaries(self, scope_type: Optional[str] = None) -> List[EmailSummaryRecord]:
        with self._get_connection() as conn:
            if scope_type:
                st = scope_type.strip().lower()
                rows = conn.execute(
                    "SELECT * FROM email_summaries WHERE scope_type = ? ORDER BY updated_at DESC",
                    (st,),
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM email_summaries ORDER BY scope_type, updated_at DESC"
                ).fetchall()
            return [self._row_to_record(r) for r in rows]

    def has_history(self, scope_type: str, scope_key: str) -> bool:
        st, sk = self._norm_key(scope_type, scope_key)
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT 1 FROM email_summary_history WHERE scope_type = ? AND scope_key = ? LIMIT 1",
                (st, sk),
            ).fetchone()
            return row is not None

    def upsert_summary(self, record: EmailSummaryRecord, *, save_history: bool = True) -> EmailSummaryRecord:
        st, sk = self._norm_key(record.scope_type, record.scope_key)
        if st not in VALID_SCOPE_TYPES:
            raise ValueError(f"Unsupported scope_type: {record.scope_type}")
        if not sk:
            raise ValueError("scope_key must not be empty")

        now_iso = utc_now_iso()
        dedup_ids = list(dict.fromkeys(str(x).strip() for x in (record.covered_message_ids or []) if str(x).strip()))
        msg_cnt = max(int(record.message_count or 0), len(dedup_ids))

        with self._get_connection() as conn:
            existing = conn.execute(
                "SELECT * FROM email_summaries WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            ).fetchone()
            created_at = str(existing["created_at"]) if existing and existing["created_at"] else (record.created_at or now_iso)

            if existing and save_history:
                conn.execute(
                    """
                    INSERT INTO email_summary_history (
                        scope_type, scope_key, title, summary_md, delta_md,
                        structured_json, deterministic_facts_json, cutoff_utc,
                        covered_message_ids, message_count, incremental_hops,
                        provider, model, reasoning_effort, focus_prompt, saved_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        existing["scope_type"],
                        existing["scope_key"],
                        existing["title"],
                        existing["summary_md"],
                        existing["delta_md"],
                        existing["structured_json"],
                        existing["deterministic_facts_json"],
                        existing["cutoff_utc"],
                        existing["covered_message_ids"],
                        existing["message_count"],
                        existing["incremental_hops"],
                        existing["provider"],
                        existing["model"],
                        existing["reasoning_effort"],
                        existing["focus_prompt"],
                        existing["updated_at"] or now_iso,
                    ),
                )
                conn.execute(
                    """
                    DELETE FROM email_summary_history
                    WHERE scope_type = ? AND scope_key = ?
                      AND id NOT IN (
                          SELECT id FROM email_summary_history
                          WHERE scope_type = ? AND scope_key = ?
                          ORDER BY id DESC LIMIT ?
                      )
                    """,
                    (st, sk, st, sk, self.MAX_HISTORY_PER_SCOPE),
                )

            updated_at = record.updated_at or now_iso
            conn.execute(
                """
                INSERT INTO email_summaries (
                    scope_type, scope_key, title, summary_md, delta_md,
                    structured_json, deterministic_facts_json, cutoff_utc,
                    covered_message_ids, message_count, incremental_hops,
                    provider, model, reasoning_effort, focus_prompt,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(scope_type, scope_key) DO UPDATE SET
                    title = excluded.title,
                    summary_md = excluded.summary_md,
                    delta_md = excluded.delta_md,
                    structured_json = excluded.structured_json,
                    deterministic_facts_json = excluded.deterministic_facts_json,
                    cutoff_utc = excluded.cutoff_utc,
                    covered_message_ids = excluded.covered_message_ids,
                    message_count = excluded.message_count,
                    incremental_hops = excluded.incremental_hops,
                    provider = excluded.provider,
                    model = excluded.model,
                    reasoning_effort = excluded.reasoning_effort,
                    focus_prompt = excluded.focus_prompt,
                    updated_at = excluded.updated_at
                """,
                (
                    st,
                    sk,
                    record.title or sk,
                    record.summary_md,
                    record.delta_md or "",
                    json.dumps(record.structured_meta or {}, ensure_ascii=False),
                    json.dumps(record.deterministic_facts or {}, ensure_ascii=False),
                    record.cutoff_utc or "",
                    json.dumps(dedup_ids, ensure_ascii=False),
                    msg_cnt,
                    int(record.incremental_hops or 0),
                    record.provider or "",
                    record.model or "",
                    record.reasoning_effort or "",
                    record.focus_prompt or "",
                    created_at,
                    updated_at,
                ),
            )
            conn.commit()
            saved = conn.execute(
                "SELECT * FROM email_summaries WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            ).fetchone()
            return self._row_to_record(saved)

    def revert_previous_version(self, scope_type: str, scope_key: str) -> Optional[EmailSummaryRecord]:
        """Restore the most recent history snapshot for `(scope_type, scope_key)`."""
        st, sk = self._norm_key(scope_type, scope_key)
        with self._get_connection() as conn:
            hist = conn.execute(
                """
                SELECT * FROM email_summary_history
                WHERE scope_type = ? AND scope_key = ?
                ORDER BY id DESC LIMIT 1
                """,
                (st, sk),
            ).fetchone()
            if not hist:
                return None
            now_iso = utc_now_iso()
            conn.execute(
                """
                INSERT INTO email_summaries (
                    scope_type, scope_key, title, summary_md, delta_md,
                    structured_json, deterministic_facts_json, cutoff_utc,
                    covered_message_ids, message_count, incremental_hops,
                    provider, model, reasoning_effort, focus_prompt,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(scope_type, scope_key) DO UPDATE SET
                    title = excluded.title,
                    summary_md = excluded.summary_md,
                    delta_md = excluded.delta_md,
                    structured_json = excluded.structured_json,
                    deterministic_facts_json = excluded.deterministic_facts_json,
                    cutoff_utc = excluded.cutoff_utc,
                    covered_message_ids = excluded.covered_message_ids,
                    message_count = excluded.message_count,
                    incremental_hops = excluded.incremental_hops,
                    provider = excluded.provider,
                    model = excluded.model,
                    reasoning_effort = excluded.reasoning_effort,
                    focus_prompt = excluded.focus_prompt,
                    updated_at = excluded.updated_at
                """,
                (
                    st,
                    sk,
                    hist["title"],
                    hist["summary_md"],
                    hist["delta_md"],
                    hist["structured_json"],
                    hist["deterministic_facts_json"],
                    hist["cutoff_utc"],
                    hist["covered_message_ids"],
                    hist["message_count"],
                    hist["incremental_hops"],
                    hist["provider"],
                    hist["model"],
                    hist["reasoning_effort"],
                    hist["focus_prompt"],
                    hist["saved_at"] or now_iso,
                    now_iso,
                ),
            )
            conn.execute("DELETE FROM email_summary_history WHERE id = ?", (hist["id"],))
            conn.commit()
            restored = conn.execute(
                "SELECT * FROM email_summaries WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            ).fetchone()
            return self._row_to_record(restored) if restored else None

    def delete_summary(self, scope_type: str, scope_key: str) -> bool:
        st, sk = self._norm_key(scope_type, scope_key)
        with self._get_connection() as conn:
            cur = conn.execute(
                "DELETE FROM email_summaries WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            )
            conn.execute(
                "DELETE FROM email_summary_history WHERE scope_type = ? AND scope_key = ?",
                (st, sk),
            )
            conn.commit()
            return cur.rowcount > 0

    @staticmethod
    def _build_freshness(
        record: Optional[EmailSummaryRecord],
        current_messages: Sequence[Dict[str, Any]],
    ) -> SummaryFreshness:
        current_ids = [_extract_msg_id(m) for m in current_messages if _extract_msg_id(m)]
        unique_current = list(dict.fromkeys(current_ids))
        if record is None:
            return SummaryFreshness(
                has_summary=False,
                total_messages=len(unique_current),
                covered_count=0,
                new_count=len(unique_current),
                new_message_ids=unique_current,
            )
        covered_set = set(record.covered_message_ids)
        new_ids = [mid for mid in unique_current if mid not in covered_set]
        covered_cnt = len(covered_set) if covered_set else int(record.message_count or 0)
        recommend_rebuild = (record.incremental_hops >= 5) or (covered_cnt > 0 and len(new_ids) > covered_cnt)
        return SummaryFreshness(
            has_summary=True,
            total_messages=len(unique_current),
            covered_count=covered_cnt,
            new_count=len(new_ids),
            new_message_ids=new_ids,
            cutoff_utc=record.cutoff_utc,
            incremental_hops=record.incremental_hops,
            recommend_full_rebuild=recommend_rebuild,
            provider=record.provider,
            model=record.model,
            reasoning_effort=record.reasoning_effort,
            updated_at=record.updated_at,
        )

    def compute_freshness(
        self,
        scope_type: str,
        scope_key: str,
        current_messages: Sequence[Dict[str, Any]],
    ) -> SummaryFreshness:
        rec = self.get_summary(scope_type, scope_key)
        return self._build_freshness(rec, current_messages)

    def get_freshness_map(
        self,
        scope_type: str,
        scope_messages_map: Dict[str, Sequence[Dict[str, Any]]],
    ) -> Dict[str, SummaryFreshness]:
        st = (scope_type or "").strip().lower()
        recs = {
            r.scope_key.upper() if st == "tdoc" else r.scope_key: r
            for r in self.list_summaries(scope_type=st)
        }
        out: Dict[str, SummaryFreshness] = {}
        for key, msgs in scope_messages_map.items():
            lookup_key = key.upper() if st == "tdoc" else key
            out[key] = self._build_freshness(recs.get(lookup_key), msgs)
        return out
