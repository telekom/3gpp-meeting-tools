"""TDoc agenda metadata cache upserts, affected-clause updates, and lookups."""

from __future__ import annotations

from typing import Any, Dict, List, Sequence


class SpecHistoryTdocMetadataMixin:
    """Mixin managing `tdoc_metadata_cache` and affected-clause enrichment updates."""

    def upsert_tdoc_metadata(self, entries: Sequence[Dict[str, Any]]) -> int:
        """
        Caches resolved meeting agenda metadata for TDocs in SQLite.
        Returns the number of rows inserted or updated.
        """
        if not entries:
            return 0

        inserted = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for r in entries:
                raw_tdoc = str(r.get("tdoc") or r.get("TDoc") or "").strip().upper()
                if not raw_tdoc or raw_tdoc == "-":
                    continue
                base_tdoc = raw_tdoc.split("R")[0]
                affected_cl = str(r.get("affected_clauses") or r.get("Clauses affected") or "").strip()
                cursor.execute("""
                    INSERT INTO tdoc_metadata_cache (
                        tdoc, base_tdoc, meeting, title, source, tdoc_type, status,
                        agenda_item, ai_desc, abstract, work_item, revision_of, revised_to, remarks, affected_clauses
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(tdoc) DO UPDATE SET
                        meeting = excluded.meeting,
                        title = excluded.title,
                        source = excluded.source,
                        tdoc_type = excluded.tdoc_type,
                        status = excluded.status,
                        agenda_item = excluded.agenda_item,
                        ai_desc = excluded.ai_desc,
                        abstract = excluded.abstract,
                        work_item = excluded.work_item,
                        revision_of = excluded.revision_of,
                        revised_to = excluded.revised_to,
                        remarks = excluded.remarks,
                        affected_clauses = CASE WHEN excluded.affected_clauses != '' THEN excluded.affected_clauses ELSE tdoc_metadata_cache.affected_clauses END,
                        cached_at = CURRENT_TIMESTAMP;
                """, (
                    raw_tdoc,
                    base_tdoc,
                    str(r.get("meeting", "")).strip(),
                    str(r.get("title") or r.get("Title") or "").strip(),
                    str(r.get("source") or r.get("Source") or "").strip(),
                    str(r.get("type") or r.get("Type") or "").strip(),
                    str(r.get("status") or r.get("TDoc Status") or "").strip().lower(),
                    str(r.get("agenda_item") or r.get("Agenda Item") or "").strip(),
                    str(r.get("ai_desc") or r.get("Agenda item description") or "").strip(),
                    str(r.get("abstract") or r.get("Abstract") or "").strip(),
                    str(r.get("wi") or r.get("Related WIs") or "").strip(),
                    str(r.get("revision_of") or r.get("Revision of") or "").strip(),
                    str(r.get("revised_to") or r.get("Revised to") or "").strip(),
                    str(r.get("remarks") or r.get("Secretary Remarks") or "").strip(),
                    affected_cl,
                ))
                inserted += 1

        return inserted

    def update_history_affected_clauses(self, history_id: int, affected_clauses: str) -> bool:
        """Updates the affected_clauses column for a specific change history entry."""
        if not history_id:
            return False
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE spec_change_history SET affected_clauses = ? WHERE id = ?",
                (affected_clauses.strip(), history_id),
            )
            return cursor.rowcount > 0

    def update_tdoc_affected_clauses(self, tdoc: str, affected_clauses: str) -> int:
        """
        Updates the affected_clauses column for a TDoc across tdoc_metadata_cache
        and any linked spec_change_history rows. Returns total rows modified.
        """
        clean_tdoc = tdoc.strip().upper()
        clean_base = clean_tdoc.split("R")[0]
        clean_clauses = affected_clauses.strip()
        if not clean_tdoc or not clean_clauses:
            return 0

        modified = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE tdoc_metadata_cache SET affected_clauses = ? WHERE tdoc = ? OR base_tdoc = ?",
                (clean_clauses, clean_tdoc, clean_base),
            )
            modified += cursor.rowcount

            cursor.execute("""
                UPDATE spec_change_history
                SET affected_clauses = ?
                WHERE id IN (
                    SELECT history_id FROM spec_tdoc_cross_ref
                    WHERE tdoc = ? OR raw_tdoc = ?
                )
            """, (clean_clauses, clean_base, clean_tdoc))
            modified += cursor.rowcount

        return modified

    def get_tdoc_metadata(self, tdocs: Sequence[str]) -> Dict[str, Dict[str, Any]]:
        """
        Fast O(1) retrieval of cached TDoc metadata.
        Returns a dict mapping uppercase TDoc token to its metadata dict.
        """
        if not tdocs:
            return {}

        clean_tokens = [t.strip().upper() for t in tdocs if t and t.strip() != "-"]
        if not clean_tokens:
            return {}

        results: Dict[str, Dict[str, Any]] = {}
        with self._get_connection() as conn:
            chunk_size = 500
            for i in range(0, len(clean_tokens), chunk_size):
                batch = clean_tokens[i:i + chunk_size]
                placeholders = ",".join("?" for _ in batch)
                rows = conn.execute(
                    f"SELECT * FROM tdoc_metadata_cache WHERE tdoc IN ({placeholders}) OR base_tdoc IN ({placeholders})",
                    batch + batch,
                ).fetchall()
                for row in rows:
                    d = dict(row)
                    results[d["tdoc"]] = d
                    if d["base_tdoc"] not in results:
                        results[d["base_tdoc"]] = d

        return results

    def get_meeting_tdoc_metadata(self, meeting_name: str) -> List[Dict[str, Any]]:
        """Retrieves all cached TDoc metadata records for a given meeting."""
        clean_mtg = meeting_name.strip()
        if not clean_mtg:
            return []

        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT * FROM tdoc_metadata_cache WHERE meeting = ? ORDER BY tdoc ASC",
                (clean_mtg,),
            ).fetchall()
            return [dict(r) for r in rows]
