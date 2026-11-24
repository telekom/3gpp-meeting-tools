"""Change-history retrieval, flexible search, reverse TDoc lookup, and statistics queries."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple, Union


class SpecHistoryQueriesMixin:
    """Mixin providing change history lookup, filtered search, version summaries, and stats."""

    def get_spec_history(
        self,
        spec_number: str,
        version: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Retrieves full change history records for a specification, optionally filtered by version.
        Each record includes an embedded 'tdocs' list of all associated TDocs.
        """
        clean_spec = spec_number.strip()
        with self._get_connection() as conn:
            query = """
                SELECT id, spec_number, version, row_order, date, meeting,
                       source_tdoc, cr_number, revision, category, subject,
                       old_version, source_file, affected_clauses, extracted_at
                FROM spec_change_history
                WHERE spec_number = ?
            """
            params: List[Any] = [clean_spec]

            if version:
                clean_ver = version.lstrip("vV").strip()
                query += " AND version = ?"
                params.append(clean_ver)

            query += " ORDER BY row_order ASC, id ASC"

            rows = conn.execute(query, params).fetchall()
            results = [dict(r) for r in rows]

            if results:
                history_ids = [r["id"] for r in results]
                placeholders = ",".join("?" for _ in history_ids)
                tdoc_rows = conn.execute(f"""
                    SELECT history_id, tdoc, raw_tdoc, is_pcr
                    FROM spec_tdoc_cross_ref
                    WHERE history_id IN ({placeholders})
                    ORDER BY id ASC
                """, history_ids).fetchall()

                tdoc_map: Dict[int, List[Dict[str, Any]]] = {}
                for t in tdoc_rows:
                    hid = t["history_id"]
                    tdoc_map.setdefault(hid, []).append({
                        "tdoc": t["tdoc"],
                        "raw_tdoc": t["raw_tdoc"],
                        "is_pcr": bool(t["is_pcr"]),
                    })

                for r in results:
                    r["tdocs"] = tdoc_map.get(r["id"], [])

            return results

    # Convenience alias for get_spec_history
    get_history_for_spec = get_spec_history

    def search_history(
        self,
        spec_number: Optional[str] = None,
        query: Optional[str] = None,
        meeting: Optional[str] = None,
        category: Optional[str] = None,
        version: Optional[str] = None,
        clause_id: Optional[str] = None,
        tdoc_id: Optional[str] = None,
        limit: int = 50,
    ) -> List[Dict[str, Any]]:
        """
        Searches across spec_change_history and spec_tdoc_cross_ref with flexible filters.
        Each returned record includes 'tdocs' (list of dicts) and 'pcr_tdocs' (list of raw TDoc strings).
        """
        sql = """
            SELECT h.id, h.spec_number, h.version, h.row_order, h.date, h.meeting,
                   h.source_tdoc, h.cr_number, h.revision, h.category, h.subject,
                   h.old_version, h.source_file, h.affected_clauses, h.extracted_at
            FROM spec_change_history h
            WHERE 1=1
        """
        params: List[Any] = []

        if spec_number:
            clean_spec = str(spec_number).strip().upper()
            for prefix in ("3GPP ", "TS ", "TR "):
                if clean_spec.startswith(prefix):
                    clean_spec = clean_spec[len(prefix):].strip()
            sql += " AND h.spec_number = ?"
            params.append(clean_spec)

        if version:
            clean_ver = str(version).strip().lstrip("vV")
            sql += " AND h.version = ?"
            params.append(clean_ver)

        if meeting:
            clean_mtg = str(meeting).strip()
            sql += " AND h.meeting LIKE ?"
            params.append(f"%{clean_mtg}%")

        if category:
            clean_cat = str(category).strip().upper()
            sql += " AND UPPER(h.category) = ?"
            params.append(clean_cat)

        if clause_id:
            clean_cl = str(clause_id).strip()
            sql += " AND (h.affected_clauses LIKE ? OR h.subject LIKE ?)"
            params.extend([f"%{clean_cl}%", f"%{clean_cl}%"])

        if tdoc_id:
            clean_tdoc = str(tdoc_id).strip().upper()
            base_tdoc = clean_tdoc.split("R")[0]
            sql += """
                AND (
                    UPPER(h.source_tdoc) LIKE ?
                    OR EXISTS (
                        SELECT 1 FROM spec_tdoc_cross_ref c
                        WHERE c.history_id = h.id
                          AND (c.tdoc = ? OR UPPER(c.raw_tdoc) LIKE ?)
                    )
                )
            """
            params.extend([f"%{clean_tdoc}%", base_tdoc, f"%{clean_tdoc}%"])

        if query:
            clean_q = str(query).strip()
            like_q = f"%{clean_q}%"
            sql += """
                AND (
                    h.subject LIKE ?
                    OR h.source_tdoc LIKE ?
                    OR h.cr_number LIKE ?
                    OR h.affected_clauses LIKE ?
                    OR h.meeting LIKE ?
                    OR EXISTS (
                        SELECT 1 FROM spec_tdoc_cross_ref c
                        WHERE c.history_id = h.id
                          AND (c.tdoc LIKE ? OR c.raw_tdoc LIKE ?)
                    )
                )
            """
            params.extend([like_q, like_q, like_q, like_q, like_q, like_q, like_q])

        safe_limit = max(1, int(limit or 50))
        sql += " ORDER BY h.row_order DESC, h.id DESC LIMIT ?"
        params.append(safe_limit)

        with self._get_connection() as conn:
            rows = conn.execute(sql, params).fetchall()
            results = [dict(r) for r in rows]

            if results:
                history_ids = [r["id"] for r in results]
                placeholders = ",".join("?" for _ in history_ids)
                tdoc_rows = conn.execute(f"""
                    SELECT history_id, tdoc, raw_tdoc, is_pcr
                    FROM spec_tdoc_cross_ref
                    WHERE history_id IN ({placeholders})
                    ORDER BY id ASC
                """, history_ids).fetchall()

                tdoc_map: Dict[int, List[Dict[str, Any]]] = {}
                pcr_map: Dict[int, List[str]] = {}
                for t in tdoc_rows:
                    hid = t["history_id"]
                    is_pcr = bool(t["is_pcr"])
                    raw_t = str(t["raw_tdoc"] or t["tdoc"] or "").strip()
                    tdoc_map.setdefault(hid, []).append({
                        "tdoc": t["tdoc"],
                        "raw_tdoc": raw_t,
                        "is_pcr": is_pcr,
                    })
                    if is_pcr and raw_t:
                        pcr_map.setdefault(hid, []).append(raw_t)

                for r in results:
                    hid = r["id"]
                    r["tdocs"] = tdoc_map.get(hid, [])
                    r["pcr_tdocs"] = pcr_map.get(hid, [])

            return results

    def get_tdocs_for_spec_version(self, spec_number: str, version: str) -> List[Dict[str, Any]]:
        """Returns all distinct TDocs cross-referenced for a given specification and version."""
        clean_spec = spec_number.strip()
        clean_ver = version.lstrip("vV").strip()
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT DISTINCT tdoc, raw_tdoc, is_pcr, meeting
                FROM spec_tdoc_cross_ref
                WHERE spec_number = ? AND version = ?
                ORDER BY tdoc ASC
            """, (clean_spec, clean_ver)).fetchall()
            return [dict(r) for r in rows]

    def get_version_summary(
        self,
        spec_number: Optional[Union[str, Sequence[str]]] = None,
    ) -> Dict[Tuple[str, str], int]:
        """
        Fast bulk pre-fetch returning a dictionary of {(spec_number, version): record_count}.
        Supports a single spec number, a sequence/list of spec numbers, or None for the full database.
        """
        summary: Dict[Tuple[str, str], int] = {}
        with self._get_connection() as conn:
            if spec_number:
                if isinstance(spec_number, str):
                    rows = conn.execute("""
                        SELECT spec_number, version, COUNT(id) AS cnt
                        FROM spec_change_history
                        WHERE spec_number = ?
                        GROUP BY spec_number, version
                    """, (spec_number.strip(),)).fetchall()
                else:
                    specs = [str(s).strip() for s in spec_number if str(s).strip()]
                    if not specs:
                        return summary
                    placeholders = ",".join("?" for _ in specs)
                    rows = conn.execute(f"""
                        SELECT spec_number, version, COUNT(id) AS cnt
                        FROM spec_change_history
                        WHERE spec_number IN ({placeholders})
                        GROUP BY spec_number, version
                    """, specs).fetchall()
            else:
                rows = conn.execute("""
                    SELECT spec_number, version, COUNT(id) AS cnt
                    FROM spec_change_history
                    GROUP BY spec_number, version
                """).fetchall()

            for r in rows:
                summary[(r["spec_number"], r["version"])] = r["cnt"]

        return summary

    def get_versions_for_spec(self, spec_number: str) -> List[str]:
        """Returns distinct version numbers indexed for a specification in natural order."""
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT DISTINCT version
                FROM spec_change_history
                WHERE spec_number = ? AND version != ''
                ORDER BY row_order DESC, id DESC
            """, (spec_number.strip(),)).fetchall()
            return [r[0] for r in rows]

    def get_tdoc_references(self, tdoc: str) -> List[Dict[str, Any]]:
        """
        Reverse lookup: Finds all specifications and versions that integrated a given TDoc ID.
        Matches either normalized base ID (S2-2606265) or raw token with revision (S2-2606265r01).
        """
        clean_tdoc = tdoc.strip().upper()
        base_tdoc = clean_tdoc.split("R")[0]
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT c.id, c.tdoc, c.raw_tdoc, c.spec_number, c.version, c.meeting, c.is_pcr,
                       h.date, h.cr_number, h.revision, h.category, h.subject
                FROM spec_tdoc_cross_ref c
                JOIN spec_change_history h ON c.history_id = h.id
                WHERE c.tdoc = ? OR c.raw_tdoc = ?
                ORDER BY h.date DESC, c.version DESC
            """, (base_tdoc, clean_tdoc)).fetchall()
            return [dict(r) for r in rows]

    def has_spec_history(self, spec_number: str) -> bool:
        """Checks if any history records exist for the specification."""
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT 1 FROM spec_change_history WHERE spec_number = ? LIMIT 1",
                (spec_number.strip(),),
            ).fetchone()
            return bool(row)

    def get_statistics(self) -> Dict[str, Any]:
        """Returns aggregate metrics on indexed specifications, versions, change entries, and TDocs."""
        with self._get_connection() as conn:
            row_rec = conn.execute("SELECT COUNT(*) FROM spec_change_history;").fetchone()
            row_tdoc = conn.execute("SELECT COUNT(*) FROM spec_tdoc_cross_ref;").fetchone()
            row_specs = conn.execute("SELECT COUNT(DISTINCT spec_number) FROM spec_change_history;").fetchone()
            row_vers = conn.execute("SELECT COUNT(DISTINCT spec_number || ':' || version) FROM spec_change_history;").fetchone()
            return {
                "total_records": row_rec[0] if row_rec else 0,
                "total_tdocs": row_tdoc[0] if row_tdoc else 0,
                "unique_specs": row_specs[0] if row_specs else 0,
                "unique_versions": row_vers[0] if row_vers else 0,
            }
