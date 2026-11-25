# --- File: src/modules/nas/core/db/queries_messages.py ---
"""Message listing, tokenized CTE search, evolution DataFrame, and message context queries."""

import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .schema import RE_TRAILING_ASN1_CAPTION, parse_version_tuple


class NasDbMessageQueriesMixin:
    """Mixin implementing protocol message search, evolution matrix, and context queries."""

    @staticmethod
    def _build_tokenized_like_clause(column_expr: str, query_str: str) -> Tuple[str, List[str]]:
        """
        Builds an AND-chained LIKE clause across whitespace-separated tokens so queries like
        'UE parameters transparent container' match 'UE parameters update transparent container'.
        """
        tokens = [t.strip() for t in re.split(r"\s+", query_str.strip()) if t.strip()]
        if not tokens:
            return "1=1", []
        clauses = [f"{column_expr} LIKE ?" for _ in tokens]
        params = [f"%{t}%" for t in tokens]
        return "(" + " AND ".join(clauses) + ")", params

    def get_messages_list(self, version_ids: Optional[List[int]] = None) -> List[Dict[str, Any]]:
        query = """
            SELECT m.message_name, m.clause,
                   MAX(COALESCE(m.description, '')) AS description,
                   MAX(COALESCE(m.direction, '')) AS direction,
                   GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number,
                   GROUP_CONCAT(DISTINCT sv.version) AS version
            FROM nas_messages m
            JOIN spec_versions sv ON m.version_id = sv.id
        """
        params = []
        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" WHERE m.version_id IN ({placeholders})"
            params.extend(version_ids)
        query += " GROUP BY m.message_name, m.clause ORDER BY m.message_name ASC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_by_ie_search(
        self,
        ie_query: str,
        version_ids: Optional[List[int]] = None,
        search_descriptions: bool = False,
    ) -> List[Dict[str, Any]]:
        """
        Searches messages matching an IE name, field path, or clause 9 / ASN.1 description.
        Supports multi-word tokenized queries and Common Table Expressions (CTE).
        """
        if not version_ids or not ie_query.strip():
            return self.get_messages_list(version_ids)

        clean_query = ie_query.strip()
        placeholders = ",".join("?" for _ in version_ids)

        ie_col_expr = "(COALESCE(i.ie_name, '') || ' ' || COALESCE(i.field_path, '') || ' ' || COALESCE(i.type_reference, '') || ' ' || COALESCE(i.iei, ''))"
        ie_tok_sql, ie_tok_params = self._build_tokenized_like_clause(ie_col_expr, clean_query)

        if search_descriptions and len(clean_query) >= 2:
            def_col_expr = "(COALESCE(d.ie_name, '') || ' ' || COALESCE(d.clause, '') || ' ' || COALESCE(d.raw_description, ''))"
            def_tok_sql, def_tok_params = self._build_tokenized_like_clause(def_col_expr, clean_query)

            msg_col_expr = "(COALESCE(m.message_name, '') || ' ' || COALESCE(m.clause, '') || ' ' || COALESCE(m.description, '') || ' ' || COALESCE(m.direction, ''))"
            msg_tok_sql, msg_tok_params = self._build_tokenized_like_clause(msg_col_expr, clean_query)

            query = f"""
                WITH matched_defs AS (
                    SELECT version_id, ie_name, clause
                    FROM ie_definitions d
                    WHERE d.version_id IN ({placeholders})
                      AND {def_tok_sql}
                ),
                matching_messages AS (
                    -- 1. Direct Information Element / ASN.1 Field Name Matches
                    SELECT m.id AS message_id, m.version_id, m.message_name, m.clause,
                           COALESCE(m.description, '') AS description,
                           COALESCE(m.direction, '') AS direction
                    FROM nas_messages m
                    JOIN message_ies i ON i.message_id = m.id
                    WHERE m.version_id IN ({placeholders})
                      AND {ie_tok_sql}

                    UNION

                    -- 2. Direct Message Description / Direction Matches
                    SELECT m.id AS message_id, m.version_id, m.message_name, m.clause,
                           COALESCE(m.description, '') AS description,
                           COALESCE(m.direction, '') AS direction
                    FROM nas_messages m
                    WHERE m.version_id IN ({placeholders})
                      AND {msg_tok_sql}

                    UNION

                    -- 3. Matches via Associated Clause 9 or ASN.1 Field / Message Definitions
                    SELECT m.id AS message_id, m.version_id, m.message_name, m.clause,
                           COALESCE(m.description, '') AS description,
                           COALESCE(m.direction, '') AS direction
                    FROM nas_messages m
                    LEFT JOIN message_ies i ON i.message_id = m.id
                    JOIN matched_defs d ON d.version_id = m.version_id
                        AND (
                            m.message_name = d.ie_name
                            OR (d.clause != '' AND m.clause = d.clause)
                            OR i.ie_name = d.ie_name
                            OR i.type_reference = d.ie_name
                            OR i.type_reference = d.clause
                            OR (d.clause != '' AND (i.type_reference = d.clause OR i.type_reference = (d.ie_name || ' ' || d.clause)))
                            OR (length(d.ie_name) >= 4 AND instr(i.type_reference, d.ie_name) > 0)
                        )
                    WHERE m.version_id IN ({placeholders})
                )
                SELECT mm.message_name, mm.clause,
                       MAX(mm.description) AS description,
                       MAX(mm.direction) AS direction,
                       GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
                FROM matching_messages mm
                JOIN spec_versions sv ON mm.version_id = sv.id
                GROUP BY mm.message_name, mm.clause
                ORDER BY mm.message_name ASC
            """
            params = (
                list(version_ids) + def_tok_params
                + list(version_ids) + ie_tok_params
                + list(version_ids) + msg_tok_params
                + list(version_ids)
            )
        else:
            query = f"""
                SELECT m.message_name, m.clause,
                       MAX(COALESCE(m.description, '')) AS description,
                       MAX(COALESCE(m.direction, '')) AS direction,
                       GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
                FROM nas_messages m
                JOIN message_ies i ON i.message_id = m.id
                JOIN spec_versions sv ON m.version_id = sv.id
                WHERE m.version_id IN ({placeholders})
                  AND {ie_tok_sql}
                GROUP BY m.message_name, m.clause
                ORDER BY m.message_name ASC
            """
            params = list(version_ids) + ie_tok_params

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_using_ie(
        self,
        clause: Optional[str] = None,
        ie_name: Optional[str] = None,
        spec_number: Optional[str] = None,
        version_ids: Optional[List[int]] = None,
    ) -> List[Dict[str, Any]]:
        if not version_ids:
            return []

        placeholders = ",".join("?" for _ in version_ids)
        where_conditions = [f"m.version_id IN ({placeholders})"]
        params: List[Any] = list(version_ids)

        if spec_number:
            where_conditions.append("sv.spec_number = ?")
            params.append(spec_number.strip())

        clean_clause = clause.strip() if clause else ""
        clean_name = ie_name.strip() if ie_name else ""

        if clean_clause and re.match(r"^(?:9|8|7|6|5|4|D\.6)(?:\.[0-9A-Za-z]+)+$", clean_clause):
            regex_pattern = rf"(?<![0-9A-Za-z.]){re.escape(clean_clause)}(?![0-9A-Za-z.])"
            where_conditions.append("i.type_reference REGEXP ?")
            params.append(regex_pattern)
        elif clean_name or clean_clause:
            target = clean_name or clean_clause
            where_conditions.append("(LOWER(TRIM(i.ie_name)) = LOWER(?) OR i.type_reference LIKE ?)")
            params.extend([target, f"%{target}%"])
        else:
            return []

        where_clause = " AND ".join(where_conditions)
        query = f"""
            SELECT m.message_name, m.clause, GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
            FROM nas_messages m
            JOIN message_ies i ON i.message_id = m.id
            JOIN spec_versions sv ON m.version_id = sv.id
            WHERE {where_clause}
            GROUP BY m.message_name, m.clause
            ORDER BY m.message_name ASC
        """

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_message_evolution_df(
        self,
        message_name: str,
        version_ids: List[int],
        include_descriptions: bool = False,
    ) -> pd.DataFrame:
        if not version_ids:
            return pd.DataFrame()

        placeholders = ",".join("?" for _ in version_ids)
        if include_descriptions:
            # Separate indexed LEFT JOINs avoid non-indexed Cartesian OR loops
            query = f"""
                SELECT
                    sv.spec_number,
                    sv.version,
                    i.iei,
                    i.ie_name,
                    i.field_path,
                    i.depth,
                    i.type_reference,
                    i.presence,
                    i.format,
                    i.length,
                    i.order_index,
                    i.applicability,
                    i.field_description,
                    COALESCE(NULLIF(i.field_description, ''), d1.raw_description, d2.raw_description, d3.raw_description) AS ie_description
                FROM message_ies i
                JOIN nas_messages m ON i.message_id = m.id
                JOIN spec_versions sv ON m.version_id = sv.id
                LEFT JOIN ie_definitions d1 ON d1.version_id = sv.id AND d1.ie_name = i.ie_name AND d1.ie_name != ''
                    AND (d1.clause = m.message_name OR d1.clause = m.clause OR d1.clause GLOB '[0-9]*' OR d1.clause GLOB 'Annex*' OR d1.clause = '')
                LEFT JOIN ie_definitions d2 ON d2.version_id = sv.id AND d2.ie_name = i.type_reference AND d2.ie_name != ''
                LEFT JOIN ie_definitions d3 ON d3.version_id = sv.id AND d3.clause != '' AND (d3.clause = i.type_reference OR (d3.ie_name || ' ' || d3.clause) = i.type_reference)
                WHERE m.message_name = ? AND sv.id IN ({placeholders})
                ORDER BY i.order_index ASC
            """
        else:
            query = f"""
                SELECT
                    sv.spec_number,
                    sv.version,
                    i.iei,
                    i.ie_name,
                    i.field_path,
                    i.depth,
                    i.type_reference,
                    i.presence,
                    i.format,
                    i.length,
                    i.order_index,
                    i.applicability,
                    i.field_description
                FROM message_ies i
                JOIN nas_messages m ON i.message_id = m.id
                JOIN spec_versions sv ON m.version_id = sv.id
                WHERE m.message_name = ? AND sv.id IN ({placeholders})
                ORDER BY i.order_index ASC
            """

        params = [message_name] + version_ids

        with self._get_connection() as conn:
            return pd.read_sql_query(query, conn, params=params)

    def get_message_context_info(
        self,
        message_name: str,
        clause: str = "",
        version_ids: Optional[List[int]] = None,
    ) -> Dict[str, str]:
        """
        Resolves the richest message description and direction across selected specification versions,
        falling back to extracting prose/direction from ie_definitions if nas_messages columns are empty
        (e.g. for databases ingested prior to message description tracking).
        """
        params: List[Any] = [message_name]
        ver_filter = ""
        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            ver_filter = f" AND m.version_id IN ({placeholders})"
            params.extend(version_ids)

        query = f"""
            SELECT m.clause, m.message_name, m.table_caption,
                   COALESCE(m.description, '') AS description,
                   COALESCE(m.direction, '') AS direction,
                   sv.spec_number, sv.version
            FROM nas_messages m
            JOIN spec_versions sv ON m.version_id = sv.id
            WHERE m.message_name = ?{ver_filter}
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = [dict(r) for r in cursor.fetchall()]

        rows.sort(key=lambda r: parse_version_tuple(r.get("version", "")), reverse=True)

        best_desc = ""
        best_dir = ""
        best_clause = clause
        best_caption = ""

        for r in rows:
            if not best_clause and r.get("clause"):
                best_clause = r["clause"]
            if not best_caption and r.get("table_caption"):
                best_caption = r["table_caption"]
            if not best_desc and r.get("description", "").strip():
                best_desc = r["description"].strip()
            if not best_dir and r.get("direction", "").strip():
                best_dir = r["direction"].strip()
            if best_desc and best_dir:
                break

        if not best_desc or not best_dir:
            defs = self.get_ie_definitions_by_clause(
                clause=best_clause or message_name,
                alt_name=message_name,
                version_ids=version_ids,
            )
            for d in defs:
                if (
                    message_name
                    and str(d.get("ie_name", "")).strip().lower() != message_name.strip().lower()
                    and str(d.get("clause", "")).strip().lower() != message_name.strip().lower()
                ):
                    continue
                raw_html = d.get("raw_description", "") or ""
                if not raw_html:
                    continue
                before_tbl = raw_html.split("<table", 1)[0]
                before_tbl = re.sub(r"<pre[^>]*>.*?</pre>", "", before_tbl, flags=re.DOTALL | re.IGNORECASE)
                p_matches = re.findall(r"<(?:p|div)[^>]*>(.*?)</(?:p|div)>", before_tbl, flags=re.DOTALL | re.IGNORECASE)
                prose_parts = []
                for pm in p_matches:
                    clean_p = re.sub(r"<[^>]+>", "", pm).strip()
                    clean_p = re.sub(r"\s+", " ", clean_p)
                    if not clean_p:
                        continue
                    low_p = clean_p.lower()
                    if low_p.startswith(("direction:", "direction :")):
                        if not best_dir:
                            best_dir = re.sub(r"(?i)^direction\s*:\s*", "", clean_p).strip()
                    elif not low_p.startswith(("note", "table ", "figure ", "asn.1 definition", "specification details")):
                        prose_parts.append(clean_p)
                if not best_desc and prose_parts:
                    best_desc = " ".join(prose_parts).strip()
                if best_desc and best_dir:
                    break

        if best_desc:
            if message_name:
                best_desc = re.sub(
                    rf"(?:\s+|^){re.escape(message_name)}\s+(?:information\s+elements?|messages?)\s*$",
                    "",
                    best_desc,
                    flags=re.IGNORECASE,
                ).strip()
            best_desc = RE_TRAILING_ASN1_CAPTION.sub("", best_desc).strip()

        return {
            "description": best_desc,
            "direction": best_dir,
            "clause": best_clause,
            "table_caption": best_caption,
        }
