# --- File: src/modules/nas/core/db/queries_ies.py ---
"""IE definition listing, evolution DataFrame, clause lookup/disambiguation, and cross-spec queries."""

import hashlib
import re
from typing import Any, Dict, List, Optional

import pandas as pd

from .schema import (
    RE_NAS_IE_EXACT_LEN,
    RE_NAS_IE_MIN_MAX_LEN,
    RE_NAS_IE_TYPE,
    parse_version_tuple,
)


class NasDbIeQueriesMixin:
    """Mixin querying Clause 9 / ASN.1 Information Element definitions and cross-spec references."""

    def get_ie_definitions_list(
        self,
        version_ids: Optional[List[int]] = None,
        ie_query: str = "",
        search_descriptions: bool = False,
    ) -> List[Dict[str, Any]]:
        """
        Returns all distinct Information Element / Type definitions across selected versions,
        excluding legacy TOC page-number artifacts and Void/structural stubs.
        """
        if not version_ids:
            return []

        placeholders = ",".join("?" for _ in version_ids)
        where_parts = [
            f"d.version_id IN ({placeholders})",
            "NOT (d.ie_name GLOB '* [0-9][0-9][0-9]*' AND LENGTH(COALESCE(d.raw_description, '')) < 250)",
            "LOWER(TRIM(d.ie_name)) NOT IN ('void', 'general', 'overview', 'other information elements', 'information elements coding')",
        ]
        params: List[Any] = list(version_ids)

        clean_query = ie_query.strip()
        if clean_query:
            if search_descriptions and len(clean_query) >= 2:
                col_expr = "(COALESCE(d.ie_name, '') || ' ' || COALESCE(d.clause, '') || ' ' || COALESCE(d.raw_description, ''))"
            else:
                col_expr = "(COALESCE(d.ie_name, '') || ' ' || COALESCE(d.clause, ''))"
            tok_sql, tok_params = self._build_tokenized_like_clause(col_expr, clean_query)
            where_parts.append(tok_sql)
            params.extend(tok_params)

        where_sql = " AND ".join(where_parts)
        query = f"""
            SELECT d.ie_name, d.clause,
                   GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number,
                   GROUP_CONCAT(DISTINCT sv.version) AS version
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE {where_sql}
            GROUP BY d.ie_name, d.clause
            ORDER BY d.ie_name ASC, d.clause ASC
        """

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_ie_definition_evolution_df(
        self,
        clause: str,
        ie_name: str,
        version_ids: List[int],
    ) -> pd.DataFrame:
        """
        Builds a release-evolution DataFrame compatible with NASEvolutionMatrixModel
        for a standalone Clause 9 / ASN.1 / PFCP Information Element definition.
        """
        if not version_ids:
            return pd.DataFrame()

        defs = self.get_ie_definitions_by_clause(
            clause=clause,
            alt_name=ie_name,
            version_ids=version_ids,
        )
        if not defs:
            return pd.DataFrame()

        rows: List[Dict[str, Any]] = []
        for d in defs:
            raw_html = d.get("raw_description", "") or ""
            plain_text = re.sub(r"<[^>]+>", " ", raw_html)
            plain_text = " ".join(plain_text.split())

            ie_type_match = RE_NAS_IE_TYPE.search(plain_text)
            fmt_str = f"Type {ie_type_match.group(1)}" if ie_type_match else f"Clause {d.get('clause', clause)}"

            min_max_match = RE_NAS_IE_MIN_MAX_LEN.search(plain_text)
            exact_len_match = RE_NAS_IE_EXACT_LEN.search(plain_text)
            if min_max_match:
                min_l = min_max_match.group(1)
                max_l = min_max_match.group(2)
                len_str = f"{min_l}-{max_l} octets" if max_l else f"Min {min_l} octets"
            elif exact_len_match:
                len_str = f"{exact_len_match.group(1)} octets"
            else:
                digest = hashlib.md5(plain_text.encode("utf-8", errors="ignore")).hexdigest()[:6]
                len_str = f"Rev #{digest}"

            resolved_clause = d.get("clause", "") or clause
            resolved_name = d.get("ie_name", "") or ie_name
            rows.append({
                "spec_number": d.get("spec_number", ""),
                "version": d.get("version", ""),
                "iei": resolved_clause if resolved_clause != resolved_name else "-",
                "ie_name": resolved_name,
                "field_path": resolved_name,
                "depth": 0,
                "type_reference": f"{resolved_name} {resolved_clause}".strip() if resolved_clause != resolved_name else resolved_clause,
                "presence": "Defined",
                "format": fmt_str,
                "length": len_str,
                "order_index": 0,
                "applicability": "",
                "ie_description": raw_html,
            })

        return pd.DataFrame(rows)

    def get_ie_definitions_by_clause(
        self,
        clause: str,
        alt_name: str = "",
        spec_number: Optional[str] = None,
        version_ids: Optional[List[int]] = None,
    ) -> List[Dict[str, Any]]:
        params = []
        where_parts = []

        clean_c = clause.strip() if clause else ""
        clean_alt = alt_name.strip() if alt_name else ""

        if clean_c:
            where_parts.append("(LOWER(TRIM(d.clause)) = LOWER(?) OR LOWER(TRIM(d.ie_name)) = LOWER(?))")
            params.extend([clean_c, clean_c])

        if clean_alt:
            where_parts.append("(LOWER(TRIM(d.clause)) = LOWER(?) OR LOWER(TRIM(d.ie_name)) = LOWER(?))")
            params.extend([clean_alt, clean_alt])

        if not where_parts:
            return []

        where_sql = " OR ".join(where_parts)
        query = f"""
            SELECT d.*, sv.version, sv.spec_number
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE ({where_sql})
              AND NOT (d.ie_name GLOB '* [0-9][0-9][0-9]*' AND LENGTH(COALESCE(d.raw_description, '')) < 250)
        """

        if spec_number:
            query += " AND sv.spec_number = ?"
            params.append(spec_number.strip())

        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" AND d.version_id IN ({placeholders})"
            params.extend(version_ids)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = [dict(row) for row in cursor.fetchall()]

        # Disambiguate umbrella clauses (e.g. TS 38.331 Clause 6.3.2 shared by hundreds of IEs)
        if rows:
            if clean_alt:
                exact_alt_rows = [
                    r for r in rows
                    if str(r.get("ie_name", "")).strip().lower() == clean_alt.lower()
                ]
                if exact_alt_rows:
                    rows = exact_alt_rows
                elif clean_c:
                    exact_c_rows = [
                        r for r in rows
                        if str(r.get("ie_name", "")).strip().lower() == clean_c.lower()
                    ]
                    if exact_c_rows:
                        rows = exact_c_rows
                    else:
                        names_per_ver: Dict[Any, set] = {}
                        for r in rows:
                            names_per_ver.setdefault(r.get("version_id"), set()).add(
                                str(r.get("ie_name", "")).strip().lower()
                            )
                        rows = [
                            r for r in rows
                            if len(names_per_ver.get(r.get("version_id"), set())) <= 3
                        ]
            elif clean_c:
                exact_c_rows = [
                    r for r in rows
                    if str(r.get("ie_name", "")).strip().lower() == clean_c.lower()
                ]
                if exact_c_rows:
                    rows = exact_c_rows

        return sorted(rows, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

    def get_latest_spec_version_by_major(
        self,
        spec_number: str,
        major_version: Optional[int] = None,
    ) -> Optional[Dict[str, Any]]:
        query = "SELECT id, spec_number, version, spec_type, import_date FROM spec_versions WHERE spec_number = ?"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number.strip(),))
            rows = [dict(r) for r in cursor.fetchall()]

        if not rows:
            return None

        if major_version is not None:
            def _extract_major(ver_str: str) -> Optional[int]:
                vt = parse_version_tuple(ver_str)
                if not vt:
                    return None
                first = vt[0]
                return first[1] if isinstance(first, tuple) and first[0] == 0 else (first if isinstance(first, int) else None)

            rows = [
                r for r in rows
                if _extract_major(r["version"]) == major_version
            ]

        if not rows:
            return None

        rows.sort(key=lambda r: parse_version_tuple(r["version"]), reverse=True)
        return rows[0]

    def get_cross_referenced_ie_definition(
        self,
        target_spec: str,
        target_clause: str,
        alt_name: str = "",
        major_version: Optional[int] = None,
    ) -> Optional[Dict[str, Any]]:
        latest_ver = self.get_latest_spec_version_by_major(target_spec, major_version)
        if not latest_ver:
            return None

        clean_c = target_clause.strip().rstrip(".") if target_clause else ""
        clean_alt = alt_name.strip() if alt_name else ""

        where_parts = []
        params: List[Any] = [latest_ver["id"]]

        if clean_c:
            where_parts.append("LOWER(TRIM(d.clause)) = LOWER(?)")
            params.append(clean_c)
        if clean_alt:
            where_parts.append("LOWER(TRIM(d.ie_name)) = LOWER(?)")
            params.append(clean_alt)

        if clean_c:
            where_parts.append("d.clause LIKE ?")
            params.append(f"%{clean_c}%")

        if not where_parts:
            return None

        where_sql = " OR ".join(where_parts)
        params.extend([clean_c, clean_alt])

        query = f"""
            SELECT d.*, sv.version, sv.spec_number
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE sv.id = ? AND ({where_sql})
              AND NOT (d.ie_name GLOB '* [0-9][0-9][0-9]*' AND LENGTH(COALESCE(d.raw_description, '')) < 250)
            ORDER BY
                CASE
                    WHEN LOWER(TRIM(d.clause)) = LOWER(?) THEN 1
                    WHEN LOWER(TRIM(d.ie_name)) = LOWER(?) THEN 2
                    ELSE 3
                END ASC,
                LENGTH(COALESCE(d.raw_description, '')) DESC
            LIMIT 1
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return dict(row) if row else None
