"""
Query execution helpers for SpecSearchDatabase:
- FTS5 trigram substring search
- Short-query (< 3 chars) SQL LIKE fallback
- Clause-only prefix/TOC search (no keyword required)
- Normalized whitespace-insensitive SHA-1 hashing for Modified badge detection
"""

from __future__ import annotations

import hashlib
import re
import sqlite3
from typing import Any, List, Optional

import pandas as pd

_CLAUSE_PREFIX_RE = re.compile(
    r"^(?:clauses?|subclauses?|sub-clauses?|sections?)\s+",
    re.IGNORECASE,
)


def norm_content_hash(text: Optional[str]) -> str:
    """Returns a 16-char SHA-1 hex digest of whitespace-normalized clause content."""
    normalized = " ".join(str(text or "").split())
    if not normalized:
        return ""
    return hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:16]


def execute_fts_substring_search(
    conn: sqlite3.Connection,
    clean_query: str,
    version_ids: List[int],
    clause_filter: Optional[str] = None,
) -> pd.DataFrame:
    """Executes FTS5 trigram substring search across the given version IDs."""
    placeholders = ",".join("?" for _ in version_ids)
    sanitized_query = clean_query.replace('"', '""')
    escaped_query = f'"{sanitized_query}"'

    sql = f"""
        SELECT
            f.spec_number,
            f.version,
            v.release_date,
            f.clause_number,
            f.clause_title,
            f.version_id,
            f.clause_pk,
            snippet(spec_fts, 6, '<mark>', '</mark>', '...', 20) AS snippet_text,
            c.order_index,
            norm_hash(c.content) AS content_hash
        FROM spec_fts f
        JOIN spec_clauses c ON f.clause_pk = c.id
        JOIN indexed_versions v ON f.version_id = v.id
        WHERE f.version_id IN ({placeholders})
          AND spec_fts MATCH ?
    """
    params: List[Any] = list(version_ids) + [escaped_query]

    if clause_filter and clause_filter.strip():
        sql += " AND (f.clause_number LIKE ? OR f.clause_title LIKE ?)"
        pat = f"%{clause_filter.strip()}%"
        params.extend([pat, pat])

    sql += " ORDER BY c.order_index ASC"
    return pd.read_sql_query(sql, conn, params=params)


def execute_like_fallback_search(
    conn: sqlite3.Connection,
    clean_query: str,
    version_ids: List[int],
    clause_filter: Optional[str] = None,
) -> pd.DataFrame:
    """Fallback LIKE scanning for queries shorter than 3 characters."""
    placeholders = ",".join("?" for _ in version_ids)
    pattern = f"%{clean_query}%"

    sql = f"""
        SELECT
            v.spec_number,
            v.version,
            v.release_date,
            c.clause_number,
            c.clause_title,
            c.version_id,
            c.id AS clause_pk,
            substr(c.content, max(1, instr(lower(c.content), lower(?)) - 25), 80) AS snippet_text,
            c.order_index,
            norm_hash(c.content) AS content_hash
        FROM spec_clauses c
        JOIN indexed_versions v ON c.version_id = v.id
        WHERE c.version_id IN ({placeholders})
          AND (c.content LIKE ? OR c.clause_title LIKE ?)
    """
    params: List[Any] = [clean_query] + list(version_ids) + [pattern, pattern]

    if clause_filter and clause_filter.strip():
        sql += " AND (c.clause_number LIKE ? OR c.clause_title LIKE ?)"
        pat = f"%{clause_filter.strip()}%"
        params.extend([pat, pat])

    sql += " ORDER BY c.order_index ASC"
    return pd.read_sql_query(sql, conn, params=params)


def execute_clause_only_search(
    conn: sqlite3.Connection,
    version_ids: List[int],
    clause_filter: str,
) -> pd.DataFrame:
    """Browses clauses by clause number prefix or title when no keyword is entered."""
    if not version_ids:
        return pd.DataFrame()

    cf = _CLAUSE_PREFIX_RE.sub("", str(clause_filter or "").strip()).strip()
    if not cf:
        return pd.DataFrame()

    cf_exact = cf.rstrip(".")
    placeholders = ",".join("?" for _ in version_ids)

    sql = f"""
        SELECT
            v.spec_number,
            v.version,
            v.release_date,
            c.clause_number,
            c.clause_title,
            c.version_id,
            c.id AS clause_pk,
            COALESCE(NULLIF(trim(substr(c.content, 1, 120)), ''), c.clause_title, ' ') AS snippet_text,
            c.order_index,
            norm_hash(c.content) AS content_hash
        FROM spec_clauses c
        JOIN indexed_versions v ON c.version_id = v.id
        WHERE c.version_id IN ({placeholders})
          AND (c.clause_number = ? OR c.clause_number LIKE ? OR c.clause_title LIKE ?)
        ORDER BY c.order_index ASC
    """
    params: List[Any] = list(version_ids) + [cf_exact, f"{cf_exact}.%", f"%{cf}%"]
    return pd.read_sql_query(sql, conn, params=params)
