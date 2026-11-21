"""Shared meeting resolution, local cache lookup, TDoc normalization, and stance helpers."""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import get_cache_root
from .common import resolve_meetings_cache_roots as _get_meetings_cache_roots

logger = logging.getLogger(__name__)


def _search_meetings_compatible(
    mdb: Any,
    *,
    wg: Optional[str] = None,
    query: Optional[str] = None,
    meeting_number: Optional[str] = None,
    year: Optional[Any] = None,
    location: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Calls MeetingsDatabase.search_meetings using its canonical signature."""
    search_term = str(query or meeting_number or "").strip() or None
    date_from = None
    date_to = None
    if year is not None and str(year).strip():
        yr_str = re.sub(r"\D", "", str(year).strip())
        if len(yr_str) == 4:
            date_from = f"{yr_str}-01-01"
            date_to = f"{yr_str}-12-31"

    rows = mdb.search_meetings(
        wg_name=wg or None,
        search_term=search_term,
        location=location or None,
        date_from=date_from,
        date_to=date_to,
    )
    normalized: List[Dict[str, Any]] = []
    for r in rows or []:
        d = dict(r)
        folder = str(d.get("folder_name") or d.get("meeting_folder_name") or "")
        d["folder_name"] = folder
        d.setdefault("meeting_folder_name", folder)
        normalized.append(d)

    if meeting_number:
        clean_num = str(meeting_number).strip().lstrip("#").lower()
        exact = [
            m for m in normalized
            if str(m.get("meeting_number") or "").strip().lstrip("#").lower() == clean_num
        ]
        if exact:
            others = [m for m in normalized if m not in exact]
            return exact + others
    return normalized


def _resolve_meeting_record(
    mdb: Any,
    *,
    wg: Optional[str] = None,
    meeting_number: Optional[str] = None,
    tdoc_id: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """Resolves a single best-matching meeting record from MeetingsDatabase."""
    if tdoc_id and hasattr(mdb, "find_meeting_by_tdoc"):
        try:
            mtg = mdb.find_meeting_by_tdoc(tdoc_id)
            if mtg:
                d = dict(mtg)
                folder = str(d.get("folder_name") or d.get("meeting_folder_name") or "")
                d["folder_name"] = folder
                d.setdefault("meeting_folder_name", folder)
                return d
        except Exception:
            pass
    if meeting_number:
        rows = _search_meetings_compatible(mdb, wg=wg, meeting_number=meeting_number)
        if rows:
            return rows[0]
    if wg:
        rows = _search_meetings_compatible(mdb, wg=wg)
        if rows:
            return rows[0]
    return None


def _resolve_meeting_cache_dir(folder_name: str, wg: Optional[str] = None) -> Path:
    """Finds an existing local cache directory for a meeting folder, or returns the default target."""
    clean_folder = str(folder_name or "").strip()
    clean_wg = str(wg or "").strip().upper()
    roots = _get_meetings_cache_roots()
    if clean_folder:
        for r in roots:
            direct = r / clean_folder
            if direct.exists():
                return direct
            if clean_wg and (r / clean_wg / clean_folder).exists():
                return r / clean_wg / clean_folder
    base = roots[0] if roots else get_cache_root()
    return (base / clean_folder) if clean_folder else base


def _normalize_tdoc_record(
    raw: Dict[str, Any],
    *,
    wg_name: str = "",
    meeting_number: str = "",
    folder_name: str = "",
) -> Dict[str, Any]:
    """Normalizes a TDoc row from Excel/JSON, tdoc_metadata_cache, or SQLite tdocs into a unified dict."""
    tdoc_num = str(
        raw.get("tdoc_number")
        or raw.get("TDoc")
        or raw.get("tdoc")
        or ""
    ).strip()
    title = str(raw.get("title") or raw.get("Title") or "").strip()
    source = str(raw.get("source") or raw.get("Source") or "").strip()
    doc_type = str(raw.get("doc_type") or raw.get("Type") or raw.get("tdoc_type") or "").strip()
    doc_for = str(raw.get("doc_for") or raw.get("For") or "").strip()
    agenda_item = str(raw.get("agenda_item") or raw.get("Agenda Item") or "").strip()
    ai_desc = str(
        raw.get("agenda_item_description")
        or raw.get("Agenda item description")
        or raw.get("ai_desc")
        or ""
    ).strip()
    work_item = str(
        raw.get("work_item")
        or raw.get("Related WIs")
        or raw.get("wi")
        or ""
    ).strip()
    spec = str(raw.get("spec") or raw.get("Spec") or "").strip()
    release = str(raw.get("release") or raw.get("Release") or "").strip()
    status = str(
        raw.get("tdoc_status")
        or raw.get("TDoc Status")
        or raw.get("status")
        or ""
    ).strip()
    remarks = str(
        raw.get("secretary_remarks")
        or raw.get("Secretary Remarks")
        or raw.get("remarks")
        or ""
    ).strip()
    revision_of = str(
        raw.get("revision_of")
        or raw.get("Is revision of")
        or raw.get("Revision of")
        or ""
    ).strip()
    revised_to = str(raw.get("revised_to") or raw.get("Revised to") or "").strip()
    abstract = str(raw.get("abstract") or raw.get("Abstract") or "").strip()
    clauses_affected = str(
        raw.get("affected_clauses")
        or raw.get("Clauses Affected")
        or raw.get("Clauses affected")
        or ""
    ).strip()

    eff_wg = str(raw.get("wg_name") or wg_name or "").strip()
    eff_mtg_num = str(raw.get("meeting_number") or meeting_number or "").strip()
    eff_folder = str(
        raw.get("folder_name")
        or raw.get("meeting_folder_name")
        or folder_name
        or ""
    ).strip()

    confirmed_cos, tentative_cos = CompanySanitizer.parse_contributors_detailed(source)
    return {
        "tdoc_number": tdoc_num,
        "title": title,
        "source": source,
        "confirmed_companies": list(confirmed_cos),
        "tentative_companies": list(tentative_cos),
        "doc_type": doc_type,
        "doc_for": doc_for,
        "agenda_item": agenda_item,
        "agenda_item_description": ai_desc,
        "work_item": work_item,
        "spec": spec,
        "release": release,
        "tdoc_status": status,
        "secretary_remarks": remarks,
        "revision_of": revision_of,
        "revised_to": revised_to,
        "abstract": abstract,
        "affected_clauses": clauses_affected,
        "wg_name": eff_wg,
        "meeting_number": eff_mtg_num,
        "folder_name": eff_folder,
        "meeting_folder_name": eff_folder,
    }


def _load_tdocs_for_meeting_or_db(
    mdb: Any,
    *,
    specs_db_path: Path,
    spec_history_db_path: Optional[Path] = None,
    wg: Optional[str] = None,
    meeting_number: Optional[str] = None,
    tdoc_id: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Loads TDoc rows from the application's real storage layers:
    1. Optional `tdocs` SQLite table in `specs_db_path` (if present in test fixtures).
    2. Meeting spreadsheet caches (`Agenda/*.xlsx.json` or `Agenda/*.xlsx`) in `MeetingsSettings().cache_dir`.
    3. `tdoc_metadata_cache` in `3gpp_spec_history.db` as fallback.
    """
    from modules.meetings.core.tdocs.tdocs_parser import TDocsParser

    results: List[Dict[str, Any]] = []
    seen_tdocs: Set[str] = set()

    # 1. Check if `tdocs` table exists in specs_db_path (e.g. disposable unit-test fixtures)
    try:
        with mdb._get_connection() as conn:
            has_tdocs_tbl = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='tdocs' LIMIT 1"
            ).fetchone()
            if has_tdocs_tbl:
                sql = """
                    SELECT t.*, m.meeting_number AS m_meeting_number, m.folder_name AS m_folder_name,
                           w.name AS m_wg_name
                    FROM tdocs t
                    LEFT JOIN meetings m ON t.meeting_id = m.id
                    LEFT JOIN working_groups w ON m.wg_id = w.id
                    WHERE 1=1
                """
                params: List[Any] = []
                if tdoc_id:
                    sql += " AND UPPER(COALESCE(t.tdoc_number, '')) LIKE ?"
                    params.append(f"%{tdoc_id.upper()}%")
                rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
                for r in rows:
                    row_wg = str(r.get("wg_name") or r.get("m_wg_name") or "")
                    row_num = str(r.get("meeting_number") or r.get("m_meeting_number") or "")
                    row_folder = str(
                        r.get("folder_name")
                        or r.get("meeting_folder_name")
                        or r.get("m_folder_name")
                        or ""
                    )
                    if wg and row_wg and row_wg.upper() != wg.upper():
                        continue
                    if meeting_number and meeting_number.lower() not in row_num.lower() and meeting_number.lower() not in row_folder.lower():
                        continue
                    norm = _normalize_tdoc_record(
                        r,
                        wg_name=row_wg or (wg or ""),
                        meeting_number=row_num or (meeting_number or ""),
                        folder_name=row_folder,
                    )
                    key = f"{norm['tdoc_number'].upper()}::{norm['folder_name']}"
                    if norm["tdoc_number"] and key not in seen_tdocs:
                        seen_tdocs.add(key)
                        results.append(norm)
    except Exception as e:
        logger.debug(f"Optional SQLite tdocs check skipped: {e}")

    # 2. Resolve candidate meetings and load their cached Agenda spreadsheets (.xlsx.json / .xlsx)
    candidate_meetings: List[Dict[str, Any]] = []
    if tdoc_id and not meeting_number:
        mtg = _resolve_meeting_record(mdb, wg=wg, tdoc_id=tdoc_id)
        if mtg:
            candidate_meetings.append(mtg)

    if meeting_number or wg:
        matched_mtgs = _search_meetings_compatible(mdb, wg=wg, meeting_number=meeting_number)
        if meeting_number:
            clean_num = meeting_number.strip().lstrip("#").lower()
            exact_mtgs = [
                m for m in matched_mtgs
                if str(m.get("meeting_number") or "").strip().lstrip("#").lower() == clean_num
                or clean_num in str(m.get("folder_name") or "").lower()
            ]
            candidate_meetings.extend(exact_mtgs or matched_mtgs[:5])
        else:
            candidate_meetings.extend(matched_mtgs[:8])
    elif not candidate_meetings:
        candidate_meetings.extend(_search_meetings_compatible(mdb)[:5])

    visited_folders: Set[str] = set()
    for mtg in candidate_meetings:
        folder = str(mtg.get("folder_name") or mtg.get("meeting_folder_name") or "").strip()
        if not folder or folder in visited_folders:
            continue
        visited_folders.add(folder)
        m_wg = str(mtg.get("wg_name") or wg or "").strip()
        m_num = str(mtg.get("meeting_number") or meeting_number or "").strip()
        mtg_dir = _resolve_meeting_cache_dir(folder, wg=m_wg)
        agenda_dir = mtg_dir / "Agenda"
        if not agenda_dir.exists():
            continue

        excel_rows: List[Dict[str, Any]] = []
        json_files = sorted(
            [
                p for p in agenda_dir.glob("*.xlsx.json")
                if not p.name.endswith(".wi_checked.json")
            ]
        )
        for jf in json_files:
            try:
                data = json.loads(jf.read_text(encoding="utf-8"))
                if isinstance(data, list) and data:
                    excel_rows = data
                    break
            except Exception:
                pass

        if not excel_rows:
            xlsx_files = sorted(
                [p for p in agenda_dir.glob("*.xlsx") if not p.name.startswith("~$")]
            )
            for xf in xlsx_files:
                try:
                    parsed = TDocsParser.parse_tdocs_excel(str(xf))
                    if parsed:
                        excel_rows = parsed
                        break
                except Exception:
                    pass

        for r in excel_rows:
            norm = _normalize_tdoc_record(
                r,
                wg_name=m_wg,
                meeting_number=m_num,
                folder_name=folder,
            )
            key = f"{norm['tdoc_number'].upper()}::{folder}"
            if norm["tdoc_number"] and key not in seen_tdocs:
                seen_tdocs.add(key)
                results.append(norm)

    # 3. Fallback to tdoc_metadata_cache in 3gpp_spec_history.db if still empty
    if not results and spec_history_db_path and Path(spec_history_db_path).exists():
        try:
            from modules.specifications.core.spec_history_db import SpecHistoryDatabase
            hdb = SpecHistoryDatabase(spec_history_db_path)
            with hdb._get_connection() as conn:
                sql = "SELECT * FROM tdoc_metadata_cache WHERE 1=1"
                params = []
                if tdoc_id:
                    sql += " AND (UPPER(tdoc) LIKE ? OR UPPER(base_tdoc) LIKE ?)"
                    params.extend([f"%{tdoc_id.upper()}%", f"%{tdoc_id.upper()}%"])
                if meeting_number:
                    sql += " AND meeting LIKE ?"
                    params.append(f"%{meeting_number}%")
                sql += " LIMIT 500"
                for r in conn.execute(sql, params).fetchall():
                    norm = _normalize_tdoc_record(
                        dict(r),
                        wg_name=wg or "",
                        meeting_number=meeting_number or str(r["meeting"] or ""),
                    )
                    if norm["tdoc_number"]:
                        results.append(norm)
        except Exception:
            pass

    return results


def _collect_user_tdoc_annotations(
    user_tdocs_db_path: Path,
    folder_names: Sequence[str] = (),
    wg: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """Merges personal TDoc stances & notes from per-meeting Agenda/user_tdocs.db and global user_tdocs.db."""
    from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase

    merged: Dict[str, Dict[str, str]] = {}
    db_candidates: List[Path] = []
    for folder in folder_names:
        if not folder:
            continue
        m_dir = _resolve_meeting_cache_dir(folder, wg=wg)
        per_mtg_db = m_dir / "Agenda" / "user_tdocs.db"
        if per_mtg_db.exists() and per_mtg_db not in db_candidates:
            db_candidates.append(per_mtg_db)
    if user_tdocs_db_path and user_tdocs_db_path not in db_candidates:
        db_candidates.append(user_tdocs_db_path)

    for db_p in db_candidates:
        if not db_p.exists() and db_p != user_tdocs_db_path:
            continue
        try:
            udb = TDocsDatabase(db_p)
            for tid, d in udb.get_all_user_data().items():
                key = str(tid).strip().upper()
                if key not in merged or (d.get("my_notes") or d.get("my_status") not in ("", "⚪ Neutral")):
                    merged[key] = {
                        "my_status": str(d.get("my_status") or d.get("status") or ""),
                        "my_notes": str(d.get("my_notes") or d.get("notes") or ""),
                    }
        except Exception:
            pass
    return merged


def _matches_company_filter(
    raw_source: str,
    source_company: str,
    confirmed_only: bool = True,
) -> bool:
    """Checks whether `raw_source` matches `source_company`, respecting confirmed vs. `?`/`[...]` tentative rules."""
    if not source_company:
        return True
    if not raw_source:
        return False

    target_clean = source_company.strip()
    target_lower = target_clean.lower()
    target_canon_list = [c.lower() for c in CompanySanitizer.get_matching_contributors(target_clean, include_tentative=True)]

    confirmed_cos, tentative_cos = CompanySanitizer.parse_contributors_detailed(raw_source)
    candidate_canon = [c.lower() for c in confirmed_cos]
    if not confirmed_only:
        candidate_canon.extend(c.lower() for c in tentative_cos if c.lower() not in candidate_canon)

    if target_canon_list:
        if any(tc in candidate_canon for tc in target_canon_list):
            return True
    if any(target_lower in c for c in candidate_canon):
        return True

    # Fallback for companies not yet in SIGNATURE_SYNONYMS_REGEX: inspect tokenized segments directly
    for seg_text, is_tent in CompanySanitizer.tokenize_source_segments(raw_source):
        if confirmed_only and is_tent:
            continue
        # Strip role brackets like [Rapporteur] before checking literal substring
        seg_clean = CompanySanitizer.ROLE_BRACKET_RE.sub(" ", seg_text).lower()
        if target_lower in seg_clean:
            return True
    return False
