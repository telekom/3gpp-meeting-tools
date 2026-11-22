# --- File: src/modules/emails/core/stats/dataset_builder.py ---
"""
Builds a unified, deduplicated email analytics dataset for `EmailStatsExporterThread`
across `GeneralEmailDatabase`, `TrackedThreadsDatabase`, and `TDocsWindow` metadata.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

from modules.emails.core.emeeting_signals import extract_base_tdoc_id
from modules.emails.core.thread_markdown_exporter import clean_tdoc_id


def build_unified_email_stats_dataset(
    *,
    all_inbox_items: Optional[Sequence[Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
    tracked_db: Optional[Any] = None,
    meeting_id: str = "",
    tdoc_ai_lookup: Optional[Dict[str, str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """
    Produces a deduplicated list of email message records normalized for `EmailStatsExporterThread`:
    `id`, `tdoc_id`, `agenda_item`, `sender_name`, `sender_email`, `company`, `date_received`, `subject`.
    """
    ai_lookup = {str(k).upper(): str(v).strip() for k, v in (tdoc_ai_lookup or {}).items() if str(k).strip()}
    meta_map = tdoc_meta_map or {}

    def _resolve_ais(tdocs: Sequence[str], explicit_ais: Sequence[str]) -> str:
        ais: List[str] = [str(a).strip() for a in (explicit_ais or ()) if str(a).strip()]
        for raw_td in (tdocs or ()):
            base_td = extract_base_tdoc_id(clean_tdoc_id(raw_td))
            if not base_td:
                continue
            ai = ai_lookup.get(base_td) or str((meta_map.get(base_td) or {}).get("ai") or "").strip()
            if ai and ai not in ais:
                ais.append(ai)
        return ", ".join(ais) if ais else "UNASSIGNED"

    if all_inbox_items is not None:
        out: List[Dict[str, Any]] = []
        for it in all_inbox_items:
            eid = str(it.get("id") or "").strip()
            if not eid:
                continue
            tdocs = [clean_tdoc_id(t) for t in (it.get("linked_tdocs") or []) if clean_tdoc_id(t)]
            ais_str = _resolve_ais(tdocs, it.get("linked_ais") or [])
            s_name = str(it.get("sender_name") or "").strip() or "Unknown"
            s_email = str(it.get("sender_email") or "").strip() or s_name.lower()
            out.append({
                "id": eid,
                "tdoc_id": tdocs[0] if tdocs else "",
                "agenda_item": ais_str,
                "sender_name": s_name,
                "sender_email": s_email,
                "company": str(it.get("company") or "").strip(),
                "date_received": str(it.get("date_str") or it.get("date_received") or "").strip(),
                "subject": str(it.get("subject") or "").strip(),
            })
        return out

    by_id: Dict[str, Dict[str, Any]] = {}
    if tracked_db is not None and hasattr(tracked_db, "get_threads_for_meeting"):
        for th in tracked_db.get_threads_for_meeting(meeting_id):
            tid = str(th.get("thread_id") or "")
            th_ais = [str(a).strip() for a in (th.get("linked_ais") or []) if str(a).strip()]
            th_tdocs = [clean_tdoc_id(t) for t in (th.get("linked_tdocs") or []) if clean_tdoc_id(t)]
            for m in tracked_db.get_thread_messages(tid):
                eid = str(m.get("id") or "").strip()
                if not eid:
                    continue
                s_name = str(m.get("sender_name") or "").strip() or "Unknown"
                s_email = str(m.get("sender_email") or "").strip() or s_name.lower()
                by_id[eid] = {
                    "id": eid,
                    "tdoc_id": th_tdocs[0] if th_tdocs else "",
                    "agenda_item": _resolve_ais(th_tdocs, th_ais),
                    "sender_name": s_name,
                    "sender_email": s_email,
                    "company": str(m.get("company") or "").strip(),
                    "date_received": str(m.get("date_sent") or "").strip(),
                    "subject": str(m.get("subject") or th.get("title") or "").strip(),
                    "_tdocs": list(th_tdocs),
                    "_ais": list(th_ais),
                }

    if general_email_db is not None and hasattr(general_email_db, "get_all_emails_with_tdocs"):
        for row in general_email_db.get_all_emails_with_tdocs(show_ignored=False):
            eid = str(row.get("id") or "").strip()
            if not eid:
                continue
            matched_tdocs = list(dict.fromkeys(
                clean_tdoc_id(t)
                for t in str(row.get("matched_tdocs") or "").split(",")
                if clean_tdoc_id(t)
            ))
            if eid in by_id:
                existing = by_id[eid]
                merged_tdocs = list(dict.fromkeys((existing.get("_tdocs") or []) + matched_tdocs))
                existing["_tdocs"] = merged_tdocs
                if not existing["tdoc_id"] and merged_tdocs:
                    existing["tdoc_id"] = merged_tdocs[0]
                existing["agenda_item"] = _resolve_ais(merged_tdocs, existing.get("_ais") or [])
            else:
                s_name = str(row.get("sender_name") or "").strip() or "Unknown"
                s_email = str(row.get("sender_email") or "").strip() or s_name.lower()
                by_id[eid] = {
                    "id": eid,
                    "tdoc_id": matched_tdocs[0] if matched_tdocs else "",
                    "agenda_item": _resolve_ais(matched_tdocs, []),
                    "sender_name": s_name,
                    "sender_email": s_email,
                    "company": str(row.get("company") or row.get("sender_company") or "").strip(),
                    "date_received": str(row.get("date_received") or "").strip(),
                    "subject": str(row.get("subject") or "").strip(),
                }

    return list(by_id.values())
