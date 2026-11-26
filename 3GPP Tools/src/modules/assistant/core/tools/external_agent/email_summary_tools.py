"""MCP handlers for reading, computing deltas, and saving persistent email summaries."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from modules.emails.core.ai_summary import (
    EmailSummaryRecord,
    EmailSummaryStore,
    extract_clean_direct_body,
    get_message_timestamp_str,
    parse_summary_output,
    sort_messages_chronologically,
)
from modules.emails.core.emeeting_signals import extract_base_tdoc_id
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

from .common import resolve_meetings_cache_roots as _get_meetings_cache_roots
from .meeting_helpers import _resolve_meeting_record


def _resolve_meeting_agenda_dir(
    *,
    specs_db_path: Path,
    wg: Optional[str] = None,
    meeting_number: Optional[str] = None,
    meeting_hint: Optional[str] = None,
    scope_key: str = "",
    scope_type: str = "",
) -> Optional[Path]:
    """Resolve `<meeting_dir>/Agenda` for email_summaries.db lookup."""
    hint = str(meeting_hint or "").strip()
    if hint:
        p = Path(hint)
        if p.is_dir():
            return (p / "Agenda") if (p / "Agenda").is_dir() or p.name != "Agenda" else p
        for root in _get_meetings_cache_roots():
            cand = root / hint
            if cand.is_dir():
                return cand / "Agenda"

    if specs_db_path and specs_db_path.is_file():
        try:
            from modules.meetings.core.meetings_db import MeetingsDatabase

            mdb = MeetingsDatabase(specs_db_path)
            tdoc_arg = scope_key if scope_type == "tdoc" else None
            rec = _resolve_meeting_record(
                mdb,
                wg=wg,
                meeting_number=meeting_number or hint or None,
                tdoc_id=tdoc_arg,
            )
            if rec and rec.get("folder_name"):
                folder = str(rec["folder_name"]).strip()
                clean_wg = str(rec.get("wg_name") or wg or "").strip().upper()
                roots = _get_meetings_cache_roots()
                for root in roots:
                    if (root / folder).exists():
                        return root / folder / "Agenda"
                    if clean_wg and (root / clean_wg / folder).exists():
                        return root / clean_wg / folder / "Agenda"
                if roots:
                    return roots[0] / folder / "Agenda"
        except Exception:
            pass

    for root in _get_meetings_cache_roots():
        if not root.is_dir():
            continue
        for child in sorted(root.iterdir()):
            if child.is_dir() and (child / "Agenda").is_dir():
                if not meeting_number or str(meeting_number) in child.name:
                    return child / "Agenda"
    return None


def _collect_scope_messages(
    agenda_dir: Path,
    scope_type: str,
    scope_key: str,
) -> List[Dict[str, Any]]:
    """Collect deduplicated chronological messages for `thread` or `tdoc` from local meeting DBs."""
    by_id: Dict[str, Dict[str, Any]] = {}
    tdb_path = agenda_dir / "tracked_threads.db"
    gdb_path = agenda_dir / "emails.db"

    if scope_type == "thread" and tdb_path.is_file():
        tdb = TrackedThreadsDatabase(tdb_path)
        for m in tdb.get_thread_messages(scope_key):
            mid = str(m.get("id") or "").strip()
            if mid:
                by_id[mid] = dict(m)
    elif scope_type == "tdoc":
        base_tdoc = extract_base_tdoc_id(scope_key)
        if gdb_path.is_file():
            gdb = GeneralEmailDatabase(gdb_path)
            for m in gdb.get_emails_for_tdocs({base_tdoc}, show_ignored=False):
                mid = str(m.get("id") or "").strip()
                if mid:
                    by_id[mid] = dict(m)
        if tdb_path.is_file():
            tdb = TrackedThreadsDatabase(tdb_path)
            for th in tdb.get_threads():
                linked = {extract_base_tdoc_id(t) for t in (th.get("linked_tdocs") or [])}
                if base_tdoc in linked:
                    for m in tdb.get_thread_messages(th["thread_id"]):
                        mid = str(m.get("id") or "").strip()
                        if mid and mid not in by_id:
                            by_id[mid] = dict(m)

    return sort_messages_chronologically(list(by_id.values()))


def handle_email_summary_mcp_action(
    action: str,
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
) -> Dict[str, Any]:
    """Dispatch `list_email_summaries`, `get_email_summary`, and `save_email_summary`."""
    scope_type = str(args.get("scope_type") or "").strip().lower()
    scope_key = str(args.get("scope_key") or args.get("query") or "").strip()
    wg = str(args.get("wg") or "").strip().upper() or None
    meeting_number = str(args.get("meeting_number") or "").strip() or None
    meeting_hint = str(args.get("meeting") or "").strip() or None

    agenda_dir = _resolve_meeting_agenda_dir(
        specs_db_path=specs_db_path,
        wg=wg,
        meeting_number=meeting_number,
        meeting_hint=meeting_hint,
        scope_key=scope_key,
        scope_type=scope_type,
    )
    if agenda_dir is None:
        return {"ok": False, "error": "Could not resolve target meeting cache directory."}

    store = EmailSummaryStore(agenda_dir / "email_summaries.db")

    if action == "list_email_summaries":
        recs = store.list_summaries(scope_type=scope_type or None)
        limit = max(1, int(args.get("limit") or 50))
        return {
            "ok": True,
            "db_path": str(store.db_path),
            "count": len(recs[:limit]),
            "summaries": [
                {
                    "scope_type": r.scope_type,
                    "scope_key": r.scope_key,
                    "title": r.title,
                    "cutoff_utc": r.cutoff_utc,
                    "message_count": r.message_count,
                    "incremental_hops": r.incremental_hops,
                    "provider": r.provider,
                    "model": r.model,
                    "reasoning_effort": r.reasoning_effort,
                    "updated_at": r.updated_at,
                }
                for r in recs[:limit]
            ],
        }

    if not scope_type or not scope_key:
        return {"ok": False, "error": f"scope_type and scope_key are required for action='{action}'."}

    if action == "get_email_summary":
        rec = store.get_summary(scope_type, scope_key)
        include_delta = bool(args.get("include_delta") or args.get("include_delta_messages"))
        delta_payload: List[Dict[str, Any]] = []
        freshness_info: Dict[str, Any] = {}
        if include_delta and scope_type in ("thread", "tdoc"):
            msgs = _collect_scope_messages(agenda_dir, scope_type, scope_key)
            fresh = store.compute_freshness(scope_type, scope_key, msgs)
            new_set = set(fresh.new_message_ids)
            freshness_info = {
                "total_messages": fresh.total_messages,
                "covered_count": fresh.covered_count,
                "new_count": fresh.new_count,
                "recommend_full_rebuild": fresh.recommend_full_rebuild,
            }
            for m in msgs:
                mid = str(m.get("id") or "").strip()
                if mid in new_set:
                    delta_payload.append({
                        "id": mid,
                        "sender_name": m.get("sender_name", ""),
                        "company": m.get("company", ""),
                        "date": get_message_timestamp_str(m),
                        "subject": m.get("subject", ""),
                        "body": extract_clean_direct_body(m)[:4000],
                    })

        return {
            "ok": True,
            "found": rec is not None,
            "summary": (
                {
                    "scope_type": rec.scope_type,
                    "scope_key": rec.scope_key,
                    "title": rec.title,
                    "summary_md": rec.summary_md,
                    "delta_md": rec.delta_md,
                    "structured_meta": rec.structured_meta,
                    "cutoff_utc": rec.cutoff_utc,
                    "covered_message_ids": rec.covered_message_ids,
                    "message_count": rec.message_count,
                    "incremental_hops": rec.incremental_hops,
                    "provider": rec.provider,
                    "model": rec.model,
                    "reasoning_effort": rec.reasoning_effort,
                    "updated_at": rec.updated_at,
                }
                if rec is not None
                else None
            ),
            "freshness": freshness_info,
            "delta_messages": delta_payload,
        }

    if action == "save_email_summary":
        raw_md = str(args.get("summary_md") or "").strip()
        if not raw_md:
            return {"ok": False, "error": "summary_md is required for action='save_email_summary'."}
        clean_md, parsed_delta, structured_meta = parse_summary_output(raw_md)
        delta_md = str(args.get("delta_md") or parsed_delta or "").strip()
        raw_cov = args.get("covered_ids") if args.get("covered_ids") is not None else args.get("covered_message_ids")
        if isinstance(raw_cov, str):
            covered_ids = [x.strip() for x in raw_cov.split(",") if x.strip()]
        else:
            covered_ids = [str(x).strip() for x in (raw_cov or []) if str(x).strip()]
        cutoff_utc = str(args.get("cutoff_utc") or "").strip()
        existing = store.get_summary(scope_type, scope_key)
        if existing and not covered_ids:
            covered_ids = list(existing.covered_message_ids)
        if not cutoff_utc and scope_type in ("thread", "tdoc"):
            msgs = _collect_scope_messages(agenda_dir, scope_type, scope_key)
            if msgs:
                if not covered_ids:
                    covered_ids = [str(m.get("id") or "").strip() for m in msgs if m.get("id")]
                cutoff_utc = get_message_timestamp_str(msgs[-1])

        saved = store.upsert_summary(
            EmailSummaryRecord(
                scope_type=scope_type,
                scope_key=scope_key,
                title=str(args.get("title") or (existing.title if existing else scope_key)),
                summary_md=clean_md,
                delta_md=delta_md,
                structured_meta=structured_meta,
                cutoff_utc=cutoff_utc,
                covered_message_ids=covered_ids,
                message_count=len(covered_ids),
                incremental_hops=int(args.get("incremental_hops") or (existing.incremental_hops + 1 if existing else 0)),
                provider=str(args.get("provider") or "mcp").strip(),
                model=str(args.get("model") or "external-agent").strip(),
                reasoning_effort=str(args.get("reasoning_effort") or "").strip(),
            )
        )
        return {
            "ok": True,
            "scope_type": saved.scope_type,
            "scope_key": saved.scope_key,
            "cutoff_utc": saved.cutoff_utc,
            "message_count": saved.message_count,
            "updated_at": saved.updated_at,
        }

    return {"ok": False, "error": f"Unsupported email summary action '{action}'."}
