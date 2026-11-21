"""External agent handlers for Meetings, TDocs, Chair's Notes, and SA2 Session Plans."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.company_sanitizer import CompanySanitizer
from .common import _parse_csv_list
from .meeting_helpers import (
    _collect_user_tdoc_annotations,
    _load_tdocs_for_meeting_or_db,
    _matches_company_filter,
    _resolve_meeting_cache_dir,
    _resolve_meeting_record,
    _search_meetings_compatible,
)


# =========================================================================
# DOMAIN 6: Meetings, TDocs, Triage Stances & Dossiers
# =========================================================================

def _handle_meetings_and_tdocs(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
    user_tdocs_db_path: Path,
    spec_history_db_path: Optional[Path] = None,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search_tdocs").strip().lower()
    wg = str(args.get("wg") or "").strip().upper() or None
    meeting_number = str(args.get("meeting_number") or "").strip()
    tdoc_id = str(args.get("tdoc_id") or "").strip()
    query = str(args.get("query") or "").strip()
    limit = max(1, int(args.get("limit") or 25))

    from modules.meetings.core.meetings_db import MeetingsDatabase
    from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase

    mdb = MeetingsDatabase(specs_db_path)

    if action == "search_meetings":
        rows = _search_meetings_compatible(
            mdb,
            wg=wg,
            query=query or None,
            meeting_number=meeting_number or None,
            year=args.get("year"),
            location=str(args.get("location") or "").strip() or None,
        )
        return {
            "count": len(rows[:limit]),
            "meetings": rows[:limit],
        }

    if action == "search_tdocs":
        source_company = str(args.get("source_company") or "").strip()
        confirmed_only = bool(args.get("confirmed_company_only", True))
        agenda_filter = str(args.get("agenda_item") or "").strip().lower()
        wi_filter = str(args.get("work_item") or "").strip().lower()
        status_filter = str(args.get("status") or "").strip().lower()
        query_lower = query.lower()
        tdoc_upper = tdoc_id.upper()

        all_rows = _load_tdocs_for_meeting_or_db(
            mdb,
            specs_db_path=specs_db_path,
            spec_history_db_path=spec_history_db_path,
            wg=wg,
            meeting_number=meeting_number or None,
            tdoc_id=tdoc_id or None,
        )

        filtered: List[Dict[str, Any]] = []
        for r in all_rows:
            if tdoc_upper and tdoc_upper not in str(r.get("tdoc_number") or "").upper():
                continue
            if wg and r.get("wg_name") and str(r.get("wg_name")).upper() != wg:
                continue
            if meeting_number:
                m_num = str(r.get("meeting_number") or "").lower()
                m_fol = str(r.get("folder_name") or "").lower()
                if meeting_number.lower() not in m_num and meeting_number.lower() not in m_fol:
                    continue
            if agenda_filter and agenda_filter not in str(r.get("agenda_item") or "").lower():
                continue
            if wi_filter and wi_filter not in str(r.get("work_item") or "").lower():
                continue
            if status_filter and status_filter not in str(r.get("tdoc_status") or "").lower():
                continue
            if query_lower:
                haystack = " ".join(
                    str(r.get(k) or "")
                    for k in ("tdoc_number", "title", "source", "work_item", "agenda_item", "agenda_item_description", "spec", "secretary_remarks")
                ).lower()
                if query_lower not in haystack:
                    continue
            if source_company and not _matches_company_filter(
                str(r.get("source") or ""),
                source_company=source_company,
                confirmed_only=confirmed_only,
            ):
                continue
            filtered.append(r)

        selected = filtered[:limit]
        folders = list({str(r.get("folder_name") or "") for r in selected if r.get("folder_name")})
        user_map = _collect_user_tdoc_annotations(user_tdocs_db_path, folder_names=folders, wg=wg)
        for r in selected:
            tn = str(r.get("tdoc_number") or "").strip().upper()
            base_tn = tn.split("R")[0]
            u_entry = user_map.get(tn) or user_map.get(base_tn)
            if u_entry:
                r["my_status"] = u_entry.get("my_status", "")
                r["my_notes"] = u_entry.get("my_notes", "")
            else:
                r.setdefault("my_status", "")
                r.setdefault("my_notes", "")

        return {
            "count": len(selected),
            "total_matched": len(filtered),
            "tdocs": selected,
        }

    if action == "read_tdoc_markdown":
        if not tdoc_id:
            return {"error": "tdoc_id is required for action='read_tdoc_markdown'."}
        from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        matched_rows = _load_tdocs_for_meeting_or_db(
            mdb,
            specs_db_path=specs_db_path,
            spec_history_db_path=spec_history_db_path,
            wg=wg,
            meeting_number=meeting_number or None,
            tdoc_id=tdoc_id,
        )
        exact_meta = next(
            (r for r in matched_rows if str(r.get("tdoc_number") or "").upper() == tdoc_id.upper()),
            matched_rows[0] if matched_rows else {},
        )
        mtg = _resolve_meeting_record(
            mdb,
            wg=wg or str(exact_meta.get("wg_name") or "") or None,
            meeting_number=meeting_number or str(exact_meta.get("meeting_number") or "") or None,
            tdoc_id=tdoc_id,
        ) or {}

        eff_wg = wg or str(exact_meta.get("wg_name") or mtg.get("wg_name") or "SA2")
        mtg_folder = str(
            exact_meta.get("folder_name")
            or mtg.get("folder_name")
            or mtg.get("meeting_folder_name")
            or f"{eff_wg}_{meeting_number or 'Cache'}"
        )
        docs_ftp = str(mtg.get("docs_folder_url") or "")
        rev_url = str(mtg.get("revisions_url") or "")
        meeting_dir = _resolve_meeting_cache_dir(mtg_folder, wg=eff_wg)

        # Check if already extracted on disk before attempting network download
        clean_tdoc = tdoc_id.strip().upper()
        base_tdoc = clean_tdoc.split("R")[0]
        existing_docs: List[Path] = []
        for candidate_dir in (
            meeting_dir / clean_tdoc,
            meeting_dir / base_tdoc,
            meeting_dir / "Docs" / clean_tdoc,
            meeting_dir / "Docs" / base_tdoc,
        ):
            if candidate_dir.exists():
                existing_docs.extend(
                    sorted(
                        p for p in candidate_dir.rglob("*")
                        if p.is_file() and p.suffix.lower() in (".docx", ".doc", ".md") and not p.name.startswith("~$")
                    )
                )
        files: List[Path] = existing_docs
        err = None
        if not files:
            files, err = TDocFileHandler.download_and_extract_tdoc(
                tdoc_id=tdoc_id,
                docs_ftp_url=docs_ftp,
                meeting_dir=meeting_dir,
                wg_name=eff_wg,
                revisions_url=rev_url,
            )
        if err or not files:
            return {"error": err or f"Could not locate or download TDoc {tdoc_id}."}

        existing_md = next((f for f in files if f.suffix.lower() == ".md"), None)
        target_doc = next((f for f in files if f.suffix.lower() in (".docx", ".doc")), files[0])
        rev_mode = str(args.get("revisions") or "markup").strip().lower()

        if target_doc.suffix.lower() in (".docx", ".doc"):
            if target_doc.suffix.lower() == ".doc":
                from modules.word_tools.core.word_converter import convert_doc_to_docx
                conv = convert_doc_to_docx(str(target_doc))
                if conv and Path(conv).exists():
                    target_doc = Path(conv)
            md_path = target_doc.with_suffix(".md")
            convert_document_to_markdown(target_doc, md_path, revisions=rev_mode)
        elif existing_md is not None:
            md_path = existing_md
        else:
            md_path = target_doc.with_suffix(".md")

        md_content = md_path.read_text(encoding="utf-8", errors="replace") if md_path.exists() else ""
        max_chars = int(args.get("max_chars") or 12000)
        return {
            "tdoc_id": tdoc_id,
            "metadata": {
                k: exact_meta.get(k)
                for k in (
                    "tdoc_number",
                    "title",
                    "source",
                    "confirmed_companies",
                    "tentative_companies",
                    "doc_type",
                    "agenda_item",
                    "work_item",
                    "tdoc_status",
                    "secretary_remarks",
                    "wg_name",
                    "meeting_number",
                    "folder_name",
                )
                if k in exact_meta
            },
            "source_file": str(target_doc),
            "markdown_path": str(md_path),
            "char_length": len(md_content),
            "truncated": len(md_content) > max_chars,
            "markdown": md_content[:max_chars],
        }

    if action == "get_or_set_stance":
        if not tdoc_id:
            return {"error": "tdoc_id is required for action='get_or_set_stance'."}
        clean_tid = tdoc_id.strip().upper()
        user_db = TDocsDatabase(user_tdocs_db_path)
        stance = args.get("stance")
        notes = args.get("notes")
        if stance is not None or notes is not None:
            existing = user_db.get_all_user_data().get(clean_tid) or user_db.get_all_user_data().get(tdoc_id) or {}
            new_status = str(stance if stance is not None else existing.get("my_status", "")).strip()
            new_notes = str(notes if notes is not None else existing.get("my_notes", "")).strip()
            user_db.update_tdoc_user_data(clean_tid, new_status, new_notes)

        mtg = _resolve_meeting_record(mdb, wg=wg, meeting_number=meeting_number or None, tdoc_id=clean_tid)
        folders = [str(mtg.get("folder_name") or "")] if mtg and mtg.get("folder_name") else []
        merged_map = _collect_user_tdoc_annotations(user_tdocs_db_path, folder_names=folders, wg=wg)
        data = merged_map.get(clean_tid) or merged_map.get(clean_tid.split("R")[0]) or {"my_status": "", "my_notes": ""}
        return {
            "tdoc_id": clean_tid,
            "my_status": data.get("my_status", ""),
            "my_notes": data.get("my_notes", ""),
        }

    if action == "meeting_stats":
        rows = _load_tdocs_for_meeting_or_db(
            mdb,
            specs_db_path=specs_db_path,
            spec_history_db_path=spec_history_db_path,
            wg=wg,
            meeting_number=meeting_number or None,
        )
        if not rows:
            return {"error": "No TDocs found for the specified meeting filter."}

        company_counts: Dict[str, int] = {}
        status_counts: Dict[str, int] = {}
        wi_counts: Dict[str, int] = {}
        for r in rows:
            st = str(r.get("tdoc_status") or "Unknown").strip() or "Unknown"
            status_counts[st] = status_counts.get(st, 0) + 1
            for comp in CompanySanitizer.get_matching_contributors(str(r.get("source") or "")):
                company_counts[comp] = company_counts.get(comp, 0) + 1
            for wi in _parse_csv_list(r.get("work_item")):
                wi_counts[wi] = wi_counts.get(wi, 0) + 1

        top_companies = sorted(company_counts.items(), key=lambda x: -x[1])[:20]
        top_wis = sorted(wi_counts.items(), key=lambda x: -x[1])[:20]
        return {
            "wg": wg or (rows[0].get("wg_name") if rows else None),
            "meeting_number": meeting_number or (rows[0].get("meeting_number") if rows else None),
            "folder_name": rows[0].get("folder_name") if rows else None,
            "total_tdocs": len(rows),
            "status_breakdown": status_counts,
            "top_confirmed_contributors": [{"company": k, "tdoc_count": v} for k, v in top_companies],
            "top_work_items": [{"work_item": k, "tdoc_count": v} for k, v in top_wis],
        }

    return {"error": f"Unsupported action '{action}' for meetings_and_tdocs."}


# =========================================================================
# DOMAIN 7: Chairman's Notes & SA2 Session Plan Management
# =========================================================================

def _handle_chair_notes_and_session_plan(
    args: Dict[str, Any],
    *,
    session_plan_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "query_session_plan").strip().lower()
    meeting_number = str(args.get("meeting_number") or "174").strip()
    file_path = str(args.get("file_path") or "").strip()

    from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
    from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffEngine
    from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
    from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser

    if action == "parse_chair_notes_file":
        if not file_path or not Path(file_path).exists():
            return {"error": "Valid file_path to a Chairman's Notes .docx file is required."}
        outcomes = ChairNotesAgreementsParser.parse_chair_notes_file(Path(file_path))
        return {
            "file_path": file_path,
            "outcomes_count": len(outcomes),
            "outcomes": [
                {
                    "tdoc_id": o.tdoc_id,
                    "canonical_status": o.canonical_status,
                    "raw_outcome": o.raw_outcome,
                    "ai_number": o.ai_number,
                    "chair_name": o.chair_name,
                    "note_version": o.note_version,
                    "follow_up_tdoc": o.follow_up_tdoc,
                }
                for o in outcomes
            ],
        }

    sp_db = SessionPlanDatabase(session_plan_db_path)

    if action == "ingest_session_plan":
        if file_path and Path(file_path).exists():
            snap = SessionPlanParser.parse_docx(
                Path(file_path),
                meeting_number=meeting_number,
                tdoc_id=str(args.get("tdoc_id") or ""),
                snapshot_label=str(args.get("snapshot_label") or ""),
            )
        elif args.get("markdown_text"):
            snap = SessionPlanParser.parse_markdown_or_text(
                str(args["markdown_text"]),
                meeting_number=meeting_number,
                tdoc_id=str(args.get("tdoc_id") or ""),
                snapshot_label=str(args.get("snapshot_label") or ""),
            )
        else:
            return {"error": "Provide file_path (.docx) or markdown_text for action='ingest_session_plan'."}
        snap_id = sp_db.save_snapshot(snap)
        return {
            "snapshot_id": snap_id,
            "meeting_number": meeting_number,
            "tdoc_id": snap.tdoc_id,
            "snapshot_label": snap.snapshot_label,
            "streams_count": len(snap.streams),
            "slots_count": len(snap.slots),
        }

    if action == "query_session_plan":
        snap = sp_db.get_active_snapshot(meeting_number)
        if not snap:
            return {"error": f"No active session plan snapshot found for meeting {meeting_number}."}
        q = str(args.get("query") or "").strip().lower()
        day_f = str(args.get("day_name") or "").strip().lower()
        slots_out = []
        for s in snap.slots:
            if day_f and not s.day_name.lower().startswith(day_f[:3]):
                continue
            if q:
                hay = f"{s.raw_text} {' '.join(i.ai_number + ' ' + i.topic_label for i in s.items)}".lower()
                if q not in hay:
                    continue
            slots_out.append({
                "day_name": s.day_name,
                "quarter": s.quarter,
                "sub_slot": f"TU {s.sub_slot}" if s.sub_slot else "Full",
                "time": f"{s.start_time_local}-{s.end_time_local}",
                "stream": s.stream_number,
                "room_override": s.room_override,
                "items": [
                    {"ai": i.ai_number, "topic": i.topic_label, "tdocs": i.tdoc_count_str}
                    for i in s.items
                ],
            })
        return {
            "meeting_number": meeting_number,
            "snapshot_id": snap.id,
            "snapshot_label": snap.snapshot_label,
            "matching_slots_count": len(slots_out),
            "slots": slots_out[: int(args.get("limit") or 50)],
        }

    if action == "diff_session_plans":
        snaps = sp_db.list_snapshots(meeting_number)
        if len(snaps) < 2:
            return {"error": f"At least 2 session plan snapshots are required for meeting {meeting_number} to diff."}
        old_id = int(args.get("old_snapshot_id") or snaps[1]["id"])
        new_id = int(args.get("new_snapshot_id") or snaps[0]["id"])
        old_snap = sp_db.get_snapshot_by_id(old_id)
        new_snap = sp_db.get_snapshot_by_id(new_id)
        if not old_snap or not new_snap:
            return {"error": "Requested snapshot IDs were not found."}
        changes = SessionPlanDiffEngine.compare_snapshots(old_snap, new_snap)
        return {
            "meeting_number": meeting_number,
            "old_snapshot": old_snap.snapshot_label,
            "new_snapshot": new_snap.snapshot_label,
            "changes_count": len(changes),
            "changes": [
                {
                    "change_type": c.change_type,
                    "ai_number": c.ai_number,
                    "topic_label": c.topic_label,
                    "old_slot_desc": c.old_slot_desc,
                    "new_slot_desc": c.new_slot_desc,
                    "details": c.details,
                }
                for c in changes
            ],
        }

    return {"error": f"Unsupported action '{action}' for chair_notes_and_session_plan."}
