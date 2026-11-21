"""External agent handler for On-Site 10.10.10.10 Meeting Cockpit and live LAN operations."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import difflib
import logging
import re
import urllib.parse
from pathlib import Path
from typing import Any, Dict, List, Set

from bs4 import BeautifulSoup

from core.network.network_state import NetworkState
from core.network.session import NetworkSession
from .common import _parse_csv_list
from .meeting_helpers import _resolve_meeting_cache_dir, _resolve_meeting_record

logger = logging.getLogger(__name__)


# =========================================================================
# DOMAIN 5: On-Site 10.10.10.10 Meeting Cockpit & Live LAN Operations
# =========================================================================

def _parse_10_10_10_10_directory_html(html_text: str, base_url: str) -> List[Dict[str, Any]]:
    """Parses IIS / Apache directory listings from http://10.10.10.10/ftp/... or www.3gpp.org/ftp/..."""
    soup = BeautifulSoup(html_text or "", "html.parser")
    entries: List[Dict[str, Any]] = []
    seen: Set[str] = set()
    for a in soup.find_all("a", href=True):
        href = str(a.get("href") or "").strip()
        label = a.get_text(strip=True)
        if not href or href in ("/", "../") or "Parent Directory" in label or href.startswith("?"):
            continue
        full_url = urllib.parse.urljoin(base_url.rstrip("/") + "/", href)
        if full_url in seen:
            continue
        seen.add(full_url)
        name = urllib.parse.unquote(full_url.rstrip("/").split("/")[-1])
        if not name or name.lower() in ("ftp", "www.3gpp.org"):
            continue
        is_dir = href.endswith("/") or "." not in name
        entries.append({
            "name": name,
            "is_dir": is_dir,
            "url": full_url,
        })
    return entries


def _handle_onsite_meeting_cockpit(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
    session_plan_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "network_status").strip().lower()
    wg = str(args.get("wg") or "SA2").strip().upper()
    meeting_number = str(args.get("meeting_number") or "").strip()
    tdoc_id = str(args.get("tdoc_id") or "").strip()

    from modules.meetings.core.meetings_db import MeetingsDatabase
    from modules.meetings.core.url_router import URLRouter

    net_state = NetworkState.get_instance()
    net_status = net_state.ensure_probed(force=bool(args.get("force_probe", False)))

    if action == "network_status":
        sample_urls = URLRouter.build_priority_url_list(
            wg_name=wg,
            folder_name=f"TSG{wg}_{meeting_number or '174'}",
            main_ftp_url=f"https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_{meeting_number or '174'}",
            is_active_sync=True,
            target_filename="S2-2600001r01.zip",
        )
        return {
            **net_status,
            "wg": wg,
            "local_server_base": URLRouter._get_local_server_base(wg),
            "priority_url_tiers": sample_urls,
        }

    if action == "browse_lan_server":
        subfolder = str(args.get("subfolder") or "").strip().strip("/")
        base_lan = URLRouter._get_local_server_base(wg)
        target_url = f"{base_lan}/{subfolder}" if subfolder else base_lan

        if not net_status.get("is_local_reachable"):
            return {
                "is_local_reachable": False,
                "attempted_url": target_url,
                "error": (
                    "On-site server 10.10.10.10 is not reachable on the current network "
                    f"(active profile: '{net_status.get('network_name') or 'Unknown'}'). "
                    "Connect to 3GPPWIFI at the meeting venue or use public 3GPP FTP actions."
                ),
            }

        html_text = NetworkSession.get_html(target_url, timeout=8, apply_delay=False)
        if not html_text:
            return {"error": f"No directory listing returned from {target_url}"}
        entries = _parse_10_10_10_10_directory_html(html_text, target_url)
        filt = str(args.get("query") or "").strip().lower()
        if filt:
            entries = [e for e in entries if filt in e["name"].lower()]
        return {
            "url": target_url,
            "total_entries": len(entries),
            "entries": entries[: int(args.get("limit") or 100)],
        }

    if action == "fetch_live_revisions":
        if not tdoc_id:
            return {"error": "tdoc_id (e.g. 'S2-2601234') is required for action='fetch_live_revisions'."}
        base_m = re.match(r"^([A-Z][A-Z0-9]{1,3}-\d{5,7})", tdoc_id.upper())
        base_tdoc = base_m.group(1) if base_m else tdoc_id.upper()

        mdb = MeetingsDatabase(specs_db_path)
        mtg = _resolve_meeting_record(mdb, wg=wg, meeting_number=meeting_number) or {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        main_ftp = docs_ftp.rsplit("/Docs", 1)[0] if "/Docs" in docs_ftp else docs_ftp
        mtg_folder = str(mtg.get("folder_name") or mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        meeting_dir = _resolve_meeting_cache_dir(mtg_folder, wg=wg)
        revs_cache_dir = meeting_dir / "Revisions" / base_tdoc
        revs_cache_dir.mkdir(parents=True, exist_ok=True)

        candidate_dirs = URLRouter.build_priority_url_list(
            wg_name=wg,
            folder_name=mtg_folder,
            main_ftp_url=main_ftp,
            is_active_sync=True,
            target_filename=f"{base_tdoc}r01.zip",
        )
        discovered_files: List[Dict[str, str]] = []
        seen_names: Set[str] = set()

        for dir_url in candidate_dirs:
            html_text = NetworkSession.get_html(dir_url, timeout=8)
            if not html_text:
                continue
            for entry in _parse_10_10_10_10_directory_html(html_text, dir_url):
                if entry["is_dir"]:
                    continue
                fname = entry["name"]
                if base_tdoc.lower() in fname.lower() and fname.lower() not in seen_names:
                    seen_names.add(fname.lower())
                    discovered_files.append({
                        "name": fname,
                        "url": entry["url"],
                        "source_dir": dir_url,
                    })

        from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        downloaded_revs: List[Dict[str, Any]] = []
        for item in sorted(discovered_files, key=lambda x: x["name"].lower()):
            dl_target = revs_cache_dir / item["name"]
            if not dl_target.exists() or dl_target.stat().st_size == 0:
                NetworkSession.download_file(item["url"], dl_target)
            if not dl_target.exists():
                continue

            doc_files: List[Path] = []
            if dl_target.suffix.lower() == ".zip":
                ext_dir = revs_cache_dir / dl_target.stem
                ext_dir.mkdir(parents=True, exist_ok=True)
                doc_files = TDocFileHandler.extract_zip_safely(dl_target, ext_dir)
            elif dl_target.suffix.lower() in (".docx", ".doc"):
                doc_files = [dl_target]

            md_path = None
            md_text = ""
            if doc_files:
                primary_doc = doc_files[0]
                if primary_doc.suffix.lower() == ".doc":
                    from modules.word_tools.core.word_converter import convert_doc_to_docx
                    conv = convert_doc_to_docx(str(primary_doc))
                    if conv and Path(conv).exists():
                        primary_doc = Path(conv)
                md_out = primary_doc.with_suffix(".md")
                res = convert_document_to_markdown(primary_doc, md_out, revisions="markup")
                if res and md_out.exists():
                    md_path = str(md_out)
                    md_text = md_out.read_text(encoding="utf-8", errors="replace")

            downloaded_revs.append({
                "archive_or_file": item["name"],
                "source_url": item["url"],
                "local_path": str(dl_target),
                "markdown_path": md_path,
                "markdown_excerpt": md_text[:2500] if md_text else "",
                "_full_md": md_text,
            })

        diff_text = ""
        if len(downloaded_revs) >= 2 and downloaded_revs[-2].get("_full_md") and downloaded_revs[-1].get("_full_md"):
            prev_rev = downloaded_revs[-2]
            curr_rev = downloaded_revs[-1]
            diff_lines = list(
                difflib.unified_diff(
                    prev_rev["_full_md"].splitlines(),
                    curr_rev["_full_md"].splitlines(),
                    fromfile=prev_rev["archive_or_file"],
                    tofile=curr_rev["archive_or_file"],
                    lineterm="",
                )
            )
            diff_text = "\n".join(diff_lines[:400])

        for r in downloaded_revs:
            r.pop("_full_md", None)

        return {
            "base_tdoc": base_tdoc,
            "wg": wg,
            "revisions_found": len(downloaded_revs),
            "revisions": downloaded_revs,
            "latest_revision_diff": diff_text,
        }

    if action == "sync_chair_notes":
        from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
        from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser

        mdb = MeetingsDatabase(specs_db_path)
        mtg = _resolve_meeting_record(mdb, wg=wg, meeting_number=meeting_number) or {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        main_ftp = docs_ftp.rsplit("/Docs", 1)[0] if "/Docs" in docs_ftp else docs_ftp
        mtg_folder = str(mtg.get("folder_name") or mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        chair_dir = _resolve_meeting_cache_dir(mtg_folder, wg=wg) / "Chair_Notes"
        chair_dir.mkdir(parents=True, exist_ok=True)

        candidates = []
        if net_status.get("is_local_reachable"):
            candidates.append(URLRouter._get_local_server_base(wg))
        candidates.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{wg}")
        if main_ftp:
            candidates.append(main_ftp)

        session = NetworkSession.get_instance()
        found_url, html_content = ChairNotesDownloaderThread.locate_chair_notes_dir(session, candidates)
        downloaded: List[str] = []
        if found_url and html_content:
            file_tuples = ChairNotesDownloaderThread.parse_file_list_static(html_content, found_url)
            if len(file_tuples) > 1:
                file_tuples = ChairNotesDownloaderThread.select_best_chair_notes_files(file_tuples)
            if file_tuples:
                downloaded = ChairNotesDownloaderThread.download_files_sync(session, file_tuples, chair_dir)

        local_files = sorted(
            [p for p in chair_dir.rglob("*") if p.is_file() and p.suffix.lower() in (".docx", ".doc") and not p.name.startswith("~$")],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        agreements = []
        for f in local_files[:10]:
            try:
                agreements.extend(ChairNotesAgreementsParser.parse_chair_notes_file(f))
            except Exception as e:
                logger.debug(f"Could not parse chair notes {f.name}: {e}")

        filt_q = str(args.get("query") or "").strip().lower()
        if filt_q:
            agreements = [
                a for a in agreements
                if filt_q in str(getattr(a, "tdoc_id", "")).lower()
                or filt_q in str(getattr(a, "ai_number", "")).lower()
                or filt_q in str(getattr(a, "raw_outcome", "")).lower()
            ]

        return {
            "wg": wg,
            "meeting_folder": mtg_folder,
            "remote_source_url": found_url,
            "newly_downloaded_count": len(downloaded),
            "cached_files": [str(p) for p in local_files[:15]],
            "parsed_outcomes_count": len(agreements),
            "outcomes": [
                {
                    "tdoc_id": getattr(a, "tdoc_id", ""),
                    "canonical_status": getattr(a, "canonical_status", ""),
                    "raw_outcome": getattr(a, "raw_outcome", ""),
                    "ai_number": getattr(a, "ai_number", ""),
                    "chair_name": getattr(a, "chair_name", ""),
                    "note_version": getattr(a, "note_version", ""),
                    "follow_up_tdoc": getattr(a, "follow_up_tdoc", ""),
                }
                for a in agreements[: int(args.get("limit") or 50)]
            ],
        }

    if action == "where_should_we_go_now":
        from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
        from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone

        sp_db = SessionPlanDatabase(session_plan_db_path)
        mtg_num = meeting_number
        if not mtg_num:
            with sp_db._get_connection() as conn:
                r = conn.execute("SELECT meeting_number FROM session_plan_snapshots ORDER BY id DESC LIMIT 1").fetchone()
                mtg_num = str(r["meeting_number"]) if r else "174"

        meta = sp_db.get_meeting_meta(mtg_num) or {}
        tz_offset = meta.get("tz_offset_minutes")
        if tz_offset is None:
            tz_offset = SessionPlanTimezone.get_laptop_tz_offset_minutes()
        tz_offset = int(tz_offset)

        venue_dt = SessionPlanTimezone.get_venue_current_datetime(tz_offset)
        day_filter = str(args.get("day_name") or venue_dt.strftime("%A")).strip()
        time_filter = str(args.get("time_hhmm") or venue_dt.strftime("%H:%M")).strip()

        snap = sp_db.get_active_snapshot(mtg_num)
        if not snap:
            return {
                "meeting_number": mtg_num,
                "venue_current_time": venue_dt.strftime("%Y-%m-%d %H:%M (%A)"),
                "laptop_tz": SessionPlanTimezone.get_laptop_tz_label(),
                "error": f"No active Session Plan snapshot found for meeting {mtg_num}.",
            }

        plan = sp_db.get_delegation_plan(mtg_num)
        cur_min = SessionPlanTimezone._time_str_to_minutes(time_filter) or (9 * 60)

        active_or_next_slots = []
        day_slots = [s for s in snap.slots if s.day_name.lower().startswith(day_filter[:3].lower())]
        if not day_slots:
            day_slots = snap.slots

        for s in day_slots:
            s_min = SessionPlanTimezone._time_str_to_minutes(s.start_time_local) or 0
            e_min = SessionPlanTimezone._time_str_to_minutes(s.end_time_local) or 0
            is_now = s_min <= cur_min <= e_min
            is_upcoming = 0 < (s_min - cur_min) <= 120
            if is_now or is_upcoming or args.get("day_name"):
                stream_info = snap.streams.get(s.stream_number)
                room = s.room_override or (stream_info.room_name if stream_info else f"Stream {s.stream_number}")
                assigned_delegates = set()
                items_out = []
                for it in s.items:
                    owner_obj = plan.topic_owners.get(it.ai_number)
                    owner_name = owner_obj.owner_name if owner_obj else ""
                    if owner_name:
                        assigned_delegates.add(owner_name)
                    items_out.append({
                        "ai_number": it.ai_number,
                        "topic_label": it.topic_label,
                        "tdoc_count": it.tdoc_count_str,
                        "owner": owner_name,
                    })
                slot_key = f"{s.day_name}_{s.quarter}_TU{s.sub_slot}_S{s.stream_number}"
                appt = plan.appointments.get(slot_key)
                if appt and appt.appointed_delegate:
                    assigned_delegates.add(appt.appointed_delegate)

                active_or_next_slots.append({
                    "state": "LIVE_NOW" if is_now else ("UPCOMING" if is_upcoming else "SCHEDULED"),
                    "day_name": s.day_name,
                    "quarter": s.quarter,
                    "sub_slot": f"TU {s.sub_slot}" if s.sub_slot else "Full",
                    "venue_time": f"{s.start_time_local} - {s.end_time_local}",
                    "dual_time": SessionPlanTimezone.format_dual_time_range(s.start_time_local, s.end_time_local, tz_offset),
                    "stream": s.stream_number,
                    "room": room,
                    "assigned_delegates": sorted(assigned_delegates),
                    "items": items_out,
                    "raw_text": s.raw_text,
                })

        all_roster = sorted(plan.members.keys())
        busy_now = {
            d
            for slot in active_or_next_slots
            if slot["state"] == "LIVE_NOW"
            for d in slot["assigned_delegates"]
        }
        free_now = [d for d in all_roster if d not in busy_now]

        return {
            "meeting_number": mtg_num,
            "snapshot_label": snap.snapshot_label,
            "venue_timezone": meta.get("tz_name") or f"UTC+{tz_offset // 60:02d}:00",
            "venue_current_time": venue_dt.strftime("%Y-%m-%d %H:%M (%A)"),
            "laptop_timezone": SessionPlanTimezone.get_laptop_tz_label(),
            "roster": all_roster,
            "free_colleagues_now": free_now,
            "slots": active_or_next_slots[: int(args.get("limit") or 30)],
        }

    if action == "batch_cache_tdocs":
        from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        tdoc_list = _parse_csv_list(args.get("tdoc_id") or args.get("query"))
        if not tdoc_list:
            return {"error": "Provide one or more TDoc IDs in tdoc_id or query (comma-separated)."}

        mdb = MeetingsDatabase(specs_db_path)
        mtg = _resolve_meeting_record(mdb, wg=wg, meeting_number=meeting_number) or {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        rev_url = str(mtg.get("revisions_url") or "")
        mtg_folder = str(mtg.get("folder_name") or mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        meeting_dir = _resolve_meeting_cache_dir(mtg_folder, wg=wg)
        convert_md = bool(args.get("convert_to_md", True))

        results: List[Dict[str, Any]] = []

        def _fetch_one(tid: str) -> Dict[str, Any]:
            files, err = TDocFileHandler.download_and_extract_tdoc(
                tdoc_id=tid,
                docs_ftp_url=docs_ftp,
                meeting_dir=meeting_dir,
                wg_name=wg,
                revisions_url=rev_url,
            )
            if err or not files:
                return {"tdoc_id": tid, "success": False, "error": err or "Not found"}
            md_files = []
            if convert_md:
                for f in files:
                    if f.suffix.lower() in (".docx", ".doc"):
                        md_out = f.with_suffix(".md")
                        res = convert_document_to_markdown(f, md_out, revisions="markup")
                        if res and md_out.exists():
                            md_files.append(str(md_out))
            return {
                "tdoc_id": tid,
                "success": True,
                "files": [str(f) for f in files],
                "markdown_files": md_files,
            }

        with ThreadPoolExecutor(max_workers=min(6, len(tdoc_list))) as pool:
            futs = [pool.submit(_fetch_one, t) for t in tdoc_list[:50]]
            for fut in as_completed(futs):
                results.append(fut.result())

        return {
            "wg": wg,
            "meeting_folder": mtg_folder,
            "requested_count": len(tdoc_list[:50]),
            "succeeded_count": sum(1 for r in results if r.get("success")),
            "results": results,
        }

    return {"error": f"Unsupported action '{action}' for onsite_meeting_cockpit."}
