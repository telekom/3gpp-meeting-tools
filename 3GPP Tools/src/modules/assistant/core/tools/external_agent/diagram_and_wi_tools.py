"""External agent handlers for PlantUML/Office Studio and Work Items / Emails."""

from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.paths import get_output_root, get_temp_root
from .common import (
    _COM_LOCK,
    _disable_faulthandler_for_com,
    _parse_csv_list,
    resolve_meetings_cache_roots as _get_meetings_cache_roots,
)


# =========================================================================
# DOMAIN 11: PlantUML, Visio (.vsdx) & PowerPoint (.pptx) Studio
# =========================================================================

def _handle_plantuml_office_studio(args: Dict[str, Any]) -> Dict[str, Any]:
    action = str(args.get("action") or "validate_or_render_svg").strip().lower()
    puml_code = str(args.get("puml_code") or "").strip()
    input_path_raw = str(args.get("input_path") or "").strip()
    output_path_raw = str(args.get("output_path") or "").strip()

    from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
    from modules.puml2visio.utils.paths import get_puml2visio_asset_path
    from modules.puml2visio.utils.utils import generate_cleaned_svg

    jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)

    if action == "list_presets":
        from modules.puml2visio.templates.plantuml_snippets import PLANTUML_SNIPPETS, reload_snippets

        snippets = reload_snippets() or PLANTUML_SNIPPETS
        q = str(args.get("query") or "").strip().lower()
        cat = str(args.get("category") or "").strip().lower()
        matched = []
        for s in snippets:
            if cat and cat not in s.category.lower():
                continue
            if q and q not in f"{s.id} {s.label} {s.description} {s.category}".lower():
                continue
            matched.append({
                "id": s.id,
                "label": s.label,
                "category": s.category,
                "description": s.description,
                "code": s.code.replace("{cursor}", "").strip(),
            })
        return {
            "count": len(matched),
            "presets": matched[: int(args.get("limit") or 50)],
        }

    if action == "extract_puml_from_vsdx":
        if not input_path_raw or not Path(input_path_raw).exists():
            return {"error": "Valid input_path to a .vsdx file is required."}
        from modules.puml2visio.ui.cards.batch_tab_utils import extract_plantuml_from_vsdx_fast

        vsdx_path = Path(input_path_raw)
        extracted = extract_plantuml_from_vsdx_fast(vsdx_path)
        if not extracted:
            return {
                "input_path": str(vsdx_path),
                "has_embedded_plantuml": False,
                "error": (
                    f"No embedded PlantUML metadata found in '{vsdx_path.name}'. "
                    "Note: Visio .vsdx reverse engineering only works on .vsdx files previously generated "
                    "by 3GPP Tools (with 'Embed PlantUML Source Code' enabled), as arbitrary hand-drawn "
                    "Visio diagrams do not contain the 'PlantUML Source' metadata page."
                ),
            }
        return {
            "input_path": str(vsdx_path),
            "has_embedded_plantuml": True,
            "plantuml_code": extracted,
        }

    if action in ("validate_or_render_svg", "export_vsdx", "export_pptx"):
        work_dir = Path(tempfile.mkdtemp(prefix="puml_agent_", dir=str(get_temp_root())))
        if input_path_raw and Path(input_path_raw).exists():
            puml_file = Path(input_path_raw)
            puml_code = puml_file.read_text(encoding="utf-8", errors="replace")
        elif puml_code:
            if "@startuml" not in puml_code:
                puml_code = f"@startuml\n{puml_code}\n@enduml"
            puml_file = work_dir / "diagram.puml"
            puml_file.write_text(puml_code, encoding="utf-8")
        else:
            return {"error": "Provide puml_code or input_path (.puml) for PlantUML rendering/export."}

        if action == "validate_or_render_svg":
            try:
                svg_path = generate_cleaned_svg(puml_file, jar_path)
                out_svg = Path(output_path_raw) if output_path_raw else (get_output_root() / f"{puml_file.stem}.svg")
                out_svg.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(svg_path, out_svg)
                return {
                    "valid": True,
                    "svg_path": str(out_svg),
                }
            except Exception as e:
                return {
                    "valid": False,
                    "error": str(e),
                }

        if action == "export_vsdx":
            from modules.puml2visio.core.visio_converter import ConverterThread

            out_dir = Path(output_path_raw).parent if output_path_raw else get_output_root()
            out_dir.mkdir(parents=True, exist_ok=True)
            _disable_faulthandler_for_com()
            with _COM_LOCK:
                worker = ConverterThread(
                    files=[puml_file],
                    jar_path=jar_path,
                    output_dir=out_dir,
                    open_visio=bool(args.get("open_after", False)),
                    include_source=bool(args.get("embed_source", True)),
                )
                worker.run()
            expected_vsdx = out_dir / f"{puml_file.stem}.vsdx"
            if output_path_raw and expected_vsdx.exists() and Path(output_path_raw).resolve() != expected_vsdx.resolve():
                shutil.move(str(expected_vsdx), output_path_raw)
                expected_vsdx = Path(output_path_raw)
            return {
                "output_path": str(expected_vsdx),
                "exists": expected_vsdx.exists(),
                "embedded_plantuml_metadata": bool(args.get("embed_source", True)),
            }

        if action == "export_pptx":
            from modules.puml2visio.core.powerpoint_converter import PptxConverterThread

            out_pptx = Path(output_path_raw) if output_path_raw else (get_output_root() / f"{puml_file.stem}.pptx")
            out_pptx.parent.mkdir(parents=True, exist_ok=True)
            _disable_faulthandler_for_com()
            with _COM_LOCK:
                worker = PptxConverterThread(
                    puml_path=puml_file,
                    jar_path=jar_path,
                    output_pptx=out_pptx,
                    include_notes=True,
                    open_after=bool(args.get("open_after", False)),
                    append_to_open=bool(args.get("append_to_open", False)),
                    insert_position=str(args.get("insert_position") or "end"),
                )
                worker.run()
            return {
                "output_path": str(out_pptx),
                "exists": out_pptx.exists(),
                "appended_to_open": bool(args.get("append_to_open", False)),
            }

    if action == "bridge_pptx_vsdx":
        if not input_path_raw or not Path(input_path_raw).exists():
            return {"error": "Valid input_path (.pptx or .vsdx) is required for action='bridge_pptx_vsdx'."}
        src = Path(input_path_raw)
        _disable_faulthandler_for_com()
        with _COM_LOCK:
            if src.suffix.lower() == ".pptx":
                from modules.puml2visio.core.pptx2visio_converter import PptxToVisioConverterThread
                w = PptxToVisioConverterThread(src)
                w.run()
                out = src.with_suffix(".vsdx")
            elif src.suffix.lower() == ".vsdx":
                from modules.puml2visio.core.visio2pptx_converter import VisioToPptxConverterThread
                w = VisioToPptxConverterThread(src)
                w.run()
                out = src.with_suffix(".pptx")
            else:
                return {"error": "input_path must be a .pptx or .vsdx file."}
        return {
            "input_path": str(src),
            "output_path": str(out),
            "exists": out.exists(),
        }

    return {"error": f"Unsupported action '{action}' for plantuml_office_studio."}


# =========================================================================
# DOMAIN 12: 3GPP Work Items & Working Group Reflector Emails
# =========================================================================

def _normalize_release_filters(release: Optional[str]) -> Optional[List[str]]:
    """Normalizes a release argument (e.g. 'Rel-20', '20', 'R99') into all matching DB forms."""
    if not release:
        return None
    raw_items = _parse_csv_list(release)
    if not raw_items:
        return None
    expanded: List[str] = []
    for item in raw_items:
        low = item.lower()
        if low in ("all", "all releases", "*"):
            return None
        if item not in expanded:
            expanded.append(item)
        m = re.match(r"^(?:rel(?:ease)?[-_\s]*)?(\d+)$", item, re.IGNORECASE)
        if m:
            num = m.group(1)
            for variant in (f"Rel-{num}", num, f"R{num}"):
                if variant not in expanded:
                    expanded.append(variant)
    return expanded or None


def _handle_work_items_and_emails(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search_work_items").strip().lower()
    query = str(args.get("query") or "").strip()
    wg = str(args.get("wg") or "").strip().upper() or None
    release = str(args.get("release") or "").strip() or None
    limit = max(1, int(args.get("limit") or 25))

    from modules.work_items.core.wi_database import WorkItemsDatabase
    from modules.work_items.core.wi_settings import WorkItemsSettings

    wi_db = WorkItemsDatabase(specs_db_path)
    wi_settings = WorkItemsSettings()

    if action == "search_work_items":
        starred, following = wi_settings.get_tracking_sets()
        wg_list = [w for w in _parse_csv_list(wg) if w.lower() not in ("all", "all wgs", "*")] if wg else None
        rel_list = _normalize_release_filters(release)
        tracked_only = bool(args.get("tracked_only", False))

        if tracked_only:
            tracked_codes = starred | following
            if not tracked_codes:
                return {
                    "count": 0,
                    "starred_codes": sorted(starred),
                    "following_codes": sorted(following),
                    "work_items": [],
                }
            all_rows = wi_db.search_work_items(
                search_term=query or None,
                releases=rel_list,
                wg_names=wg_list,
                limit=None,
            )
            rows = [
                r for r in all_rows
                if str(r.get("code") or "") in tracked_codes
                or str(r.get("acronym") or "") in tracked_codes
            ]
        else:
            rows = wi_db.search_work_items(
                search_term=query or None,
                releases=rel_list,
                wg_names=wg_list,
                limit=limit,
            )

        selected = rows[:limit]
        for r in selected:
            c = str(r.get("code") or "")
            acr = str(r.get("acronym") or "")
            r["starred"] = (c in starred) or (bool(acr) and acr in starred)
            r["following"] = (c in following) or (bool(acr) and acr in following)
        return {
            "count": len(selected),
            "starred_codes": sorted(starred),
            "following_codes": sorted(following),
            "work_items": selected,
        }

    if action == "fetch_wi_from_portal":
        if not query:
            return {"error": "query (Work Item numeric code or URL) is required for action='fetch_wi_from_portal'."}
        from modules.work_items.core.wi_scraper import fetch_wi_details_from_portal

        return fetch_wi_details_from_portal(query)

    if action == "toggle_wi_tracking":
        wi_code = str(args.get("wi_code") or query).strip()
        if not wi_code:
            return {"error": "wi_code is required for action='toggle_wi_tracking'."}
        starred, following = wi_settings.get_tracking_sets()
        if args.get("starred") is not None:
            if bool(args["starred"]):
                starred.add(wi_code)
            else:
                starred.discard(wi_code)
        if args.get("following") is not None:
            if bool(args["following"]):
                following.add(wi_code)
            else:
                following.discard(wi_code)
        wi_settings.save_tracking_sets(starred, following)
        return {
            "wi_code": wi_code,
            "starred": wi_code in starred,
            "following": wi_code in following,
        }

    if action == "search_emails":
        from modules.emails.core.email_db import EmailDatabase
        from modules.emails.core.general_email_db import GeneralEmailDatabase

        results: List[Dict[str, Any]] = []
        q_like = f"%{query}%"

        # 1. Check specs_db_path first (supports both canonical `emails` / `general_emails` and legacy fixture tables)
        for db_p in [specs_db_path]:
            if not db_p or not db_p.exists():
                continue
            try:
                email_db = EmailDatabase(db_p)
                with email_db._get_connection() as conn:
                    tbls = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
                    if "emails" in tbls:
                        sql = """
                            SELECT id, tdoc_id AS tdoc_number, subject, sender_name, company AS sender_company,
                                   date_received AS sent_date, substr(COALESCE(free_text, short_text, ''), 1, 800) AS snippet
                            FROM emails
                            WHERE (? = '' OR tdoc_id LIKE ? OR subject LIKE ? OR free_text LIKE ? OR short_text LIKE ? OR company LIKE ?)
                            ORDER BY date_received DESC LIMIT ?
                        """
                        rows = conn.execute(sql, (query, q_like, q_like, q_like, q_like, q_like, limit)).fetchall()
                        results.extend([dict(r) for r in rows])
                    if "meeting_emails" in tbls:
                        sql = """
                            SELECT id, tdoc_number, subject, sender_name, sender_company, sent_date, substr(body_clean, 1, 800) AS snippet
                            FROM meeting_emails
                            WHERE (? = '' OR tdoc_number LIKE ? OR subject LIKE ? OR body_clean LIKE ? OR sender_company LIKE ?)
                            ORDER BY sent_date DESC LIMIT ?
                        """
                        rows = conn.execute(sql, (query, q_like, q_like, q_like, q_like, limit)).fetchall()
                        results.extend([dict(r) for r in rows])
            except Exception:
                pass

            try:
                gen_db = GeneralEmailDatabase(db_p)
                with gen_db._get_connection() as conn:
                    tbls = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
                    if "general_emails" in tbls:
                        sql = """
                            SELECT id, folder_tag AS wg_name, subject, sender_name, company AS sender_company,
                                   date_received AS sent_date, substr(COALESCE(body_text, ''), 1, 800) AS snippet
                            FROM general_emails
                            WHERE (? = '' OR subject LIKE ? OR body_text LIKE ? OR company LIKE ?)
                            ORDER BY date_received DESC LIMIT ?
                        """
                        rows = conn.execute(sql, (query, q_like, q_like, q_like, limit)).fetchall()
                        results.extend([dict(r) for r in rows])
                    if "general_wg_emails" in tbls:
                        sql = """
                            SELECT id, wg_name, subject, sender_name, sender_company, sent_date, substr(body_clean, 1, 800) AS snippet
                            FROM general_wg_emails
                            WHERE (? = '' OR subject LIKE ? OR body_clean LIKE ? OR sender_company LIKE ?)
                            ORDER BY sent_date DESC LIMIT ?
                        """
                        rows = conn.execute(sql, (query, q_like, q_like, q_like, limit)).fetchall()
                        results.extend([dict(r) for r in rows])
            except Exception:
                pass

        # 2. Also check per-meeting Agenda/emails.db and global general_emails.db in cache_dir if needed
        if len(results) < limit:
            for c_root in _get_meetings_cache_roots():
                if not c_root.exists():
                    continue
                gen_db_file = c_root / "General_Emails" / "general_emails.db"
                if gen_db_file.exists():
                    try:
                        gdb = GeneralEmailDatabase(gen_db_file)
                        with gdb._get_connection() as conn:
                            sql = """
                                SELECT id, folder_tag AS wg_name, subject, sender_name, company AS sender_company,
                                       date_received AS sent_date, substr(COALESCE(body_text, ''), 1, 800) AS snippet
                                FROM general_emails
                                WHERE (? = '' OR subject LIKE ? OR body_text LIKE ? OR company LIKE ?)
                                ORDER BY date_received DESC LIMIT ?
                            """
                            rows = conn.execute(sql, (query, q_like, q_like, q_like, limit)).fetchall()
                            results.extend([dict(r) for r in rows])
                    except Exception:
                        pass

        return {
            "query": query,
            "count": len(results[:limit]),
            "emails": results[:limit],
        }

    if action in ("list_email_summaries", "get_email_summary", "save_email_summary"):
        from .email_summary_tools import handle_email_summary_mcp_action

        return handle_email_summary_mcp_action(action, args, specs_db_path=specs_db_path)

    return {"error": f"Unsupported action '{action}' for work_items_and_emails."}
