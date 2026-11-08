"""Compact, action-multiplexed 3GPP Tools suite for external CLI agents (Claude CLI, Codex CLI, MCP).

Design goals:
- Expose the complete 39-capability 3GPP Tools catalogue (including on-site 10.10.10.10 meeting LAN capabilities).
- Keep the MCP `tools/list` schema ultra-compact (~12 tools with short descriptions and action enums)
  so Claude CLI and Codex CLI do not suffer high system-prompt token overhead on every turn.
- Run 100% headlessly without blocking Qt modal dialogs, serializing Win32 COM operations safely.
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import difflib
import faulthandler
import json
import logging
import os
from pathlib import Path
import re
import shutil
import tempfile
import threading
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple
import urllib.parse

from bs4 import BeautifulSoup

from core.network.network_state import NetworkState
from core.network.session import NetworkSession
from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import (
    get_db_root,
    get_meetings_root,
    get_output_root,
    get_project_root,
    get_specs_root,
    get_temp_root,
    to_extended_path,
)
from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.assistant.core.knowledge import (
    ProtocolKnowledgeService,
    SpecificationKnowledgeService,
    SpecSearchKnowledgeService,
)
from modules.assistant.core.tool_registry import (
    ArgumentType,
    ToolArgument,
    ToolDefinition,
    ToolPermission,
    ToolRegistry,
)

logger = logging.getLogger(__name__)


def ToolArgumentSpec(
    name: str,
    arg_type: str,
    description: str,
    *,
    required: bool = False,
    default: Any = None,
    minimum: Optional[int] = None,
    maximum: Optional[int] = None,
    max_length: Optional[int] = None,
    enum: Tuple[Any, ...] = (),
) -> ToolArgument:
    """Helper constructing a `ToolArgument` from compact string type names."""
    type_map = {
        "string": ArgumentType.STRING,
        "integer": ArgumentType.INTEGER,
        "boolean": ArgumentType.BOOLEAN,
    }
    return ToolArgument(
        name=name,
        type=type_map.get(arg_type, ArgumentType.STRING),
        description=description,
        required=required,
        default=default,
        minimum=minimum,
        maximum=maximum,
        max_length=max_length,
        enum=tuple(enum),
    )


def _wrap_handler(fn: Callable[[Dict[str, Any]], Any]) -> Callable[..., ToolResult]:
    """Adapts a `fn(args_dict) -> Dict | ToolResult` handler to `fn(**kwargs) -> ToolResult`."""
    def _wrapped(**kwargs: Any) -> ToolResult:
        res = fn(dict(kwargs))
        if isinstance(res, ToolResult):
            is_ok = res.status in (ToolStatus.FOUND, ToolStatus.SOURCE_INCOMPLETE)
            if isinstance(res.data, dict):
                payload = dict(res.data)
                payload.setdefault("ok", is_ok)
                if not is_ok and res.message and "error" not in payload:
                    payload["error"] = res.message
                if res.metadata:
                    payload.setdefault("metadata", res.metadata)
            elif isinstance(res.data, list):
                payload = {
                    "ok": is_ok,
                    "items": res.data,
                    "matches": res.data,
                    "hits": res.data,
                    "count": len(res.data),
                }
                if not is_ok and res.message:
                    payload["error"] = res.message
                if res.metadata:
                    payload["metadata"] = res.metadata
            else:
                payload = {
                    "ok": is_ok,
                    "result": res.data,
                }
                if not is_ok and res.message:
                    payload["error"] = res.message
                if res.metadata:
                    payload["metadata"] = res.metadata
            return ToolResult(
                status=ToolStatus.FOUND,
                data=payload,
                message=res.message,
                evidence_ids=res.evidence_ids,
                metadata=res.metadata,
            )
        if isinstance(res, dict):
            if "error" in res and "ok" not in res:
                res["ok"] = False
            elif "ok" not in res:
                res["ok"] = True
            return ToolResult(
                status=ToolStatus.FOUND,
                data=res,
                message=str(res.get("error") or "ok"),
            )
        return ToolResult(status=ToolStatus.FOUND, data={"ok": True, "result": res})
    return _wrapped

# Global re-entrant lock for Win32 COM automation (Word, Visio, PowerPoint)
_COM_LOCK = threading.RLock()


def _disable_faulthandler_for_com() -> None:
    """Disables faulthandler on Windows so recoverable Win32 COM RPC_E_CALL_REJECTED SEH exceptions
    are caught cleanly by Python try/except blocks instead of printing alarming stderr tracebacks."""
    if os.name == "nt":
        try:
            faulthandler.disable()
        except Exception:
            pass


def _resolve_db_paths(
    specs_db_path: Optional[Path] = None,
    spec_search_db_path: Optional[Path] = None,
    protocol_db_path: Optional[Path] = None,
    spec_history_db_path: Optional[Path] = None,
    user_tdocs_db_path: Optional[Path] = None,
    session_plan_db_path: Optional[Path] = None,
) -> Dict[str, Path]:
    db_root = get_db_root()
    return {
        "specs_db": Path(specs_db_path or (db_root / "3gpp_data.db")),
        "spec_search_db": Path(spec_search_db_path or (db_root / "spec_search.db")),
        "protocol_db": Path(protocol_db_path or (db_root / "nas_knowledge.db")),
        "spec_history_db": Path(spec_history_db_path or (db_root / "3gpp_spec_history.db")),
        "user_tdocs_db": Path(user_tdocs_db_path or (db_root / "user_tdocs.db")),
        "session_plan_db": Path(session_plan_db_path or (db_root / "session_plans.db")),
    }


def _parse_csv_list(val: Any) -> List[str]:
    if not val:
        return []
    if isinstance(val, (list, tuple, set)):
        return [str(x).strip() for x in val if str(x).strip()]
    return [x.strip() for x in re.split(r"[,;]+", str(val)) if x.strip()]


# =========================================================================
# DOMAIN 1: Specification Catalogue, Files & Full Spec Markdown Conversion
# =========================================================================

def _handle_spec_catalogue_and_files(
    args: Dict[str, Any],
    *,
    spec_service: SpecificationKnowledgeService,
    spec_search_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search").strip().lower()
    spec_number = str(args.get("spec_number") or "").strip()
    version = str(args.get("version") or "").strip() or None
    release = str(args.get("release") or "").strip() or None
    query = str(args.get("query") or "").strip()
    limit = int(args.get("limit") or 10)

    if action == "search":
        return spec_service.find_specifications(
            query=query or spec_number,
            limit=limit,
        )

    if action == "get":
        if not spec_number:
            return {"error": "spec_number is required for action='get'."}
        res = spec_service.get_specification(specification=spec_number)
        if isinstance(res.data, dict) and "known_versions" in res.data:
            res.data.setdefault("versions", res.data["known_versions"])
            res.data.setdefault("count", len(res.data["known_versions"]))
        return res

    if action == "resolve_files":
        if not spec_number:
            return {"error": "spec_number is required for action='resolve_files'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.specifications.core.spec_history_extractor import SpecHistoryExtractor

        resolved_parts = [str(p) for p in SpecDocResolver.resolve_spec_files(spec_number, version=version)]
        terminal_doc = SpecHistoryExtractor.resolve_spec_document(spec_number, version=version)
        spec_dir = get_specs_root() / spec_number.replace(".", "")
        alt_dir = get_specs_root() / spec_number
        zips: List[str] = []
        for d in (spec_dir, alt_dir):
            if d.exists():
                zips.extend(str(p) for p in sorted(d.glob("*.zip")))
        return {
            "spec_number": spec_number,
            "version": version,
            "resolved_doc_parts": resolved_parts,
            "terminal_or_history_doc": str(terminal_doc) if terminal_doc else None,
            "cached_zip_archives": zips,
            "local_ready": bool(resolved_parts),
        }

    if action == "download":
        if not spec_number:
            return {"error": "spec_number is required for action='download'."}
        from modules.specifications.core.spec_autonomous_downloader import extract_word_from_zip
        from modules.specifications.core.database import SpecsDatabase

        specs_db = SpecsDatabase(spec_service.db_path)
        file_rows = specs_db.search_files(spec_number=spec_number)
        if not file_rows:
            return {"error": f"No versions found in catalogue for {spec_number}."}

        versions = []
        for r in file_rows:
            if isinstance(r, (tuple, list)) and len(r) >= 8:
                ver_s = str(r[5] or "")
                rel_s = ver_s.split(".")[0] if "." in ver_s else ""
                versions.append({"version": ver_s, "release": rel_s, "zip_url": str(r[6] or ""), "upload_date": str(r[7] or "")})

        chosen = None
        if version:
            clean_v = version.lstrip("vV").strip()
            for v in versions:
                if str(v.get("version", "")).lstrip("vV").strip() == clean_v:
                    chosen = v
                    break
        elif release:
            rel_digits = re.sub(r"\D", "", release)
            for v in versions:
                if str(v.get("release", "")) == rel_digits or str(v.get("version", "")).startswith(f"{rel_digits}."):
                    chosen = v
                    break
        if chosen is None:
            chosen = versions[0]

        zip_url = str(chosen.get("zip_url") or "").strip()
        if not zip_url:
            return {"error": f"No 3GPP FTP zip_url recorded for {spec_number} v{chosen.get('version')}."}

        clean_num = spec_number.replace(".", "")
        target_dir = get_specs_root() / clean_num
        target_dir.mkdir(parents=True, exist_ok=True)
        zip_name = urllib.parse.unquote(zip_url.split("/")[-1])
        if not zip_name.lower().endswith(".zip"):
            zip_name = f"{clean_num}-{chosen.get('version', 'latest')}.zip"
        local_zip = target_dir / zip_name

        if not local_zip.exists() or local_zip.stat().st_size == 0:
            ok = NetworkSession.download_file(zip_url, local_zip, apply_delay=True)
            if not ok or not local_zip.exists():
                return {"error": f"Failed to download {zip_url}"}

        extracted = extract_word_from_zip(local_zip, target_dir)
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        parts = [str(p) for p in SpecDocResolver.resolve_spec_files(spec_number, version=str(chosen.get("version") or ""))]
        return {
            "spec_number": spec_number,
            "version": chosen.get("version"),
            "release": chosen.get("release"),
            "zip_url": zip_url,
            "local_zip": str(local_zip),
            "extracted_doc": str(extracted) if extracted else None,
            "resolved_doc_parts": parts,
        }

    if action == "index_for_search":
        if not spec_number:
            return {"error": "spec_number is required for action='index_for_search'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.spec_search.core.spec_search_db import SpecSearchDatabase
        from modules.spec_search.core.spec_search_threads import extract_spec_metadata_from_filename
        from modules.spec_search.core.spec_text_extractor import SpecDocxExtractor
        from modules.specifications.core.database import SpecsDatabase

        parts = SpecDocResolver.resolve_spec_files(spec_number, version=version)
        if not parts:
            return {"error": f"No local .docx/.doc files found for {spec_number}. Run action='download' first."}

        search_db = SpecSearchDatabase(spec_search_db_path)
        _, inferred_ver, inferred_rel = extract_spec_metadata_from_filename(parts[0])
        actual_ver = (version or inferred_ver or "unknown").lstrip("vV")
        actual_rel = release or inferred_rel or (f"Rel-{actual_ver.split('.')[0]}" if "." in actual_ver else "Unknown")

        all_clauses = []
        for p in parts:
            if p.suffix.lower() == ".doc":
                from modules.word_tools.core.word_converter import convert_doc_to_docx
                conv = convert_doc_to_docx(str(p))
                if conv and Path(conv).exists():
                    p = Path(conv)
            clauses = SpecDocxExtractor.extract_clauses(p)
            all_clauses.extend(clauses)

        search_db.insert_parsed_spec(
            spec_number=spec_number,
            version=actual_ver,
            clauses=all_clauses,
            release_date="",
        )
        return {
            "spec_number": spec_number,
            "version": actual_ver,
            "release": actual_rel,
            "parts_indexed": [str(p) for p in parts],
            "clause_count": len(all_clauses),
        }

    if action == "convert_to_md":
        if not spec_number:
            return {"error": "spec_number is required for action='convert_to_md'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.specifications.core.spec_markdown_converter import convert_spec_to_markdown
        from modules.specifications.core.database import SpecsDatabase

        parts = SpecDocResolver.resolve_spec_files(spec_number, version=version)
        if not parts:
            return {"error": f"No local Word files found for {spec_number}. Run action='download' first."}

        docx_parts: List[Path] = []
        for p in parts:
            if p.suffix.lower() == ".doc":
                from modules.word_tools.core.word_converter import convert_doc_to_docx
                conv = convert_doc_to_docx(str(p))
                if conv and Path(conv).exists():
                    docx_parts.append(Path(conv))
            else:
                docx_parts.append(p)

        specs_db = SpecsDatabase(spec_service.db_path)
        info = specs_db.get_spec_details(spec_number) or {}
        out_dir = Path(args.get("output_dir") or (get_output_root() / "Spec_Markdown" / f"{spec_number}_{version or 'latest'}"))
        mode = str(args.get("mode") or "single").strip().lower()
        res = convert_spec_to_markdown(
            source_paths=docx_parts,
            output_dir=out_dir,
            spec_number=spec_number,
            version=version or "",
            spec_title=str(info.get("title") or ""),
            mode=mode if mode in ("single", "split") else "single",
            include_toc=True,
            extract_figures=bool(args.get("extract_figures", True)),
        )
        return {
            "spec_number": res.spec_number,
            "version": res.version,
            "mode": res.mode,
            "output_dir": str(res.output_dir),
            "master_file": str(res.master_file),
            "generated_files": [str(p) for p in res.generated_files[:50]],
            "total_clauses": res.total_clauses,
            "total_tables": res.total_tables,
            "total_figures": res.total_figures,
        }

    return {"error": f"Unsupported action '{action}' for spec_catalogue_and_files."}


# =========================================================================
# DOMAIN 2: Specification Text Search, Clause Retrieval, Diff & Extraction
# =========================================================================

def _handle_spec_text_and_clauses(
    args: Dict[str, Any],
    *,
    search_service: SpecSearchKnowledgeService,
    spec_search_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search_text").strip().lower()
    spec_number = str(args.get("spec_number") or "").strip()
    version = str(args.get("version") or "").strip() or None
    clause_id = str(args.get("clause_id") or "").strip()
    query = str(args.get("query") or "").strip()
    limit = int(args.get("limit") or 10)

    from modules.spec_search.core.spec_search_db import SpecSearchDatabase
    search_db = SpecSearchDatabase(spec_search_db_path)

    if action == "search_text":
        if not query:
            return {"error": "query is required for action='search_text'."}
        rel_raw = str(args.get("release") or "").strip()
        rel_int = None
        if rel_raw:
            m = re.search(r"(\d+)", rel_raw)
            if m:
                rel_int = int(m.group(1))
        return search_service.search_text(
            query=query,
            specification=spec_number or None,
            version=version,
            release=rel_int,
            clause=clause_id or None,
            limit=limit,
        )

    if action == "get_clause":
        result_id = str(args.get("result_id") or "").strip()
        if result_id:
            return search_service.get_clause(result_id=result_id)
        if not spec_number or not clause_id:
            return {"error": "Provide result_id OR (spec_number and clause_id)."}
        clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number, flags=re.IGNORECASE).strip()
        versions = search_db.get_versions_for_spec(clean_spec)
        if not versions and "." not in clean_spec and len(clean_spec) == 5:
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:]}"
            versions = search_db.get_versions_for_spec(clean_spec)
        if not versions:
            return {"error": f"Specification {spec_number} is not indexed in spec_search.db yet."}
        chosen_ver = version.lstrip("vV").strip() if version else str(versions[0]["version"])
        row = search_db.get_clause_content_by_spec_ver(clean_spec, chosen_ver, clause_id)
        if not row:
            return {"error": f"Clause {clause_id} not found in {clean_spec} v{chosen_ver}."}
        return dict(row)

    if action == "evolution_matrix":
        if not spec_number or not clause_id:
            return {"error": "spec_number and clause_id are required for action='evolution_matrix'."}
        clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number, flags=re.IGNORECASE).strip()
        docs = search_db.get_versions_for_spec(clean_spec)
        if not docs and "." not in clean_spec and len(clean_spec) == 5:
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:]}"
            docs = search_db.get_versions_for_spec(clean_spec)
        if not docs:
            return {"error": f"No indexed versions found for {spec_number} in spec_search.db."}

        def _ver_key(d: Dict[str, Any]) -> Tuple[int, ...]:
            parts = re.findall(r"\d+", str(d.get("version", "")))
            return tuple(int(x) for x in parts) if parts else (0,)

        docs_sorted = sorted(docs, key=_ver_key)
        timeline = []
        prev_text = None
        for d in docs_sorted:
            ver = str(d.get("version", ""))
            row = search_db.get_clause_content_by_spec_ver(clean_spec, ver, clause_id)
            content = str(row.get("content", "")) if isinstance(row, dict) else str(row or "")
            exists = bool(content and content.strip())
            norm = content.strip()
            status = "absent"
            if exists:
                if prev_text is None:
                    status = "introduced"
                elif norm != prev_text:
                    status = "modified"
                else:
                    status = "unchanged"
                prev_text = norm
            timeline.append({
                "version_id": d.get("id"),
                "version": ver,
                "release_date": d.get("release_date"),
                "status": status,
                "char_length": len(norm),
            })
        return {
            "spec_number": clean_spec,
            "clause_id": clause_id,
            "versions_checked": len(timeline),
            "timeline": timeline,
        }

    if action == "diff_versions":
        old_ver = str(args.get("old_version") or "").strip().lstrip("vV")
        new_ver = str(args.get("new_version") or version or "").strip().lstrip("vV")
        if not spec_number or not clause_id or not old_ver or not new_ver:
            return {"error": "spec_number, clause_id, old_version, and new_version are required for action='diff_versions'."}
        from modules.spec_search.core.spec_clause_diff import generate_unified_diff

        clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number, flags=re.IGNORECASE).strip()
        old_row = search_db.get_clause_content_by_spec_ver(clean_spec, old_ver, clause_id)
        if not old_row and "." not in clean_spec and len(clean_spec) == 5:
            clean_spec = f"{clean_spec[:2]}.{clean_spec[2:]}"
            old_row = search_db.get_clause_content_by_spec_ver(clean_spec, old_ver, clause_id)
        new_row = search_db.get_clause_content_by_spec_ver(clean_spec, new_ver, clause_id)

        old_text = str(old_row.get("content", "")) if isinstance(old_row, dict) else str(old_row or "")
        new_text = str(new_row.get("content", "")) if isinstance(new_row, dict) else str(new_row or "")
        if not old_text and not new_text:
            return {"error": f"Clause {clause_id} not found in either v{old_ver} or v{new_ver} of {spec_number}."}
        diff_text = generate_unified_diff(
            base_text=old_text,
            target_text=new_text,
            base_label=f"TS {clean_spec} v{old_ver} (Clause {clause_id})",
            target_label=f"TS {clean_spec} v{new_ver} (Clause {clause_id})",
        )
        added = sum(1 for ln in diff_text.splitlines() if ln.startswith("+") and not ln.startswith("+++"))
        removed = sum(1 for ln in diff_text.splitlines() if ln.startswith("-") and not ln.startswith("---"))
        sim = difflib.SequenceMatcher(None, old_text, new_text).ratio()
        return {
            "spec_number": clean_spec,
            "clause_id": clause_id,
            "old_version": old_ver,
            "new_version": new_ver,
            "added_lines": added,
            "removed_lines": removed,
            "similarity_ratio": round(sim, 4),
            "unified_diff": diff_text,
        }

    if action == "extract_clause":
        if not spec_number or not clause_id:
            return {"error": "spec_number and clause_id are required for action='extract_clause'."}
        from modules.spec_search.core.spec_clause_exporter import match_heading_for_clause
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown
        from modules.word_tools.core.word_excerpt_extractor import WordExcerptEngine, WordHeadingScanner

        parts = SpecDocResolver.resolve_spec_files(spec_number, version=version)
        if not parts:
            return {"error": f"No local Word files found for {spec_number}. Download the spec first."}

        scanner = WordHeadingScanner()
        headings = scanner.scan([str(p) for p in parts])
        target_clauses = _parse_csv_list(clause_id)
        matched_ids: List[str] = []
        matched_titles: List[str] = []
        for cid in target_clauses:
            h = match_heading_for_clause(headings, cid)
            if h:
                matched_ids.append(h.id)
                matched_titles.append(h.text)

        if not matched_ids:
            return {"error": f"Could not find heading for clause '{clause_id}' in {spec_number}."}

        out_dir = Path(args.get("output_dir") or (get_output_root() / "Clause_Excerpts"))
        out_dir.mkdir(parents=True, exist_ok=True)
        safe_clause = re.sub(r"[^A-Za-z0-9._-]+", "_", clause_id)
        out_docx = out_dir / f"{spec_number}_{version or 'latest'}_clause_{safe_clause}.docx"

        _disable_faulthandler_for_com()
        with _COM_LOCK:
            engine = WordExcerptEngine([str(p) for p in parts], headings)
            engine.extract(matched_ids, str(out_docx))

        out_md = out_docx.with_suffix(".md")
        md_res = convert_document_to_markdown(out_docx, out_md, revisions="accept")
        md_text = out_md.read_text(encoding="utf-8", errors="replace") if md_res and out_md.exists() else ""
        return {
            "spec_number": spec_number,
            "version": version,
            "clauses_extracted": matched_titles,
            "docx_path": str(out_docx),
            "markdown_path": str(out_md) if md_res else None,
            "markdown_preview": md_text[:4000],
        }

    return {"error": f"Unsupported action '{action}' for spec_text_and_clauses."}


# =========================================================================
# DOMAIN 3: Specification Change History & Clause-to-TDoc Attribution
# =========================================================================

def _handle_spec_history_and_attribution(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
    spec_search_db_path: Path,
    spec_history_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "query_history").strip().lower()
    spec_number = str(args.get("spec_number") or "").strip()
    version = str(args.get("version") or "").strip() or None
    clause_id = str(args.get("clause_id") or "").strip()
    query = str(args.get("query") or "").strip()
    limit = int(args.get("limit") or 25)

    from modules.specifications.core.spec_history_db import SpecHistoryDatabase
    history_db = SpecHistoryDatabase(spec_history_db_path)

    if action == "query_history":
        rows = history_db.search_history(
            spec_number=spec_number or None,
            query=query or None,
            meeting=args.get("meeting"),
            category=args.get("category"),
            version=version,
            limit=limit,
        )
        return {
            "spec_number": spec_number or None,
            "query": query or None,
            "count": len(rows),
            "records": rows,
        }

    if action == "ingest_history":
        if not spec_number:
            return {"error": "spec_number is required for action='ingest_history'."}
        from modules.specifications.core.spec_history_extractor import SpecHistoryExtractor

        doc_path = SpecHistoryExtractor.resolve_spec_document(spec_number, version=version)
        if not doc_path or not doc_path.exists():
            return {"error": f"No local Word document found for {spec_number}. Run spec_catalogue_and_files(action='download') first."}
        records = SpecHistoryExtractor.extract_from_docx(doc_path, spec_number, version_hint=version)
        inserted = history_db.upsert_history_records(records) if records else 0
        return {
            "spec_number": spec_number,
            "source_file": str(doc_path),
            "extracted_rows": len(records),
            "upserted_rows": inserted,
        }

    if action == "attribute_clause":
        if not spec_number or not clause_id:
            return {"error": "spec_number and clause_id are required for action='attribute_clause'."}
        from modules.spec_search.core.spec_search_db import SpecSearchDatabase
        from modules.specifications.core.spec_clause_attribution import (
            SpecClauseAttributionEngine,
            verify_candidate_in_docx,
        )
        from modules.specifications.core.database import SpecsDatabase

        search_db = SpecSearchDatabase(spec_search_db_path)
        specs_db = SpecsDatabase(specs_db_path)
        engine = SpecClauseAttributionEngine(
            history_db=history_db,
            specs_db=specs_db,
            spec_search_db=search_db,
        )
        report = engine.attribute_clause(
            spec_number=spec_number,
            clause_id=clause_id,
            clause_title=str(args.get("clause_title") or ""),
            clause_text=str(args.get("clause_text") or ""),
            target_version=version,
            include_all_versions=bool(args.get("include_all_versions", False)),
            limit=limit,
        )
        if bool(args.get("deep_verify_docx", False)) and report.candidates:
            for cand in report.candidates[: min(10, len(report.candidates))]:
                verify_candidate_in_docx(cand, clause_id, specs_db=specs_db)
            report.candidates.sort(key=lambda c: (-c.confidence_score, c.version, c.tdoc_id))

        return {
            "spec_number": report.spec_number,
            "clause_id": report.clause_id,
            "clause_title": report.clause_title,
            "first_added_version": report.first_added_version,
            "first_added_release": report.first_added_release,
            "origin_candidates": [
                {
                    "tdoc_id": c.tdoc_id,
                    "parent_tdoc": c.parent_tdoc,
                    "cr_number": c.cr_number,
                    "category": c.category,
                    "version": c.version,
                    "meeting": c.meeting,
                    "subject": c.subject,
                    "title": c.title,
                    "source_company": c.source_company,
                    "work_item": c.work_item,
                    "clauses_affected": c.clauses_affected,
                    "confidence_tier": c.confidence_tier,
                    "confidence_score": c.confidence_score,
                    "match_reasons": c.match_reasons,
                    "is_origin_candidate": c.is_origin_candidate,
                    "docx_verified": c.docx_verified,
                }
                for c in report.origin_candidates[:10]
            ],
            "all_candidates_count": len(report.candidates),
            "candidates": [
                {
                    "tdoc_id": c.tdoc_id,
                    "cr_number": c.cr_number,
                    "category": c.category,
                    "version": c.version,
                    "meeting": c.meeting,
                    "subject": c.subject,
                    "source_company": c.source_company,
                    "work_item": c.work_item,
                    "clauses_affected": c.clauses_affected,
                    "confidence_tier": c.confidence_tier,
                    "confidence_score": c.confidence_score,
                    "match_reasons": c.match_reasons,
                    "docx_verified": c.docx_verified,
                }
                for c in report.candidates[:limit]
            ],
        }

    if action == "rank_origin_crs":
        if not query:
            return {"error": "query is required for action='rank_origin_crs'."}
        from modules.spec_search.core.spec_search_db import SpecSearchDatabase
        from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
        from modules.specifications.core.database import SpecsDatabase

        search_db = SpecSearchDatabase(spec_search_db_path)
        specs_db = SpecsDatabase(specs_db_path)
        engine = SpecClauseAttributionEngine(
            history_db=history_db,
            specs_db=specs_db,
            spec_search_db=search_db,
        )
        all_imported = search_db.get_imported_versions()
        if spec_number:
            clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number, flags=re.IGNORECASE).strip()
            v_ids = [int(v["id"]) for v in all_imported if str(v.get("spec_number")) in (spec_number, clean_spec)]
        else:
            v_ids = [int(v["id"]) for v in all_imported]
        df_hits = search_db.search_substring(query, v_ids)
        hits_resp = df_hits.to_dict(orient="records") if hasattr(df_hits, "to_dict") else []
        ranked = engine.rank_origin_crs_for_query(
            search_query=query,
            clause_hits=hits_resp[: max(30, limit * 3)],
            max_clauses=int(args.get("max_clauses") or 12),
            top_n=limit,
        )
        return {
            "query": query,
            "spec_number": spec_number or None,
            "clause_hits_analyzed": len(hits_resp),
            "ranked_origin_crs": ranked,
        }

    if action == "tdoc_dossier":
        if not spec_number:
            return {"error": "spec_number is required for action='tdoc_dossier'."}
        from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

        rows = history_db.search_history(spec_number=spec_number, version=version, limit=500)
        tdoc_tokens: List[str] = []
        for r in rows:
            st = str(r.get("source_tdoc") or "").strip()
            if st and st != "-":
                tdoc_tokens.append(st)
            for pcr in r.get("pcr_tdocs") or []:
                if pcr and pcr != "-":
                    tdoc_tokens.append(str(pcr).strip())

        gen = SpecTDocReportGenerator(data_db_path=specs_db_path, history_db=history_db)
        resolved_items = gen.resolve_tdocs_metadata(tdoc_tokens)
        md_report = SpecTDocReportGenerator.generate_markdown_report(
            spec_number=spec_number,
            version=version or "all",
            tdoc_entries=resolved_items,
        )
        out_dir = Path(args.get("output_dir") or (get_output_root() / "Spec_TDoc_Dossiers"))
        out_dir.mkdir(parents=True, exist_ok=True)
        out_md = out_dir / f"{spec_number}_{version or 'all'}_tdoc_dossier.md"
        out_md.write_text(md_report, encoding="utf-8")
        return {
            "spec_number": spec_number,
            "version": version,
            "total_tdocs": len(resolved_items),
            "markdown_dossier_path": str(out_md),
            "preview_items": resolved_items[: min(15, len(resolved_items))],
        }

    return {"error": f"Unsupported action '{action}' for spec_history_and_attribution."}


# =========================================================================
# DOMAIN 4: Protocol Inspection & Evolution Matrix (NAS, RRC, NGAP, GTPv1-U, UP, PFCP)
# =========================================================================

def _handle_protocol_inspector(
    args: Dict[str, Any],
    *,
    protocol_service: ProtocolKnowledgeService,
    protocol_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "list_protocols").strip().lower()
    protocol = str(args.get("protocol") or "").strip().upper()
    query = str(args.get("query") or "").strip()
    version = str(args.get("version") or "").strip() or None
    limit = int(args.get("limit") or 20)

    if action == "list_protocols":
        return protocol_service.list_protocols()

    if action == "list_messages":
        if not protocol:
            return {"error": "protocol is required for action='list_messages'."}
        return protocol_service.list_messages(
            protocol=protocol,
            version=version,
            limit=limit,
        )

    if action == "find_message":
        if not query:
            return {"error": "query (message name) is required for action='find_message'."}
        res = protocol_service.find_message(message=query, protocol=protocol or None, version=version)
        if isinstance(res.data, dict) and "fields" in res.data:
            res.data.setdefault("information_elements", res.data["fields"])
        return res

    if action == "find_ie":
        if not query:
            return {"error": "query (IE name) is required for action='find_ie'."}
        return protocol_service.find_ie(ie=query, protocol=protocol or None, version=version)

    if action == "evolution_matrix":
        if not protocol or not query:
            return {"error": "protocol and query (message or IE name) are required for action='evolution_matrix'."}
        from modules.nas.core.nas_db import NASDatabase
        from modules.nas.core.nas_exporter import build_matrix_markdown, build_matrix_tsv

        nas_db = NASDatabase(protocol_db_path)
        target_type = str(args.get("target_type") or "message").strip().lower()
        fmt = str(args.get("format") or "markdown").strip().lower()

        if target_type == "ie":
            df = nas_db.get_ie_definition_evolution_df(query, protocol=protocol)
            title = f"{protocol} IE Evolution: {query}"
        else:
            df = nas_db.get_message_evolution_df(query, protocol=protocol)
            title = f"{protocol} Message Evolution: {query}"

        if df is None or df.empty:
            return {"error": f"No evolution matrix rows found for {protocol} {target_type} '{query}'."}

        rendered = build_matrix_tsv(title, df) if fmt == "tsv" else build_matrix_markdown(title, df)
        return {
            "protocol": protocol,
            "target_type": target_type,
            "target_name": query,
            "format": fmt,
            "row_count": len(df),
            "columns": list(df.columns),
            "matrix": rendered,
        }

    return {"error": f"Unsupported action '{action}' for protocol_inspector."}


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
        mtg = None
        if meeting_number:
            rows = mdb.search_meetings(query=meeting_number, wg=wg)
            mtg = rows[0] if rows else None
        if not mtg:
            rows = mdb.search_meetings(wg=wg)
            mtg = rows[0] if rows else {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        main_ftp = docs_ftp.rsplit("/Docs", 1)[0] if "/Docs" in docs_ftp else docs_ftp
        mtg_folder = str(mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        meeting_dir = get_meetings_root() / wg / mtg_folder
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
        mtg = None
        if meeting_number:
            rows = mdb.search_meetings(query=meeting_number, wg=wg)
            mtg = rows[0] if rows else None
        if not mtg:
            rows = mdb.search_meetings(wg=wg)
            mtg = rows[0] if rows else {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        main_ftp = docs_ftp.rsplit("/Docs", 1)[0] if "/Docs" in docs_ftp else docs_ftp
        mtg_folder = str(mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        chair_dir = get_meetings_root() / wg / mtg_folder / "Chair_Notes"
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
        mtg = None
        if meeting_number:
            rows = mdb.search_meetings(query=meeting_number, wg=wg)
            mtg = rows[0] if rows else None
        if not mtg:
            rows = mdb.search_meetings(wg=wg)
            mtg = rows[0] if rows else {}

        docs_ftp = str(mtg.get("docs_folder_url") or "")
        rev_url = str(mtg.get("revisions_url") or "")
        mtg_folder = str(mtg.get("meeting_folder_name") or f"{wg}_{meeting_number or 'Live'}")
        meeting_dir = get_meetings_root() / wg / mtg_folder
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


# =========================================================================
# DOMAIN 6: Meetings, TDocs, Triage Stances & Dossiers
# =========================================================================

def _handle_meetings_and_tdocs(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
    user_tdocs_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search_tdocs").strip().lower()
    wg = str(args.get("wg") or "").strip().upper() or None
    meeting_number = str(args.get("meeting_number") or "").strip()
    tdoc_id = str(args.get("tdoc_id") or "").strip()
    query = str(args.get("query") or "").strip()
    limit = int(args.get("limit") or 25)

    from modules.meetings.core.meetings_db import MeetingsDatabase
    from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase

    mdb = MeetingsDatabase(specs_db_path)

    if action == "search_meetings":
        rows = mdb.search_meetings(
            query=query or meeting_number or None,
            wg=wg,
            year=args.get("year"),
        )
        return {
            "count": len(rows[:limit]),
            "meetings": rows[:limit],
        }

    if action == "search_tdocs":
        source_company = str(args.get("source_company") or "").strip()
        confirmed_only = bool(args.get("confirmed_company_only", True))
        with mdb._get_connection() as conn:
            sql = """
                SELECT t.tdoc_number, t.title, t.source, t.doc_type, t.doc_for,
                       t.agenda_item, t.agenda_item_description, t.work_item,
                       t.spec, t.release, t.tdoc_status, t.secretary_remarks,
                       t.revision_of, t.revised_to, m.wg_name, m.meeting_number, m.meeting_folder_name
                FROM tdocs t
                LEFT JOIN meetings m ON t.meeting_id = m.id
                WHERE 1=1
            """
            params: List[Any] = []
            if tdoc_id:
                sql += " AND UPPER(t.tdoc_number) LIKE ?"
                params.append(f"%{tdoc_id.upper()}%")
            if wg:
                sql += " AND UPPER(m.wg_name) = ?"
                params.append(wg)
            if meeting_number:
                sql += " AND (m.meeting_number LIKE ? OR m.meeting_folder_name LIKE ?)"
                params.extend([f"%{meeting_number}%", f"%{meeting_number}%"])
            if args.get("agenda_item"):
                sql += " AND t.agenda_item LIKE ?"
                params.append(f"%{str(args['agenda_item']).strip()}%")
            if args.get("work_item"):
                sql += " AND t.work_item LIKE ?"
                params.append(f"%{str(args['work_item']).strip()}%")
            if args.get("status"):
                sql += " AND t.tdoc_status LIKE ?"
                params.append(f"%{str(args['status']).strip()}%")
            if query:
                sql += " AND (t.title LIKE ? OR t.tdoc_number LIKE ? OR t.work_item LIKE ? OR t.source LIKE ?)"
                params.extend([f"%{query}%"] * 4)
            if source_company:
                sql += " AND t.source LIKE ?"
                params.append(f"%{source_company}%")

            sql += " ORDER BY t.id DESC LIMIT ?"
            params.append(max(limit * 4, 100) if source_company else limit)
            raw_rows = [dict(r) for r in conn.execute(sql, params).fetchall()]

        if source_company:
            filtered = []
            target_canon = CompanySanitizer.sanitize(source_company).lower()
            for r in raw_rows:
                raw_src = str(r.get("source") or "")
                if confirmed_only:
                    contributors = [c.lower() for c in CompanySanitizer.get_matching_contributors(raw_src)]
                    if any(target_canon in c or source_company.lower() in c for c in contributors):
                        filtered.append(r)
                else:
                    if source_company.lower() in raw_src.lower():
                        filtered.append(r)
            raw_rows = filtered[:limit]

        user_db = TDocsDatabase(user_tdocs_db_path)
        user_map = user_db.get_all_user_data()
        for r in raw_rows:
            tn = str(r.get("tdoc_number") or "")
            if tn in user_map:
                r["my_status"] = user_map[tn].get("my_status", "")
                r["my_notes"] = user_map[tn].get("my_notes", "")

        return {
            "count": len(raw_rows[:limit]),
            "tdocs": raw_rows[:limit],
        }

    if action == "read_tdoc_markdown":
        if not tdoc_id:
            return {"error": "tdoc_id is required for action='read_tdoc_markdown'."}
        from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        with mdb._get_connection() as conn:
            row = conn.execute(
                """
                SELECT t.*, m.wg_name, m.meeting_folder_name, m.docs_folder_url, m.revisions_url
                FROM tdocs t
                LEFT JOIN meetings m ON t.meeting_id = m.id
                WHERE UPPER(t.tdoc_number) = ?
                ORDER BY t.id DESC LIMIT 1
                """,
                (tdoc_id.upper(),),
            ).fetchone()
        meta = dict(row) if row else {}
        eff_wg = wg or str(meta.get("wg_name") or "SA2")
        mtg_folder = str(meta.get("meeting_folder_name") or f"{eff_wg}_{meeting_number or 'Cache'}")
        docs_ftp = str(meta.get("docs_folder_url") or "")
        rev_url = str(meta.get("revisions_url") or "")
        meeting_dir = get_meetings_root() / eff_wg / mtg_folder

        files, err = TDocFileHandler.download_and_extract_tdoc(
            tdoc_id=tdoc_id,
            docs_ftp_url=docs_ftp,
            meeting_dir=meeting_dir,
            wg_name=eff_wg,
            revisions_url=rev_url,
        )
        if err or not files:
            return {"error": err or f"Could not locate or download TDoc {tdoc_id}."}

        target_doc = next((f for f in files if f.suffix.lower() in (".docx", ".doc")), files[0])
        if target_doc.suffix.lower() == ".doc":
            from modules.word_tools.core.word_converter import convert_doc_to_docx
            conv = convert_doc_to_docx(str(target_doc))
            if conv and Path(conv).exists():
                target_doc = Path(conv)

        rev_mode = str(args.get("revisions") or "markup").strip().lower()
        md_path = target_doc.with_suffix(".md")
        convert_document_to_markdown(target_doc, md_path, revisions=rev_mode)
        md_content = md_path.read_text(encoding="utf-8", errors="replace") if md_path.exists() else ""
        max_chars = int(args.get("max_chars") or 12000)
        return {
            "tdoc_id": tdoc_id,
            "metadata": {
                k: meta.get(k)
                for k in ("tdoc_number", "title", "source", "doc_type", "agenda_item", "work_item", "tdoc_status", "secretary_remarks")
                if k in meta
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
        user_db = TDocsDatabase(user_tdocs_db_path)
        stance = args.get("stance")
        notes = args.get("notes")
        if stance is not None or notes is not None:
            existing = user_db.get_all_user_data().get(tdoc_id, {})
            new_status = str(stance if stance is not None else existing.get("my_status", "")).strip()
            new_notes = str(notes if notes is not None else existing.get("my_notes", "")).strip()
            user_db.update_tdoc_user_data(tdoc_id, new_status, new_notes)
        data = user_db.get_all_user_data().get(tdoc_id, {"my_status": "", "my_notes": ""})
        return {
            "tdoc_id": tdoc_id,
            "my_status": data.get("my_status", ""),
            "my_notes": data.get("my_notes", ""),
        }

    if action == "meeting_stats":
        with mdb._get_connection() as conn:
            sql = """
                SELECT t.tdoc_number, t.source, t.doc_type, t.agenda_item, t.work_item, t.tdoc_status
                FROM tdocs t
                JOIN meetings m ON t.meeting_id = m.id
                WHERE 1=1
            """
            params = []
            if wg:
                sql += " AND UPPER(m.wg_name) = ?"
                params.append(wg)
            if meeting_number:
                sql += " AND (m.meeting_number LIKE ? OR m.meeting_folder_name LIKE ?)"
                params.extend([f"%{meeting_number}%", f"%{meeting_number}%"])
            rows = [dict(r) for r in conn.execute(sql, params).fetchall()]

        if not rows:
            return {"error": "No TDocs found for the specified meeting filter."}

        company_counts: Dict[str, int] = {}
        status_counts: Dict[str, int] = {}
        wi_counts: Dict[str, int] = {}
        for r in rows:
            st = str(r.get("tdoc_status") or "Unknown").strip()
            status_counts[st] = status_counts.get(st, 0) + 1
            for comp in CompanySanitizer.get_matching_contributors(str(r.get("source") or "")):
                company_counts[comp] = company_counts.get(comp, 0) + 1
            for wi in _parse_csv_list(r.get("work_item")):
                wi_counts[wi] = wi_counts.get(wi, 0) + 1

        top_companies = sorted(company_counts.items(), key=lambda x: -x[1])[:20]
        top_wis = sorted(wi_counts.items(), key=lambda x: -x[1])[:20]
        return {
            "wg": wg,
            "meeting_number": meeting_number,
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


# =========================================================================
# DOMAIN 8: Word ⇄ Markdown Studio, Clause Slicing & Visio Extraction
# =========================================================================

def _handle_word_markdown_studio(args: Dict[str, Any]) -> Dict[str, Any]:
    action = str(args.get("action") or "word_to_md").strip().lower()
    input_path = Path(str(args.get("input_path") or "").strip())
    output_path_raw = str(args.get("output_path") or "").strip()

    if not input_path.exists():
        return {"error": f"input_path does not exist: {input_path}"}

    if action == "word_to_md":
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        src = input_path
        if src.suffix.lower() == ".doc":
            from modules.word_tools.core.word_converter import convert_doc_to_docx
            conv = convert_doc_to_docx(str(src))
            if conv and Path(conv).exists():
                src = Path(conv)
        out_md = Path(output_path_raw) if output_path_raw else src.with_suffix(".md")
        rev_mode = str(args.get("revisions") or "markup").strip().lower()
        res = convert_document_to_markdown(src, out_md, revisions=rev_mode)
        if not res or not out_md.exists():
            return {"error": f"Failed converting {src} to Markdown."}
        text = out_md.read_text(encoding="utf-8", errors="replace")
        return {
            "input_path": str(input_path),
            "output_path": str(out_md),
            "revisions_mode": rev_mode,
            "media_dir": str(res.media_dir) if getattr(res, "media_dir", None) else None,
            "char_length": len(text),
            "preview": text[:4000],
        }

    if action == "md_to_word":
        from modules.word_tools.core.markdown_to_docx import convert_markdown_to_docx

        out_docx = Path(output_path_raw) if output_path_raw else input_path.with_suffix(".docx")
        _disable_faulthandler_for_com()
        with _COM_LOCK:
            res = convert_markdown_to_docx(
                md_path=input_path,
                output_docx_path=out_docx,
                apply_3gpp_styles=bool(args.get("apply_3gpp_styles", True)),
                render_diagrams=bool(args.get("compile_plantuml", True)),
            )
        stats = res.get("stats", {}) if isinstance(res, dict) else {}
        return {
            "input_path": str(input_path),
            "output_path": str(res.get("output_path", out_docx)),
            "stats": stats,
        }

    if action == "slice_clauses":
        from modules.spec_search.core.spec_clause_exporter import match_heading_for_clause
        from modules.word_tools.core.word_excerpt_extractor import WordExcerptEngine, WordHeadingScanner

        clauses = _parse_csv_list(args.get("clauses"))
        scanner = WordHeadingScanner()
        headings = scanner.scan([str(input_path)])
        if not clauses:
            return {
                "input_path": str(input_path),
                "headings_count": len(headings),
                "headings": [{"id": h.id, "level": h.level, "text": h.text} for h in headings[:200]],
            }
        matched_ids = []
        matched_texts = []
        for c in clauses:
            h = match_heading_for_clause(headings, c)
            if h:
                matched_ids.append(h.id)
                matched_texts.append(h.text)
        if not matched_ids:
            return {"error": f"No matching headings found for clauses {clauses} in {input_path.name}."}
        out_docx = Path(output_path_raw) if output_path_raw else input_path.with_name(f"{input_path.stem}_excerpt.docx")
        _disable_faulthandler_for_com()
        with _COM_LOCK:
            engine = WordExcerptEngine([str(input_path)], headings)
            engine.extract(matched_ids, str(out_docx))
        return {
            "input_path": str(input_path),
            "output_path": str(out_docx),
            "clauses_extracted": matched_texts,
        }

    if action == "extract_visio":
        from modules.word_tools.core.word_extractor import WordExtractorThread

        _disable_faulthandler_for_com()
        with _COM_LOCK:
            worker = WordExtractorThread(str(input_path))
            worker.run()
        extracted = [str(p) for p in sorted(input_path.parent.glob(f"{input_path.stem}*.vsdx"))]
        return {
            "input_path": str(input_path),
            "output_dir": str(input_path.parent),
            "extracted_count": len(extracted),
            "files": extracted,
        }

    return {"error": f"Unsupported action '{action}' for word_markdown_studio."}


# =========================================================================
# DOMAIN 9: 3GPP CR & Pseudo-CR (pCR) Contribution Generator
# =========================================================================

def _handle_cr_pcr_generator(args: Dict[str, Any]) -> Dict[str, Any]:
    action = str(args.get("action") or "generate_pcr").strip().lower()
    from modules.word_tools.core.cr_pcr_generator import ContributionMetadata, CrPcrGeneratorEngine
    from modules.word_tools.core.word_excerpt_naming import generate_cr_pcr_export_name

    mode = "cr" if action == "generate_cr" or str(args.get("mode", "")).lower() == "cr" else "pcr"
    spec_number = str(args.get("spec_number") or "23.501").strip()
    spec_version = str(args.get("version") or "19.0.0").strip()
    clauses = _parse_csv_list(args.get("clauses"))
    tdoc_id = str(args.get("tdoc_id") or "S2-26xxxx").strip()
    title = str(args.get("title") or "Proposed clarification").strip()
    source_company = str(args.get("source_company") or "Deutsche Telekom").strip()
    work_item = str(args.get("work_item") or "TEI19").strip()

    if action == "preview_filename":
        proposed_filename = generate_cr_pcr_export_name(
            mode=mode,
            spec_num=spec_number,
            version=spec_version,
            tdoc_id=tdoc_id,
            work_item=work_item,
            title=title,
        )
        return {"proposed_filename": proposed_filename}

    source_doc = str(args.get("source_docx_path") or "").strip()
    source_paths: List[str] = []
    if source_doc and Path(source_doc).exists():
        source_paths = [source_doc]
    else:
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        resolved = SpecDocResolver.resolve_spec_files(spec_number, version=spec_version)
        source_paths = [str(p) for p in resolved]

    if not source_paths:
        return {"error": f"No baseline Word document found for {spec_number} v{spec_version}. Provide source_docx_path or download the spec first."}

    from modules.spec_search.core.spec_clause_exporter import match_heading_for_clause
    from modules.word_tools.core.word_excerpt_extractor import WordHeadingScanner

    scanner = WordHeadingScanner()
    headings = scanner.scan(source_paths)
    matched_ids: List[str] = []
    for c in clauses:
        h = match_heading_for_clause(headings, c)
        if h:
            matched_ids.append(h.id)

    if not matched_ids and headings:
        matched_ids = [headings[0].id]

    proposed_filename = generate_cr_pcr_export_name(
        mode=mode,
        spec_num=spec_number,
        version=spec_version,
        tdoc_id=tdoc_id,
        work_item=work_item,
        title=title,
        headings=headings,
        selected_ids=matched_ids,
    )

    out_dir = Path(args.get("output_dir") or (get_output_root() / "Contributions"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = Path(args.get("output_path") or (out_dir / proposed_filename))

    meta = ContributionMetadata(
        mode=mode,
        meeting_wg=str(args.get("wg") or "SA2"),
        meeting_num=str(args.get("meeting_number") or "174"),
        meeting_dates=str(args.get("meeting_dates") or ""),
        meeting_location=str(args.get("meeting_location") or ""),
        tdoc_id=tdoc_id,
        title=title,
        source_company=source_company,
        agenda_item=str(args.get("agenda_item") or ""),
        work_item=work_item,
        spec_num=spec_number,
        spec_version=spec_version,
        spec_release=str(args.get("release") or "Rel-19"),
        cr_num=str(args.get("cr_number") or "<CR#>"),
        cr_rev=str(args.get("cr_revision") or "-"),
        category=str(args.get("category") or "F"),
        reason_for_change=str(args.get("reason_for_change") or ""),
        summary_of_change=str(args.get("summary_of_change") or ""),
        consequences=str(args.get("consequences_if_not_approved") or ""),
        clauses_affected=", ".join(clauses),
    )

    _disable_faulthandler_for_com()
    with _COM_LOCK:
        engine = CrPcrGeneratorEngine(
            metadata=meta,
            source_paths=source_paths,
            headings=headings,
            selected_ids=matched_ids,
        )
        final_path = engine.generate(str(out_path))

    return {
        "mode": mode.upper(),
        "spec_number": spec_number,
        "spec_version": spec_version,
        "clauses": clauses,
        "output_path": str(final_path),
    }


# =========================================================================
# DOMAIN 10: Word Hygiene Sanitizer (F-07), Track Changes Compare & MIP Labels
# =========================================================================

def _handle_word_sanitizer_and_compliance(args: Dict[str, Any]) -> Dict[str, Any]:
    action = str(args.get("action") or "audit_hygiene").strip().lower()
    input_path = Path(str(args.get("input_path") or "").strip())

    if not input_path.exists():
        return {"error": f"input_path does not exist: {input_path}"}

    if action == "audit_hygiene":
        from modules.word_tools.core.document_sanitizer import inspect_document

        report = inspect_document(input_path)
        return {
            "file_path": str(input_path),
            "primary_language": report.language_stats.primary_language,
            "language_counts": report.language_stats.language_counts,
            "language_change_marks": report.language_stats.language_change_marks,
            "pure_language_change_marks": report.language_stats.pure_language_change_marks,
            "total_insertions": report.total_insertions,
            "total_deletions": report.total_deletions,
            "total_formatting_revisions": report.total_formatting,
            "total_comments": report.total_comments,
            "doc_creator": report.doc_creator,
            "doc_last_modified_by": report.doc_last_modified_by,
            "authors": [
                {
                    "name": a.name,
                    "insertions": a.insertions,
                    "deletions": a.deletions,
                    "formatting": a.formatting,
                    "comments": a.comments,
                }
                for a in report.authors
            ],
        }

    if action == "sanitize_document":
        from modules.word_tools.core.document_sanitizer import SanitizeConfig, inspect_document, sanitize_document

        target_author = str(args.get("target_author") or "").strip()
        author_map: Dict[str, str] = {}
        if target_author:
            audit = inspect_document(input_path)
            author_map = {a.name: target_author for a in audit.authors}

        out_path = Path(str(args.get("output_path") or input_path).strip())
        cfg = SanitizeConfig(
            unify_language=bool(args.get("unify_language", True)),
            target_language=str(args.get("target_language") or "en-GB"),
            purge_language_changes=bool(args.get("purge_lang_revisions", True)),
            update_authors=bool(target_author),
            author_mapping=author_map,
            update_metadata=bool(target_author),
            backup=bool(args.get("create_backup", True)),
        )
        ok, mod_count, msg = sanitize_document(input_path, out_path, cfg)
        return {
            "success": ok,
            "file_path": str(out_path),
            "modifications_count": mod_count,
            "message": msg,
        }

    if action == "compare_documents":
        revised_path = Path(str(args.get("revised_path") or "").strip())
        if not revised_path.exists():
            return {"error": "Valid revised_path is required for action='compare_documents'."}
        from modules.word_tools.core.word_comparator import WordComparatorThread

        _disable_faulthandler_for_com()
        with _COM_LOCK:
            worker = WordComparatorThread(
                doc_a=str(input_path),
                doc_b=str(revised_path),
                keep_open=bool(args.get("keep_open", False)),
            )
            worker.run()
        return {
            "original_path": str(input_path),
            "revised_path": str(revised_path),
            "status": "Comparison completed in Microsoft Word.",
        }

    if action == "inspect_label":
        from modules.word_tools.core.sensitivity_label import get_document_sensitivity_label

        info = get_document_sensitivity_label(input_path)
        return {
            "file_path": str(input_path),
            "sensitivity_label": info,
        }

    if action == "apply_label":
        from modules.word_tools.core.sensitivity_label import batch_set_sensitivity_label
        from modules.word_tools.core.word_config import WordConfig

        preset_name = str(args.get("label_preset") or "").strip()
        label_id = None
        label_name = None
        if preset_name:
            for p in WordConfig.get_sensitivity_presets():
                if p.get("name", "").lower() == preset_name.lower():
                    label_id = p.get("id")
                    label_name = p.get("name")
                    break
        _disable_faulthandler_for_com()
        with _COM_LOCK:
            summary = batch_set_sensitivity_label(
                file_paths=[input_path],
                label_id=label_id,
                label_name=label_name,
                overwrite_existing=True,
            )
        return {
            "file_path": str(input_path),
            "summary": summary,
        }

    return {"error": f"Unsupported action '{action}' for word_sanitizer_and_compliance."}


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

def _handle_work_items_and_emails(
    args: Dict[str, Any],
    *,
    specs_db_path: Path,
) -> Dict[str, Any]:
    action = str(args.get("action") or "search_work_items").strip().lower()
    query = str(args.get("query") or "").strip()
    wg = str(args.get("wg") or "").strip().upper() or None
    release = str(args.get("release") or "").strip() or None
    limit = int(args.get("limit") or 25)

    from modules.work_items.core.wi_database import WorkItemsDatabase
    from modules.work_items.core.wi_settings import WorkItemsSettings

    wi_db = WorkItemsDatabase(specs_db_path)
    wi_settings = WorkItemsSettings()

    if action == "search_work_items":
        starred, following = wi_settings.get_tracking_sets()
        rows = wi_db.search_work_items(
            search_text=query,
            release=release or "All Releases",
            wg=wg or "All WGs",
        )
        if bool(args.get("tracked_only", False)):
            tracked_codes = starred | following
            rows = [r for r in rows if str(r.get("code", "")) in tracked_codes or str(r.get("acronym", "")) in tracked_codes]
        for r in rows[:limit]:
            c = str(r.get("code") or "")
            r["starred"] = c in starred
            r["following"] = c in following
        return {
            "count": len(rows[:limit]),
            "starred_codes": sorted(starred),
            "following_codes": sorted(following),
            "work_items": rows[:limit],
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

        email_db = EmailDatabase(specs_db_path)
        gen_db = GeneralEmailDatabase(specs_db_path)
        results = []
        with email_db._get_connection() as conn:
            try:
                sql = """
                    SELECT id, tdoc_number, subject, sender_name, sender_company, sent_date, substr(body_clean, 1, 800) AS snippet
                    FROM meeting_emails
                    WHERE (? = '' OR tdoc_number LIKE ? OR subject LIKE ? OR body_clean LIKE ? OR sender_company LIKE ?)
                    ORDER BY sent_date DESC LIMIT ?
                """
                q_like = f"%{query}%"
                rows = conn.execute(sql, (query, q_like, q_like, q_like, q_like, limit)).fetchall()
                results.extend([dict(r) for r in rows])
            except Exception:
                pass
        with gen_db._get_connection() as conn:
            try:
                sql = """
                    SELECT id, wg_name, subject, sender_name, sender_company, sent_date, substr(body_clean, 1, 800) AS snippet
                    FROM general_wg_emails
                    WHERE (? = '' OR subject LIKE ? OR body_clean LIKE ? OR sender_company LIKE ?)
                    ORDER BY sent_date DESC LIMIT ?
                """
                q_like = f"%{query}%"
                rows = conn.execute(sql, (query, q_like, q_like, q_like, limit)).fetchall()
                results.extend([dict(r) for r in rows])
            except Exception:
                pass
        return {
            "query": query,
            "count": len(results[:limit]),
            "emails": results[:limit],
        }

    return {"error": f"Unsupported action '{action}' for work_items_and_emails."}


# =========================================================================
# REGISTRY BUILDER (Compact 12-Tool Schema + Profile Filtering)
# =========================================================================

PROFILE_TOOLS: Dict[str, Tuple[str, ...]] = {
    "all": (
        "spec_catalogue_and_files",
        "spec_text_and_clauses",
        "spec_history_and_attribution",
        "protocol_inspector",
        "onsite_meeting_cockpit",
        "meetings_and_tdocs",
        "chair_notes_and_session_plan",
        "word_markdown_studio",
        "cr_pcr_generator",
        "word_sanitizer_and_compliance",
        "plantuml_office_studio",
        "work_items_and_emails",
    ),
    "knowledge": (
        "spec_catalogue_and_files",
        "spec_text_and_clauses",
        "spec_history_and_attribution",
        "protocol_inspector",
        "work_items_and_emails",
    ),
    "onsite": (
        "onsite_meeting_cockpit",
        "meetings_and_tdocs",
        "chair_notes_and_session_plan",
        "spec_text_and_clauses",
        "cr_pcr_generator",
    ),
    "authoring": (
        "word_markdown_studio",
        "cr_pcr_generator",
        "word_sanitizer_and_compliance",
        "plantuml_office_studio",
        "spec_text_and_clauses",
    ),
}


def build_external_agent_tool_registry(
    *,
    profile: str = "all",
    read_only: bool = False,
    specs_db_path: Optional[Path] = None,
    spec_search_db_path: Optional[Path] = None,
    protocol_db_path: Optional[Path] = None,
    spec_history_db_path: Optional[Path] = None,
    user_tdocs_db_path: Optional[Path] = None,
    session_plan_db_path: Optional[Path] = None,
) -> ToolRegistry:
    """Builds the token-efficient 12-tool external agent registry for Claude CLI / Codex CLI."""
    paths = _resolve_db_paths(
        specs_db_path=specs_db_path,
        spec_search_db_path=spec_search_db_path,
        protocol_db_path=protocol_db_path,
        spec_history_db_path=spec_history_db_path,
        user_tdocs_db_path=user_tdocs_db_path,
        session_plan_db_path=session_plan_db_path,
    )
    spec_service = SpecificationKnowledgeService(paths["specs_db"])
    search_service = SpecSearchKnowledgeService(paths["spec_search_db"])
    protocol_service = ProtocolKnowledgeService(paths["protocol_db"])

    all_defs: List[ToolDefinition] = [
        ToolDefinition(
            name="spec_catalogue_and_files",
            description="Search 3GPP specs/versions, download/resolve FTP archives, index for FTS5, or convert specs to Markdown.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search", "get", "resolve_files", "download", "index_for_search", "convert_to_md")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("version", "string", "e.g. '19.2.0'"),
                ToolArgumentSpec("release", "string", "e.g. 'Rel-19'"),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("mode", "string", "'single'|'split'"),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_catalogue_and_files(a, spec_service=spec_service, spec_search_db_path=paths["spec_search_db"])),
        ),
        ToolDefinition(
            name="spec_text_and_clauses",
            description="FTS5 search across 3GPP specs, get clause text, trace evolution across releases, diff versions, or extract clauses to .docx/.md.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_text", "get_clause", "evolution_matrix", "diff_versions", "extract_clause")),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("clause_id", "string", "e.g. '5.5.1'"),
                ToolArgumentSpec("result_id", "string", "Handle from search_text"),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("old_version", "string", ""),
                ToolArgumentSpec("new_version", "string", ""),
                ToolArgumentSpec("release", "string", ""),
                ToolArgumentSpec("latest_only", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_text_and_clauses(a, search_service=search_service, spec_search_db_path=paths["spec_search_db"])),
        ),
        ToolDefinition(
            name="spec_history_and_attribution",
            description="Query Change History tables, attribute clauses to origin CR/pCR/TDoc, rank origin CRs, or generate TDoc dossiers.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("query_history", "ingest_history", "attribute_clause", "rank_origin_crs", "tdoc_dossier")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("clause_id", "string", "e.g. '5.34.2'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("meeting", "string", ""),
                ToolArgumentSpec("deep_verify_docx", "boolean", "Verify CR .docx cover"),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_history_and_attribution(
                a,
                specs_db_path=paths["specs_db"],
                spec_search_db_path=paths["spec_search_db"],
                spec_history_db_path=paths["spec_history_db"],
            )),
        ),
        ToolDefinition(
            name="protocol_inspector",
            description="Inspect 3GPP protocols (5GS/EPS NAS, UE Policies, NR/LTE RRC, NGAP, GTPv1-U, UP, PFCP): messages, IEs, and evolution matrices.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("list_protocols", "list_messages", "find_message", "find_ie", "evolution_matrix")),
                ToolArgumentSpec("protocol", "string", "e.g. '5GS_NAS', 'RRC_NR', 'NGAP'"),
                ToolArgumentSpec("query", "string", "Message or IE name"),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("target_type", "string", "'message'|'ie'"),
                ToolArgumentSpec("format", "string", "'markdown'|'tsv'"),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_protocol_inspector(a, protocol_service=protocol_service, protocol_db_path=paths["protocol_db"])),
        ),
        ToolDefinition(
            name="onsite_meeting_cockpit",
            description="On-site 10.10.10.10 LAN (3GPPWIFI) & live meeting cockpit: probe LAN, browse 10.10.10.10/ftp/, fetch/diff live r01/r02 revisions, sync Chair Notes, check venue schedule & free delegates, or batch-cache TDocs.",
            permission=ToolPermission.NETWORK_FETCH,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("network_status", "browse_lan_server", "fetch_live_revisions", "sync_chair_notes", "where_should_we_go_now", "batch_cache_tdocs")),
                ToolArgumentSpec("wg", "string", "Default 'SA2'"),
                ToolArgumentSpec("meeting_number", "string", "e.g. '174'"),
                ToolArgumentSpec("tdoc_id", "string", "TDoc ID(s)"),
                ToolArgumentSpec("subfolder", "string", "e.g. 'Inbox/Revisions', 'Docs'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("day_name", "string", ""),
                ToolArgumentSpec("time_hhmm", "string", ""),
                ToolArgumentSpec("force_probe", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_onsite_meeting_cockpit(a, specs_db_path=paths["specs_db"], session_plan_db_path=paths["session_plan_db"])),
        ),
        ToolDefinition(
            name="meetings_and_tdocs",
            description="Search meetings & TDocs (disambiguating confirmed vs '?'/'[...]' tentative companies), read TDocs as Markdown, manage stances/notes, or get meeting stats.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_meetings", "search_tdocs", "read_tdoc_markdown", "get_or_set_stance", "meeting_stats")),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("meeting_number", "string", ""),
                ToolArgumentSpec("tdoc_id", "string", ""),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("source_company", "string", ""),
                ToolArgumentSpec("confirmed_company_only", "boolean", "Exclude '?'/'[...]'"),
                ToolArgumentSpec("agenda_item", "string", ""),
                ToolArgumentSpec("work_item", "string", ""),
                ToolArgumentSpec("stance", "string", ""),
                ToolArgumentSpec("notes", "string", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_meetings_and_tdocs(a, specs_db_path=paths["specs_db"], user_tdocs_db_path=paths["user_tdocs_db"])),
        ),
        ToolDefinition(
            name="chair_notes_and_session_plan",
            description="Parse Chairman's Notes .docx for TDoc outcomes, or ingest, query, and diff SA2 Session Plan snapshots (TU 1/TU 2 sub-slots & rooms).",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("parse_chair_notes_file", "ingest_session_plan", "query_session_plan", "diff_session_plans")),
                ToolArgumentSpec("meeting_number", "string", ""),
                ToolArgumentSpec("file_path", "string", ""),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("day_name", "string", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_chair_notes_and_session_plan(a, session_plan_db_path=paths["session_plan_db"])),
        ),
        ToolDefinition(
            name="word_markdown_studio",
            description="Bidirectional Word (.docx/.doc) ⇄ Markdown (.md) conversion (with Track Changes, tables, PlantUML & figures), clause slicing, or Visio (.vsdx) extraction.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("word_to_md", "md_to_word", "slice_clauses", "extract_visio")),
                ToolArgumentSpec("input_path", "string", "", required=True),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("revisions", "string", "'markup'|'accept'|'reject'"),
                ToolArgumentSpec("clauses", "string", "Comma-separated"),
            ),
            handler=_wrap_handler(_handle_word_markdown_studio),
        ),
        ToolDefinition(
            name="cr_pcr_generator",
            description="Generate official MCC-styled 3GPP CR or pCR .docx contributions with cover sheet, blue change separators, and differentiated naming.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("generate_pcr", "generate_cr", "preview_filename")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'", required=True),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("clauses", "string", "Clauses to excerpt"),
                ToolArgumentSpec("tdoc_id", "string", ""),
                ToolArgumentSpec("title", "string", ""),
                ToolArgumentSpec("source_company", "string", ""),
                ToolArgumentSpec("work_item", "string", ""),
                ToolArgumentSpec("reason_for_change", "string", ""),
                ToolArgumentSpec("summary_of_change", "string", ""),
                ToolArgumentSpec("consequences_if_not_approved", "string", ""),
                ToolArgumentSpec("source_docx_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
            ),
            handler=_wrap_handler(_handle_cr_pcr_generator),
        ),
        ToolDefinition(
            name="word_sanitizer_and_compliance",
            description="Audit/sanitize Word (.docx) hygiene (F-07: language unification, purge language revisions, author consolidation), Word COM compare, or MIP sensitivity labels.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("audit_hygiene", "sanitize_document", "compare_documents", "inspect_label", "apply_label")),
                ToolArgumentSpec("input_path", "string", "", required=True),
                ToolArgumentSpec("revised_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("target_language", "string", "Default 'en-GB'"),
                ToolArgumentSpec("target_author", "string", ""),
                ToolArgumentSpec("label_preset", "string", "'OFFEN'|'INTERN'"),
            ),
            handler=_wrap_handler(_handle_word_sanitizer_and_compliance),
        ),
        ToolDefinition(
            name="plantuml_office_studio",
            description="Validate/render PlantUML to SVG, export to editable Visio (.vsdx) or PowerPoint (.pptx/append), bridge .pptx⇄.vsdx, list 3GPP NF presets, or extract PlantUML from 3GPP-Tools-generated .vsdx.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("validate_or_render_svg", "export_vsdx", "export_pptx", "bridge_pptx_vsdx", "extract_puml_from_vsdx", "list_presets")),
                ToolArgumentSpec("puml_code", "string", ""),
                ToolArgumentSpec("input_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("append_to_open", "boolean", ""),
                ToolArgumentSpec("query", "string", ""),
            ),
            handler=_wrap_handler(_handle_plantuml_office_studio),
        ),
        ToolDefinition(
            name="work_items_and_emails",
            description="Search 3GPP Work Items (with starred/following status), fetch live WI details from 3GPP Portal, toggle WI tracking, or search WG reflector emails.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_work_items", "fetch_wi_from_portal", "toggle_wi_tracking", "search_emails")),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("release", "string", ""),
                ToolArgumentSpec("tracked_only", "boolean", ""),
                ToolArgumentSpec("starred", "boolean", ""),
                ToolArgumentSpec("following", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_work_items_and_emails(a, specs_db_path=paths["specs_db"])),
        ),
    ]

    allowed_names: Optional[Set[str]] = None
    prof_key = (profile or "all").strip().lower()
    if prof_key in PROFILE_TOOLS:
        allowed_names = set(PROFILE_TOOLS[prof_key])

    allowed_perms = (
        {ToolPermission.KNOWLEDGE_READ}
        if read_only
        else {ToolPermission.KNOWLEDGE_READ, ToolPermission.DOCUMENT_WRITE, ToolPermission.NETWORK_FETCH}
    )
    registry = ToolRegistry(allowed_permissions=allowed_perms)
    for tool in all_defs:
        if allowed_names is not None and tool.name not in allowed_names:
            continue
        if read_only and tool.permission != ToolPermission.KNOWLEDGE_READ:
            continue
        if read_only:
            orig_handler = tool.handler
            mutating_actions = {
                "download",
                "index_for_search",
                "convert_to_md",
                "extract_clause",
                "ingest_history",
                "tdoc_dossier",
                "get_or_set_stance",
                "ingest_session_plan",
                "toggle_wi_tracking",
            }
            def _ro_guard(orig=orig_handler, tname=tool.name, **kw: Any) -> ToolResult:
                act = str(kw.get("action") or "").strip().lower()
                if act in mutating_actions:
                    return ToolResult(
                        status=ToolStatus.FOUND,
                        data={"ok": False, "error": f"Read-only mode enabled: action '{act}' on '{tname}' is disabled."},
                        message="Read-only mode enabled",
                    )
                return orig(**kw)
            tool = ToolDefinition(
                name=tool.name,
                description=tool.description,
                arguments=tool.arguments,
                handler=_ro_guard,
                permission=tool.permission,
            )
        registry.register(tool)
    return registry
