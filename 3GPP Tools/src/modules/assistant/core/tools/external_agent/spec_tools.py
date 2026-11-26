"""External agent handlers for Specification Catalogue, Files, Text Search, and Clauses."""

from __future__ import annotations

import difflib
import re
import urllib.parse
from pathlib import Path
from typing import Any, Dict, List, Tuple

from core.network.session import NetworkSession
from core.utils.paths import get_output_root, get_specs_root
from modules.assistant.core.knowledge import (
    SpecificationKnowledgeService,
    SpecSearchKnowledgeService,
)
from .common import (
    _COM_LOCK,
    _disable_faulthandler_for_com,
    _export_doc_to_markdown,
    _extract_resolved_word_docs,
    _match_heading_for_clause,
    _parse_csv_list,
)


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

    wg = str(args.get("wg") or "").strip().upper()

    if action == "search":
        if wg:
            from modules.specifications.core.database import SpecsDatabase
            from modules.assistant.core.knowledge.specifications import _normalize_file_row

            specs_db = SpecsDatabase(spec_service.db_path)
            raw_rows = specs_db.search_files(spec_number=(query or spec_number or None), group=wg)
            grouped: Dict[str, Dict[str, Any]] = {}
            for r in raw_rows or []:
                norm = _normalize_file_row(r)
                if not norm:
                    continue
                sp = norm["specification"]
                entry = grouped.setdefault(
                    sp,
                    {
                        "specification": sp,
                        "type": norm.get("type", "TS"),
                        "title": norm.get("title", ""),
                        "primary_group": norm.get("primary_group", wg),
                        "versions": [],
                    },
                )
                ver = norm.get("version")
                if ver and ver not in entry["versions"]:
                    entry["versions"].append(ver)
            items = list(grouped.values())[:limit]
            return {"count": len(items), "items": items}
        return spec_service.find_specifications(
            query=query or spec_number,
            limit=limit,
        )

    if action == "get":
        if not spec_number:
            return {"error": "spec_number is required for action='get'."}
        res = spec_service.get_specification(specification=spec_number)
        if isinstance(res.data, dict):
            data_out = dict(res.data)
            if "known_versions" in data_out:
                data_out.setdefault("versions", data_out["known_versions"])
                data_out.setdefault("count", len(data_out["known_versions"]))
            indexed_versions: List[str] = []
            try:
                from modules.spec_search.core.spec_search_db import SpecSearchDatabase
                clean_spec = str(data_out.get("specification") or spec_number).strip()
                if spec_search_db_path.exists():
                    sdb = SpecSearchDatabase(spec_search_db_path)
                    indexed_versions = [
                        str(v.get("version") or "")
                        for v in sdb.get_versions_for_spec(clean_spec)
                        if v.get("version")
                    ]
            except Exception:
                pass
            data_out["indexed_versions"] = indexed_versions
            data_out["is_indexed"] = bool(indexed_versions)
            from modules.assistant.core.agent_models import ToolResult
            return ToolResult(res.status, data=data_out, message=res.message, metadata=res.metadata)
        return res

    if action == "resolve_files":
        if not spec_number:
            return {"error": "spec_number is required for action='resolve_files'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.specifications.core.database import SpecsDatabase
        from modules.specifications.core.spec_history_extractor import SpecHistoryExtractor

        specs_db = SpecsDatabase(spec_service.db_path) if spec_service.db_path.exists() else None
        info = SpecDocResolver.resolve_spec_files(spec_number, version=version or "", specs_db=specs_db)
        resolved_parts = [str(p) for p in _extract_resolved_word_docs(info)]
        terminal_doc = SpecHistoryExtractor.resolve_spec_document(spec_number, version=version)
        spec_dir = Path(info["spec_dir"]) if isinstance(info, dict) and info.get("spec_dir") else (get_specs_root() / spec_number)
        alt_dir = get_specs_root() / spec_number.replace(".", "")
        zips: List[str] = []
        if isinstance(info, dict) and info.get("zip_file"):
            zips.append(str(info["zip_file"]))
        for d in dict.fromkeys((spec_dir, alt_dir)):
            if d.exists():
                for p in sorted(d.glob("*.zip")):
                    if str(p) not in zips:
                        zips.append(str(p))
        return {
            "spec_number": spec_number,
            "version": info.get("version") if isinstance(info, dict) else version,
            "spec_dir": str(spec_dir),
            "resolved_doc_parts": resolved_parts,
            "terminal_or_history_doc": str(terminal_doc) if terminal_doc else None,
            "cached_zip_archives": zips,
            "local_ready": bool(resolved_parts),
        }

    if action == "download":
        if not spec_number:
            return {"error": "spec_number is required for action='download'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
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
        target_dir = get_specs_root() / spec_number
        target_dir.mkdir(parents=True, exist_ok=True)
        zip_name = urllib.parse.unquote(zip_url.split("/")[-1])
        if not zip_name.lower().endswith(".zip"):
            zip_name = f"{clean_num}-{chosen.get('version', 'latest')}.zip"
        local_zip = target_dir / zip_name

        if not local_zip.exists() or local_zip.stat().st_size == 0:
            ok = NetworkSession.download_file(zip_url, local_zip, apply_delay=True)
            if not ok or not local_zip.exists():
                return {"error": f"Failed to download {zip_url}"}

        extracted_docs = SpecDocResolver.extract_word_from_zip(local_zip, target_dir)
        info = SpecDocResolver.resolve_spec_files(spec_number, version=str(chosen.get("version") or ""), specs_db=specs_db)
        parts = [str(p) for p in _extract_resolved_word_docs(info)]
        return {
            "spec_number": spec_number,
            "version": chosen.get("version"),
            "release": chosen.get("release"),
            "zip_url": zip_url,
            "local_zip": str(local_zip),
            "extracted_docs": [str(p) for p in (extracted_docs or [])],
            "resolved_doc_parts": parts,
        }

    if action == "index_for_search":
        if not spec_number:
            return {"error": "spec_number is required for action='index_for_search'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver, letter_code_to_version
        from modules.spec_search.core.spec_search_db import SpecSearchDatabase
        from modules.spec_search.core.spec_text_extractor import SpecDocxExtractor
        from modules.specifications.core.database import SpecsDatabase

        specs_db = SpecsDatabase(spec_service.db_path) if spec_service.db_path.exists() else None
        info = SpecDocResolver.resolve_spec_files(spec_number, version=version or "", specs_db=specs_db)
        parts = _extract_resolved_word_docs(info)
        if not parts and isinstance(info, dict) and info.get("zip_file"):
            parts = SpecDocResolver.extract_word_from_zip(Path(info["zip_file"]), Path(info["spec_dir"]))
        if not parts:
            return {"error": f"No local .docx/.doc files found for {spec_number}. Run action='download' first."}

        search_db = SpecSearchDatabase(spec_search_db_path)
        resolved_ver = str((info.get("version") if isinstance(info, dict) else "") or "")
        if not resolved_ver:
            m_code = re.search(r"[-_\s]([a-z0-9]{3})(?:[-_.]|$)", parts[0].name.lower())
            if m_code:
                resolved_ver = letter_code_to_version(m_code.group(1)) or ""
        actual_ver = (version or resolved_ver or "unknown").lstrip("vV")
        actual_rel = release or (f"Rel-{actual_ver.split('.')[0]}" if "." in actual_ver else "Unknown")

        docx_parts: List[Path] = []
        for p in parts:
            if p.suffix.lower() == ".doc":
                from modules.word_tools.core.word_converter import convert_doc_to_docx
                conv = convert_doc_to_docx(str(p))
                if conv and Path(conv).exists():
                    docx_parts.append(Path(conv))
            else:
                docx_parts.append(p)

        all_clauses = SpecDocxExtractor(docx_parts).extract_clauses()

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
            "parts_indexed": [str(p) for p in docx_parts],
            "clause_count": len(all_clauses),
        }

    if action == "convert_to_md":
        if not spec_number:
            return {"error": "spec_number is required for action='convert_to_md'."}
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.specifications.core.spec_markdown_converter import convert_spec_to_markdown
        from modules.specifications.core.database import SpecsDatabase

        specs_db = SpecsDatabase(spec_service.db_path)
        info = SpecDocResolver.resolve_spec_files(spec_number, version=version or "", specs_db=specs_db)
        parts = _extract_resolved_word_docs(info)
        if not parts and isinstance(info, dict) and info.get("zip_file"):
            parts = SpecDocResolver.extract_word_from_zip(Path(info["zip_file"]), Path(info["spec_dir"]))
        if not parts:
            return {"error": f"No local Word files found for {spec_number}. Run action='download' first."}

        details = specs_db.get_spec_details(spec_number) or {}
        mode = str(args.get("mode") or "single").strip().lower()
        effective_mode = mode if mode in ("single", "split") else "single"
        spec_dir = Path(info["spec_dir"])
        stem = str(info.get("stem") or spec_number.replace(".", ""))
        resolved_ver = str(info.get("version") or version or "")
        out_md = convert_spec_to_markdown(
            spec_dir=spec_dir,
            stem=stem,
            mode=effective_mode,
            meta={
                "spec": spec_number,
                "version": resolved_ver,
                "title": str(details.get("title") or ""),
            },
        )
        custom_out_dir = str(args.get("output_dir") or "").strip()
        if custom_out_dir:
            import shutil
            dest_dir = Path(custom_out_dir)
            dest_dir.mkdir(parents=True, exist_ok=True)
            if effective_mode == "split" and out_md.parent.is_dir():
                target_container = dest_dir / out_md.parent.name
                if target_container.resolve() != out_md.parent.resolve():
                    shutil.copytree(out_md.parent, target_container, dirs_exist_ok=True)
                    out_md = target_container / out_md.name
            elif out_md.is_file():
                target_file = dest_dir / out_md.name
                if target_file.resolve() != out_md.resolve():
                    shutil.copy2(out_md, target_file)
                    out_md = target_file

        return {
            "spec_number": spec_number,
            "version": resolved_ver,
            "mode": effective_mode,
            "output_dir": str(out_md.parent),
            "master_file": str(out_md),
            "exists": out_md.exists(),
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
        latest_only = bool(args.get("latest_only", False))
        res = search_service.search_text(
            query=query,
            specification=spec_number or None,
            version=version,
            release=rel_int,
            clause=clause_id or None,
            limit=limit * 3 if latest_only else limit,
        )
        if latest_only and isinstance(res.data, list) and res.data:
            def _v_tup(v_str: Any) -> Tuple[int, ...]:
                nums = re.findall(r"\d+", str(v_str or ""))
                return tuple(int(n) for n in nums) if nums else (0,)

            best_ver: Dict[str, Tuple[int, ...]] = {}
            for item in res.data:
                sp = str(item.get("specification") or "")
                vt = _v_tup(item.get("version"))
                if sp not in best_ver or vt > best_ver[sp]:
                    best_ver[sp] = vt
            filtered = [
                item for item in res.data
                if _v_tup(item.get("version")) == best_ver.get(str(item.get("specification") or ""))
            ][:limit]
            from modules.assistant.core.agent_models import ToolResult
            return ToolResult(res.status, data=filtered, message=res.message, metadata=res.metadata)
        return res

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
        from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
        from modules.word_tools.core.word_excerpt_extractor import WordExcerptEngine, WordHeadingScanner

        info = SpecDocResolver.resolve_spec_files(spec_number, version=version or "")
        parts = _extract_resolved_word_docs(info)
        if not parts and isinstance(info, dict) and info.get("zip_file"):
            parts = SpecDocResolver.extract_word_from_zip(Path(info["zip_file"]), Path(info["spec_dir"]))
        if not parts:
            return {"error": f"No local Word files found for {spec_number}. Download the spec first."}

        scanner = WordHeadingScanner()
        headings = scanner.scan([str(p) for p in parts])
        target_clauses = _parse_csv_list(clause_id)
        matched_ids: List[str] = []
        matched_titles: List[str] = []
        for cid in target_clauses:
            h = _match_heading_for_clause(headings, cid)
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
        md_path, md_text, _ = _export_doc_to_markdown(out_docx, out_md, revisions="accept")
        return {
            "spec_number": spec_number,
            "version": version,
            "clauses_extracted": matched_titles,
            "docx_path": str(out_docx),
            "markdown_path": str(md_path) if md_path else None,
            "markdown_preview": md_text[:4000],
        }

    return {"error": f"Unsupported action '{action}' for spec_text_and_clauses."}
