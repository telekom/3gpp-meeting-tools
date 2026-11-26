"""External agent handlers for Specification Change History/Attribution and Protocol Inspection."""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Any, Dict, List

from core.utils.paths import get_output_root
from modules.assistant.core.agent_models import ToolResult
from modules.assistant.core.knowledge import ProtocolKnowledgeService
from .common import resolve_meetings_cache_roots as _get_meetings_cache_roots


def _strip_html(text: str) -> str:
    """Strips HTML tags and collapses whitespace for compact, LLM-friendly text."""
    cleaned = re.sub(r"<[^>]+>", " ", html.unescape(str(text or "")))
    return " ".join(cleaned.split())


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
        tdoc_id = str(args.get("tdoc_id") or "").strip() or None
        rows = history_db.search_history(
            spec_number=spec_number or None,
            query=query or None,
            meeting=args.get("meeting"),
            category=args.get("category"),
            version=version,
            clause_id=clause_id or None,
            tdoc_id=tdoc_id,
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
        from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener

        search_db = SpecSearchDatabase(spec_search_db_path)
        engine = SpecClauseAttributionEngine(
            history_db=history_db,
            search_db=search_db,
        )
        clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number, flags=re.IGNORECASE).strip()
        target_ver = version
        if not target_ver:
            hist_vers = history_db.get_versions_for_spec(clean_spec)
            target_ver = hist_vers[0] if hist_vers else "20.0.0"

        clause_title = str(args.get("clause_title") or "")
        clause_text = str(args.get("clause_text") or "")
        if not clause_title or not clause_text:
            try:
                all_imp = search_db.get_imported_versions()
                spec_v_rows = [v for v in all_imp if str(v.get("spec_number")) in (spec_number, clean_spec)]
                if spec_v_rows:
                    vid = int(spec_v_rows[0]["id"])
                    cl_row = search_db.get_clause_content(vid, clause_id)
                    if cl_row:
                        clause_title = clause_title or str(cl_row.get("clause_title") or "")
                        clause_text = clause_text or str(cl_row.get("content") or "")
            except Exception:
                pass

        candidates = engine.attribute_clause(
            spec_number=clean_spec,
            clause_number=clause_id,
            target_version=target_ver,
            base_version=str(args.get("base_version") or "") or None,
            clause_title=clause_title,
            diff_text=str(args.get("diff_text") or ""),
            clause_content=clause_text,
        )
        if bool(args.get("deep_verify_docx", False)) and candidates:
            for cand in candidates[: min(5, len(candidates))]:
                tdoc_tok = str(cand.get("tdoc") or "")
                local_doc = SpecTDocOpener.resolve_local_tdoc_file(tdoc_tok) if tdoc_tok else None
                if local_doc and local_doc.exists():
                    matched, pct, snip = engine.verify_candidate_in_docx(
                        tdoc=tdoc_tok,
                        diff_text=str(args.get("diff_text") or clause_text),
                        doc_path=local_doc,
                        cr_number=str(cand.get("cr_number") or ""),
                    )
                    if matched:
                        cand["is_verified"] = True
                        cand["verified_match_pct"] = pct
                        cand["verified_snippet"] = snip
                        cand["score"] = max(int(cand.get("score") or 0), min(100, int(cand.get("score") or 0) + 20))
            candidates.sort(key=lambda x: (bool(x.get("is_verified")), int(x.get("score") or 0)), reverse=True)

        return {
            "spec_number": clean_spec,
            "clause_id": clause_id,
            "target_version": target_ver,
            "clause_title": clause_title,
            "all_candidates_count": len(candidates),
            "candidates": candidates[:limit],
        }

    if action == "rank_origin_crs":
        if not query:
            return {"error": "query is required for action='rank_origin_crs'."}
        from modules.spec_search.core.spec_search_db import SpecSearchDatabase
        from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine

        search_db = SpecSearchDatabase(spec_search_db_path)
        engine = SpecClauseAttributionEngine(
            history_db=history_db,
            search_db=search_db,
        )
        clean_spec = re.sub(r"^(?:3GPP\s*)?(?:TS|TR)\s*", "", spec_number or "23.501", flags=re.IGNORECASE).strip()
        all_imported = search_db.get_imported_versions()
        v_ids = [int(v["id"]) for v in all_imported if str(v.get("spec_number")) in (spec_number, clean_spec)]
        clause_hints: List[str] = []
        if v_ids:
            df_hits = search_db.search_substring(query, v_ids)
            hits_resp = df_hits.to_dict(orient="records") if hasattr(df_hits, "to_dict") else []
            clause_hints = [str(h.get("clause_number") or "") for h in hits_resp[:30] if h.get("clause_number")]
        else:
            hits_resp = []

        ranked = engine.rank_origin_crs_for_query(
            spec_number=clean_spec,
            query=query,
            target_version=version,
            affected_clauses_hint=clause_hints or None,
        )
        return {
            "query": query,
            "spec_number": clean_spec,
            "clause_hits_analyzed": len(hits_resp),
            "ranked_origin_crs": ranked[:limit],
        }

    if action == "tdoc_dossier":
        tdoc_arg = str(args.get("tdoc_id") or "").strip()
        if not spec_number and not tdoc_arg:
            return {"error": "spec_number or tdoc_id is required for action='tdoc_dossier'."}
        from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

        tdoc_tokens: List[str] = []
        if tdoc_arg:
            tdoc_tokens.append(tdoc_arg)
        if spec_number:
            rows = history_db.search_history(
                spec_number=spec_number,
                version=version,
                tdoc_id=tdoc_arg or None,
                limit=500,
            )
            for r in rows:
                st = str(r.get("source_tdoc") or "").strip()
                if st and st != "-":
                    tdoc_tokens.append(st)
                for pcr in r.get("pcr_tdocs") or []:
                    if pcr and pcr != "-":
                        tdoc_tokens.append(str(pcr).strip())

        cache_roots = _get_meetings_cache_roots()
        gen = SpecTDocReportGenerator(
            data_db_path=specs_db_path,
            cache_root=cache_roots[0] if cache_roots else None,
            history_db=history_db,
        )
        resolved_items = gen.resolve_tdocs_metadata(tdoc_tokens)
        label_spec = spec_number or tdoc_arg
        md_report = SpecTDocReportGenerator.generate_markdown_report(
            spec_number=label_spec,
            version=version or "all",
            tdoc_entries=resolved_items,
        )
        out_dir = Path(args.get("output_dir") or (get_output_root() / "Spec_TDoc_Dossiers"))
        out_dir.mkdir(parents=True, exist_ok=True)
        safe_label = re.sub(r"[^A-Za-z0-9._-]+", "_", label_spec)
        out_md = out_dir / f"{safe_label}_{version or 'all'}_tdoc_dossier.md"
        out_md.write_text(md_report, encoding="utf-8")
        return {
            "spec_number": spec_number or None,
            "tdoc_id": tdoc_arg or None,
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
        if isinstance(res.data, dict):
            data_copy = dict(res.data)
            if isinstance(data_copy.get("description"), str) and "<" in data_copy["description"]:
                data_copy["description"] = _strip_html(data_copy["description"])
            if "fields" in data_copy:
                raw_fields = data_copy.pop("fields") or []
                cleaned_fields: List[Dict[str, Any]] = []
                for f in raw_fields:
                    fc = dict(f)
                    for dk in ("ie_description", "field_description", "description"):
                        if isinstance(fc.get(dk), str) and "<" in fc[dk]:
                            fc[dk] = _strip_html(fc[dk])
                    cleaned_fields.append(fc)
                data_copy["information_elements"] = cleaned_fields
            return ToolResult(res.status, data=data_copy, message=res.message, metadata=res.metadata)
        return res

    if action == "find_ie":
        if not query:
            return {"error": "query (IE name) is required for action='find_ie'."}
        return protocol_service.find_ie(ie=query, protocol=protocol or None, version=version)

    if action == "evolution_matrix":
        if not protocol or not query:
            return {"error": "protocol and query (message or IE name) are required for action='evolution_matrix'."}
        from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
        from modules.nas.core.nas_exporter import build_matrix_markdown, build_matrix_tsv

        nas_db = NASDatabase(protocol_db_path)
        imported = nas_db.get_imported_versions()
        descriptors = protocol_service._resolve_descriptors(protocol)
        if not descriptors:
            return {"error": f"Unknown protocol '{protocol}'."}
        selected_vers = protocol_service._select_versions(imported, descriptors, version, None)
        if not selected_vers:
            return {"error": f"No imported versions found in protocol DB for {protocol}."}
        selected_vers_sorted = sorted(
            selected_vers,
            key=lambda v: parse_version_tuple(str(v.get("version") or "")),
        )
        version_ids = [int(v["id"]) for v in selected_vers_sorted if v.get("id") is not None]

        target_type = str(args.get("target_type") or "message").strip().lower()
        fmt = str(args.get("format") or "markdown").strip().lower()

        if target_type == "ie":
            df = nas_db.get_ie_definition_evolution_df(clause=query, ie_name=query, version_ids=version_ids)
            title = f"{protocol} IE Evolution: {query}"
        else:
            df = nas_db.get_message_evolution_df(message_name=query, version_ids=version_ids)
            title = f"{protocol} Message Evolution: {query}"

        if df is None or df.empty:
            return {"error": f"No evolution matrix rows found for {protocol} {target_type} '{query}'."}

        # Pivot DataFrame into compact per-version columns for Markdown/TSV matrix rendering
        ver_cols = sorted(
            {str(v) for v in df["version"].dropna() if str(v)},
            key=parse_version_tuple,
        )
        headers = ["IE / Field", "Type Ref", *ver_cols]
        field_order: List[str] = []
        field_type_ref: Dict[str, str] = {}
        cell_map: Dict[str, Dict[str, str]] = {}

        for _, row in df.fillna("").iterrows():
            fname = str(row.get("field_path") or row.get("ie_name") or "").strip()
            if not fname:
                continue
            if fname not in cell_map:
                field_order.append(fname)
                cell_map[fname] = {}
                field_type_ref[fname] = str(row.get("type_reference") or "").strip()
            ver = str(row.get("version") or "").strip()
            pres = str(row.get("presence") or "").strip()
            fmt_s = str(row.get("format") or "").strip()
            len_s = str(row.get("length") or "").strip()
            detail = " ".join(p for p in (pres, f"[{fmt_s}]" if fmt_s else "", len_s) if p)
            if ver:
                cell_map[fname][ver] = detail or "Present"

        matrix_rows: List[List[Any]] = [
            [fname, field_type_ref[fname], *[cell_map[fname].get(v, "-") for v in ver_cols]]
            for fname in field_order
        ]
        rendered = build_matrix_tsv(headers, matrix_rows) if fmt == "tsv" else build_matrix_markdown(title, headers, matrix_rows)
        return {
            "protocol": protocol,
            "target_type": target_type,
            "target_name": query,
            "format": fmt,
            "row_count": len(matrix_rows),
            "columns": headers,
            "matrix": rendered,
        }

    return {"error": f"Unsupported action '{action}' for protocol_inspector."}
