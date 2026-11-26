"""External agent handlers for Word/Markdown Studio, CR/pCR Generation, and Compliance."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

from core.utils.paths import get_output_root
from .common import (
    _COM_LOCK,
    _disable_faulthandler_for_com,
    _export_doc_to_markdown,
    _extract_resolved_word_docs,
    _match_heading_for_clause,
    _parse_csv_list,
)


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
        src = input_path
        if src.suffix.lower() == ".doc":
            from modules.word_tools.core.word_converter import convert_doc_to_docx
            conv = convert_doc_to_docx(str(src))
            if conv and Path(conv).exists():
                src = Path(conv)
        out_md = Path(output_path_raw) if output_path_raw else src.with_suffix(".md")
        rev_mode = str(args.get("revisions") or "markup").strip().lower()
        produced, text, res = _export_doc_to_markdown(src, out_md, revisions=rev_mode)
        if not produced or not produced.exists():
            return {"error": f"Failed converting {src} to Markdown."}
        return {
            "input_path": str(input_path),
            "output_path": str(produced),
            "revisions_mode": rev_mode,
            "media_dir": None,
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
            h = _match_heading_for_clause(headings, c)
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
        source_paths = [str(p) for p in _extract_resolved_word_docs(resolved)]

    if not source_paths:
        return {"error": f"No baseline Word document found for {spec_number} v{spec_version}. Provide source_docx_path or download the spec first."}

    from modules.word_tools.core.word_excerpt_extractor import WordHeadingScanner

    scanner = WordHeadingScanner()
    headings = scanner.scan(source_paths)
    matched_ids: List[str] = []
    for c in clauses:
        h = _match_heading_for_clause(headings, c)
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
        labels = WordConfig.get_sensitivity_labels()
        label_id = None
        label_name = None
        if preset_name:
            needle = preset_name.lower()
            for p in labels:
                if (
                    (p.get("name") or "").strip().lower() == needle
                    or (p.get("display_name") or "").strip().lower() == needle
                    or (p.get("id") or "").strip().lower() == needle
                ):
                    label_id = p.get("id")
                    label_name = p.get("name")
                    break
            if not label_id:
                available = [p.get("name") for p in labels if p.get("name")]
                return {
                    "error": f"Sensitivity label '{preset_name}' not found in configured/Windows policy labels.",
                    "available_labels": available,
                }
        else:
            active = WordConfig.get_active_sensitivity_label()
            label_id = active.get("id")
            label_name = active.get("name")
            if not label_id or not label_name:
                return {
                    "error": "No sensitivity label configured or detected in local Windows policy.",
                }

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
            "applied_label": {"id": label_id, "name": label_name},
            "summary": summary,
        }

    return {"error": f"Unsupported action '{action}' for word_sanitizer_and_compliance."}
