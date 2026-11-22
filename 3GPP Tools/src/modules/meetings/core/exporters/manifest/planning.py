# --- File: src/modules/meetings/core/exporters/manifest/planning.py ---
"""Incremental session planning and non-destructive corpus tree merging."""

import datetime
import os
from pathlib import Path
from typing import Optional, Set

from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.tdoc_markdown_converter import sanitize_filename

from .models import ExportManifest, IncrementalPlan
from .validation import is_export_complete, sanitize_folder_name


def plan_incremental_session(
    manifest: Optional[ExportManifest],
    incoming_entries: dict,
    export_dir: Path,
    force_reconvert: bool = False,
) -> IncrementalPlan:
    """
    Analyzes incoming export entries against existing manifest and disk files to plan
    an incremental session.
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if manifest is None or not manifest.tdocs:
        # Fresh / Initial export
        new_tdocs = set()
        for rel in incoming_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        new_tdocs.add(str(doc["tdoc_id"]).strip().upper())
        return IncrementalPlan(
            is_incremental=False,
            session_id=1,
            session_timestamp=now_str,
            new_tdocs=new_tdocs,
            tdocs_to_skip=set(),
            tdocs_to_update=set(),
            preserved_tdocs=set(),
        )

    # Incremental Session
    session_id = len(manifest.sessions) + 1
    new_tdocs: Set[str] = set()
    tdocs_to_skip: Set[str] = set()
    tdocs_to_update: Set[str] = set()

    for rel_name, wis in incoming_entries.items():
        rel_slug = sanitize_folder_name(rel_name)
        for wi_name, wi_data in wis.items():
            wi_slug = sanitize_folder_name(wi_name)
            for cat_name, cat_docs in wi_data.get("categories", {}).items():
                cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                type_dir = export_dir / rel_slug / wi_slug / cat_folder
                for doc in cat_docs:
                    tid = str(doc["tdoc_id"]).strip().upper()
                    target_abs = doc.get("target_md_abs_path")
                    if not target_abs:
                        target_abs = type_dir / f"{sanitize_filename(tid)}.md"

                    # Check if target file already exists on disk
                    file_exists = False
                    existing_path = None
                    if target_abs and os.path.exists(to_extended_path(target_abs)):
                        file_exists = True
                        existing_path = Path(to_extended_path(target_abs))
                    elif tid in manifest.tdocs:
                        saved_rel = manifest.tdocs[tid].get("rel_path")
                        if saved_rel and os.path.exists(to_extended_path(export_dir / saved_rel)):
                            file_exists = True
                            existing_path = Path(to_extended_path(export_dir / saved_rel))

                    # Completeness check: is the existing export non-empty, complete, and valid?
                    complete = is_export_complete(existing_path, doc) if (file_exists and existing_path) else False

                    if tid in tdocs_to_update:
                        continue
                    if file_exists and complete and not force_reconvert:
                        tdocs_to_skip.add(tid)
                    elif file_exists and (not complete or force_reconvert):
                        tdocs_to_skip.discard(tid)
                        tdocs_to_update.add(tid)
                    else:
                        tdocs_to_skip.discard(tid)
                        new_tdocs.add(tid)

    # All TDocs already in manifest that were not in incoming entries are preserved
    incoming_all_tids = {
        str(d["tdoc_id"]).strip().upper()
        for rel in incoming_entries.values()
        for wi in rel.values()
        for cat_docs in wi["categories"].values()
        for d in cat_docs
    }
    preserved_tdocs = set(manifest.tdocs.keys()) - incoming_all_tids

    return IncrementalPlan(
        is_incremental=True,
        session_id=session_id,
        session_timestamp=now_str,
        new_tdocs=new_tdocs,
        tdocs_to_skip=tdocs_to_skip,
        tdocs_to_update=tdocs_to_update,
        preserved_tdocs=preserved_tdocs,
    )


def merge_manifest_into_tree(
    manifest: ExportManifest,
    tree: dict,
    export_dir: Path,
) -> dict:
    """
    Merges preserved documents from manifest into the export tree so master and local
    index.md reports display the complete corpus without accidental loss.
    """
    if not manifest or not manifest.tdocs:
        return tree

    # Set of all TDoc IDs already in tree
    existing_in_tree = {
        str(d["tdoc_id"]).strip().upper()
        for rel in tree.values()
        for wi in rel.values()
        for cat_docs in wi["categories"].values()
        for d in cat_docs
    }

    for tid, doc_info in manifest.tdocs.items():
        if tid in existing_in_tree:
            continue

        rel_name = doc_info.get("release") or "General"
        wi_name = doc_info.get("wi") or "General"
        cat_name = doc_info.get("category") or "DISC"
        rel_path = doc_info.get("rel_path") or f"{rel_name}/{wi_name}/{cat_name}/{sanitize_filename(tid)}.md"
        target_abs = export_dir / rel_path
        if not target_abs.is_file() or not is_export_complete(target_abs):
            continue

        if rel_name not in tree:
            tree[rel_name] = {}
        if wi_name not in tree[rel_name]:
            tree[rel_name][wi_name] = {
                "info": {"name": wi_name},
                "ais": set(),
                "ai_titles": {},
                "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
            }

        ai_num = doc_info.get("ai") or "Unassigned"
        if ai_num and ai_num != "Unassigned":
            tree[rel_name][wi_name]["ais"].add(ai_num)

        entry = {
            "tdoc_id": tid,
            "title": doc_info.get("title", ""),
            "source": doc_info.get("source", ""),
            "type": doc_info.get("type", ""),
            "status": doc_info.get("status", ""),
            "ai": ai_num,
            "spec": doc_info.get("spec", ""),
            "category": cat_name,
            "cross_wis": [],
            "replies": doc_info.get("replies", []),
            "reply_to_ids": doc_info.get("reply_to_ids", []),
            "sibling_replies": doc_info.get("sibling_replies", []),
            "supersedes": doc_info.get("supersedes", []),
            "is_reserved": (
                not is_export_complete(target_abs, doc_info)
                if os.path.exists(to_extended_path(target_abs))
                else bool(doc_info.get("is_reserved", False))
            ),
            "my_status": doc_info.get("my_status", ""),
            "my_notes": doc_info.get("my_notes", ""),
            "for_action": "",
            "sanitized_companies": [],
            "row_data": doc_info.get("row_data", {}),
            "tdoc_url": "",
            "target_md_abs_path": target_abs,
            "target_md_rel_path": rel_path,
            "wi_local_rel_path": f"{cat_name}/{sanitize_filename(tid)}.md",
            "is_preserved_from_previous": True,
        }

        if cat_name not in tree[rel_name][wi_name]["categories"]:
            tree[rel_name][wi_name]["categories"][cat_name] = []
        tree[rel_name][wi_name]["categories"][cat_name].append(entry)

    # Sort entries within categories
    for rel in tree.values():
        for wi in rel.values():
            for cat in wi["categories"]:
                wi["categories"][cat].sort(key=lambda x: str(x.get("tdoc_id", "")))

    return tree
