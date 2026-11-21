# --- File: src/modules/meetings/core/exporters/pre_meeting/mixins/export_dossiers.py ---
"""Parallel pre-download, canonical Markdown conversion, and callout assembly mixin."""

from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import os
import posixpath
import re
import shutil
from pathlib import Path
from typing import Dict, List, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.post_meeting_exporter import (
    safe_markdown_rel_link,
    sanitize_folder_name,
)
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    sanitize_filename,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url

logger = logging.getLogger("modules.meetings.core.exporters.pre_meeting_exporter")


class PreMeetingExportMixin:
    """Exports Pre-Meeting per-TDoc dossiers, callouts, stubs, and media assets."""

    def _export_all_dossiers(self, filtered_entries: dict):
        """Converts and copies documents into their respective type subfolders with Pre-Meeting callouts."""
        unique_tdocs = {}
        for rel in filtered_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        tid = doc["tdoc_id"]
                        if tid not in unique_tdocs:
                            unique_tdocs[tid] = doc

        total_unique = len(unique_tdocs)
        converted_cache = {}  # tdoc_id -> (local_md_file, local_media_dir)

        # Phase 0: Parallel pre-download of missing non-reserved documents
        if self.download_missing:
            missing_dl_tasks: List[Tuple[str, Path]] = []
            for tdoc_id, doc_entry in unique_tdocs.items():
                # Skip unuploaded reserved documents from network download unless upgrading from stub
                is_stub_upgrade = hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_update
                if doc_entry.get("is_reserved") and not is_stub_upgrade:
                    continue

                # Skip documents preserved from previous export or already existing on disk
                if doc_entry.get("is_preserved_from_previous"):
                    continue
                if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                    continue

                row_data = doc_entry["row_data"]
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                tdoc_folder = self.meeting_dir / base_tdoc
                local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                legacy_cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{CONVERTER_VERSION}.md"
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                doc_type = row_data.get("Type", "") if row_data else ""
                expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                needs_conversion = self.force_reconvert or is_conversion_stale(
                    doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                )
                if needs_conversion and not self.force_reconvert and os.path.exists(to_extended_path(legacy_cache_file)):
                    if not is_conversion_stale(doc_path, legacy_cache_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                        needs_conversion = False

                if needs_conversion and not doc_path:
                    missing_dl_tasks.append((tdoc_id, tdoc_folder))

            if missing_dl_tasks:
                total_dl = len(missing_dl_tasks)
                dl_workers = min(6, total_dl)
                self.progress.emit(f"Downloading {total_dl} missing pre-meeting TDoc(s) in parallel ({dl_workers} workers)...")
                logger.info(f"[PreMeetingExporter] Concurrently downloading {total_dl} missing TDocs using {dl_workers} workers...")
                dl_done = 0
                with ThreadPoolExecutor(max_workers=dl_workers) as executor:
                    future_to_tid = {
                        executor.submit(self._download_and_extract_tdoc, tid, folder): tid
                        for tid, folder in missing_dl_tasks
                    }
                    for future in as_completed(future_to_tid):
                        if self.isInterruptionRequested():
                            try:
                                executor.shutdown(wait=False, cancel_futures=True)
                            except TypeError:
                                executor.shutdown(wait=False)
                            return
                        tid = future_to_tid[future]
                        dl_done += 1
                        try:
                            res_path = future.result()
                            if res_path:
                                self.progress.emit(f"Downloaded TDoc {tid} ({dl_done}/{total_dl})")
                            else:
                                self.progress.emit(f"Failed to download {tid} ({dl_done}/{total_dl})")
                        except Exception as dl_err:
                            logger.error(f"[PreMeetingExporter] Error downloading {tid}: {dl_err}")

        # Phase 1: Ensure all canonical Markdown files exist in meeting cache
        for idx, (tdoc_id, doc_entry) in enumerate(unique_tdocs.items(), 1):
            if self.isInterruptionRequested():
                return

            if doc_entry.get("is_preserved_from_previous"):
                continue
            if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                continue

            row_data = doc_entry["row_data"]
            # If reserved and not present locally, we skip conversion and will write a clean reserved stub
            if doc_entry.get("is_reserved") and not self._find_word_doc(self.meeting_dir / tdoc_id, tdoc_id):
                converted_cache[tdoc_id] = (None, None)
                continue

            self.progress.emit(f"Converting {tdoc_id} to Markdown ({idx}/{total_unique})...")
            local_md_file, local_media_dir = self._ensure_tdoc_markdown(tdoc_id, row_data)
            converted_cache[tdoc_id] = (local_md_file, local_media_dir)
            if local_md_file:
                doc_entry["is_reserved"] = False

        # Pre-register target paths for all documents across the dossier structure
        self.exported_tdoc_paths: Dict[str, Path] = {}
        self.tdoc_titles: Dict[str, str] = {}
        self.tdoc_sources: Dict[str, str] = {}

        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)
                for cat_name, cat_docs in wi_data["categories"].items():
                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        target_md_file = type_dir / f"{sanitize_filename(tdoc_id)}.md"
                        doc["target_md_abs_path"] = target_md_file
                        doc["target_md_rel_path"] = f"{rel_slug}/{wi_slug}/{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        doc["wi_local_rel_path"] = f"{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        if tdoc_id not in self.exported_tdoc_paths:
                            self.exported_tdoc_paths[tdoc_id] = target_md_file
                            self.tdoc_titles[tdoc_id] = doc.get("title", "")
                            self.tdoc_sources[tdoc_id] = doc.get("source", "")

        # Phase 2: Copy dossiers into target directory structure with rich pre-meeting callouts
        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)

                for cat_name, cat_docs in wi_data["categories"].items():
                    if not cat_docs:
                        continue

                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    type_media_dir = type_dir / "media"
                    os.makedirs(to_extended_path(type_dir), exist_ok=True)

                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        if doc.get("is_preserved_from_previous"):
                            continue
                        if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                            continue

                        local_md_file, local_media_dir = converted_cache.get(tdoc_id, (None, None))
                        target_md_file = doc.get("target_md_abs_path") or (type_dir / f"{sanitize_filename(tdoc_id)}.md")

                        header_notices = []

                        # Notice 0: Incremental Addition Header
                        if hasattr(self, "session_plan") and self.session_plan.is_incremental and tdoc_id in self.session_plan.new_tdocs:
                            update_lbl = f"Incremental Update #{self.session_plan.session_id - 1}"
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Incremental Addition:** Added to this briefing on {self.session_plan.session_timestamp} ({update_lbl})."
                            )
                        elif hasattr(self, "session_plan") and self.session_plan.is_incremental and tdoc_id in self.session_plan.tdocs_to_update:
                            update_lbl = f"Incremental Update #{self.session_plan.session_id - 1}"
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Incremental Revision:** Updated and refreshed in this briefing on {self.session_plan.session_timestamp} ({update_lbl})."
                            )

                        # Notice 1: Revision Lineage Notice (Supersedes older revisions)
                        if doc.get("supersedes"):
                            preds_str = ", ".join(f"`{p}`" for p in doc["supersedes"])
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Revision History:** This active pre-meeting contribution supersedes earlier draft(s): {preds_str}."
                            )

                        # Notice 2: Cross-WI Contribution
                        if doc.get("cross_wis"):
                            other_wis_str = ", ".join(f"`{w}`" for w in doc["cross_wis"])
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Cross-WI Contribution:** This TDoc is also associated with {other_wis_str}."
                            )

                        # Notice 3: Personal Pre-Meeting Stance & Notes
                        if self.include_user_notes:
                            my_st = doc.get("my_status", "")
                            my_nt = doc.get("my_notes", "")
                            if (my_st and my_st != "⚪ Neutral") or my_nt:
                                stance_header = f"Delegate Pre-Meeting Stance ({my_st}):" if my_st else "Delegate Pre-Meeting Notes:"
                                notes_body = f"\n> {my_nt}" if my_nt else ""
                                header_notices.append(
                                    f"> [!IMPORTANT]\n"
                                    f"> **{stance_header}**{notes_body}"
                                )

                        # Notice 4: Candidate LS Replies (For LS_IN)
                        if doc.get("category") == "LS_IN" and doc.get("replies"):
                            rep_bullets = []
                            for rep_id in doc["replies"]:
                                rep_abs = self.exported_tdoc_paths.get(rep_id)
                                rep_title = self.tdoc_titles.get(rep_id, "")
                                rep_src = self.tdoc_sources.get(rep_id, "")
                                src_str = f" ({rep_src})" if rep_src else ""
                                title_str = f" — *{rep_title}*" if rep_title else ""
                                if rep_abs:
                                    rel = safe_markdown_rel_link(rep_abs, target_md_file.parent)
                                    rep_bullets.append(f"> - **[{rep_id}]({rel})**{title_str}{src_str}")
                                else:
                                    rep_bullets.append(f"> - **{rep_id}**{title_str}{src_str}")
                            n_cand = len(rep_bullets)
                            cand_lbl = f"{n_cand} candidate proposal(s)" if n_cand > 1 else "1 candidate proposal"
                            header_notices.append(
                                f"> [!TIP]\n"
                                f"> **Proposed Reply Liaison Statements ({cand_lbl}):**\n" + "\n".join(rep_bullets)
                            )

                        # Notice 5: Draft LS Reply provenance & Competing Sibling Replies (For LS_OUT)
                        if doc.get("category") == "LS_OUT" and doc.get("reply_to_ids"):
                            in_links = []
                            for in_id in doc["reply_to_ids"]:
                                in_abs = self.exported_tdoc_paths.get(in_id)
                                in_title = self.tdoc_titles.get(in_id, "")
                                title_str = f" (*{in_title}*)" if in_title else ""
                                if in_abs:
                                    rel = safe_markdown_rel_link(in_abs, target_md_file.parent)
                                    in_links.append(f"**[{in_id}]({rel})**{title_str}")
                                else:
                                    in_links.append(f"**{in_id}**{title_str}")

                            reply_notice = f"> [!NOTE]\n> **Draft Reply Liaison Statement:**\n> - **Replies to Incoming LS:** {', '.join(in_links)}"

                            if doc.get("sibling_replies"):
                                sib_links = []
                                for sib_id in doc["sibling_replies"]:
                                    sib_abs = self.exported_tdoc_paths.get(sib_id)
                                    sib_src = self.tdoc_sources.get(sib_id, "")
                                    src_str = f" ({sib_src})" if sib_src else ""
                                    if sib_abs:
                                        rel = safe_markdown_rel_link(sib_abs, target_md_file.parent)
                                        sib_links.append(f"**[{sib_id}]({rel})**{src_str}")
                                    else:
                                        sib_links.append(f"**{sib_id}**{src_str}")
                                reply_notice += f"\n> - **Competing Reply Proposals ({len(sib_links)} other(s)):** {', '.join(sib_links)}"

                            header_notices.append(reply_notice)

                        notice_block = "\n\n".join(header_notices) + ("\n\n" if header_notices else "")

                        # Navigation breadcrumbs
                        master_index_file = self.export_dir / "index.md"
                        master_rel = safe_markdown_rel_link(master_index_file, target_md_file.parent)
                        wi_index_file = self.export_dir / rel_slug / wi_slug / "index.md"
                        if self.generate_local_indices:
                            wi_rel = safe_markdown_rel_link(wi_index_file, target_md_file.parent)
                            nav_bar = f"[📁 Master Briefing]({master_rel}) · [🎯 {wi_name} Briefing]({wi_rel})\n\n---\n\n"
                        else:
                            nav_bar = f"[📁 Master Briefing]({master_rel})\n\n---\n\n"

                        full_header_block = nav_bar + notice_block

                        # If reserved and uncached, write clean informative reserved stub
                        if doc.get("is_reserved") and not local_md_file:
                            try:
                                source_lines = "\n".join(CompanySanitizer.format_source_markdown_header_lines(doc.get('source', '')))
                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(
                                        f"{full_header_block}"
                                        f"# {tdoc_id}: {doc['title']}\n\n"
                                        f"> ⏳ **Status: Reserved / Not Uploaded**\n"
                                        f"> *This document number has been allocated or reserved on the meeting agenda but has not yet been uploaded to 3GPP servers.*\n\n"
                                        f"{source_lines}\n"
                                        f"- **Type:** {doc['type']}\n"
                                        f"- **Agenda Item:** {doc['ai']}\n"
                                        f"- **Status:** {doc['status'] or 'Reserved'}\n"
                                    )
                            except Exception as err:
                                logger.error(f"[PreMeetingExporter] Error writing reserved stub for {tdoc_id}: {err}")
                            continue

                        if local_md_file and os.path.exists(to_extended_path(local_md_file)):
                            try:
                                with open(to_extended_path(local_md_file), "r", encoding="utf-8") as f:
                                    md_content = f.read()

                                tdoc_url = doc.get("tdoc_url") or resolve_tdoc_download_url(
                                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=doc.get("row_data")
                                )
                                if tdoc_url and "- **3GPP Server URL:**" not in md_content:
                                    url_bullet = f"- **3GPP Server URL:** [{tdoc_url}]({tdoc_url})\n"
                                    if "- **Original file:**" in md_content:
                                        md_content = md_content.replace("- **Original file:**", f"{url_bullet}- **Original file:**", 1)
                                    elif "- **Exporter:**" in md_content:
                                        md_content = md_content.replace("- **Exporter:**", f"{url_bullet}- **Exporter:**", 1)
                                    elif "\n---\n" in md_content:
                                        md_content = md_content.replace("\n---\n", f"\n{url_bullet}\n---\n", 1)

                                if full_header_block:
                                    if md_content.startswith("<!---"):
                                        first_break = md_content.find("\n\n")
                                        if first_break != -1:
                                            md_content = md_content[:first_break + 2] + full_header_block + md_content[first_break + 2:]
                                        else:
                                            md_content = full_header_block + md_content
                                    else:
                                        md_content = full_header_block + md_content

                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(md_content)

                                # Copy media files with guaranteed collision-free resolution
                                os.makedirs(to_extended_path(type_media_dir), exist_ok=True)
                                # 1. Copy all images explicitly referenced in Markdown content
                                all_ref_imgs = re.findall(r'!\[[^\]]*\]\(([^)]+)\)', md_content)
                                for img_ref in all_ref_imgs:
                                    img_ref = img_ref.strip().replace("\\", "/")
                                    if img_ref.startswith(("http://", "https://", "data:")):
                                        continue
                                    img_name = posixpath.basename(img_ref)
                                    dst_img = type_media_dir / img_name
                                    if not dst_img.exists():
                                        candidates = []
                                        if local_media_dir and local_media_dir.is_dir():
                                            candidates.append(local_media_dir / img_name)
                                        if local_md_file:
                                            candidates.append(local_md_file.parent / img_name)
                                            candidates.append(local_md_file.parent / "media" / img_name)
                                            candidates.append(local_md_file.parent / f"{local_md_file.stem}_media" / img_name)
                                        for cand in candidates:
                                            if cand.is_file():
                                                try:
                                                    shutil.copy2(to_extended_path(cand), to_extended_path(dst_img))
                                                except Exception:
                                                    pass
                                                break

                                # 2. Copy companion media files by TDoc prefix or slide pattern
                                if local_media_dir and local_media_dir.is_dir():
                                    clean_tid = re.sub(r'[^\w\-]', '_', str(tdoc_id).strip())
                                    patterns = [
                                        f"{tdoc_id}_*",
                                        f"{sanitize_filename(tdoc_id)}_*",
                                        f"{clean_tid}_*",
                                        "slide_*.png",
                                    ]
                                    copied_names = set()
                                    for pat in patterns:
                                        for mf in local_media_dir.glob(pat):
                                            if mf.name not in copied_names:
                                                copied_names.add(mf.name)
                                                dst_img = type_media_dir / mf.name
                                                if not dst_img.exists():
                                                    try:
                                                        shutil.copy2(to_extended_path(mf), to_extended_path(dst_img))
                                                    except Exception:
                                                        pass
                            except Exception as err:
                                logger.error(f"[PreMeetingExporter] Error writing dossier for {tdoc_id}: {err}")
                        else:
                            # Fallback stub
                            try:
                                source_lines = "\n".join(CompanySanitizer.format_source_markdown_header_lines(doc.get('source', '')))
                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(
                                        f"{full_header_block}"
                                        f"# {tdoc_id}: {doc['title']}\n\n"
                                        f"{source_lines}\n"
                                        f"- **Type:** {doc['type']}\n"
                                        f"- **Agenda Item:** {doc['ai']}\n"
                                        f"- **Status:** {doc['status']}\n\n"
                                        f"> ⚠️ *Note: Document conversion was not available in local cache.*\n"
                                    )
                            except Exception:
                                pass

