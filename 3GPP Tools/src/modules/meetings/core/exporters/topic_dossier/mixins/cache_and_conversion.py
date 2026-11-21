"""Revision lineage, shared canonical cache, download, and COM conversion mixin."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from core.utils.office_printer_guard import temporary_virtual_default_printer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import ExportManifest
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    sanitize_filename,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from ..path_utils import (
    _copy_markdown_with_media,
    _normalize_tdoc_id,
    _resolve_doc_server_url,
    call_convert_document_to_markdown as convert_document_to_markdown,
    sanitize_folder_name,
)

logger = logging.getLogger(__name__)


class TopicDossierCacheAndConversionMixin:
    """Handles revision lineage mapping, canonical cache sync, FTP downloads, and Markdown conversion."""

    @staticmethod
    def _build_revision_lineage_map(
        rows: List[Dict[str, Any]],
    ) -> Tuple[Set[str], Dict[str, List[str]], Dict[str, str]]:
        """
        Builds per-meeting revision lineage relationships:
        - superseded_set: TDoc IDs with a newer active (non-withdrawn) revision in `rows`.
        - predecessors_map: latest_leaf_tdoc_id -> sorted list of superseded predecessors.
        - leaf_map: tdoc_id -> latest_active_leaf_tdoc_id.
        """
        tdoc_map = {
            _normalize_tdoc_id(r.get("TDoc", "") or r.get("tdoc_id", "") or r.get("TD#", "")): r
            for r in rows
            if _normalize_tdoc_id(r.get("TDoc", "") or r.get("tdoc_id", "") or r.get("TD#", ""))
        }
        forward_edges: Dict[str, Set[str]] = {}

        for tdoc_id, row in tdoc_map.items():
            rev_to = str(row.get("Revised to", "") or "").strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cu = _normalize_tdoc_id(cand)
                    if cu in tdoc_map and cu != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cu)

            rev_of = str(row.get("Is revision of", "") or row.get("Revision of", "") or "").strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cu = _normalize_tdoc_id(cand)
                    if cu in tdoc_map and cu != tdoc_id:
                        forward_edges.setdefault(cu, set()).add(tdoc_id)

            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        def _is_active(tid: str) -> bool:
            r = tdoc_map.get(tid)
            if not r:
                return False
            st = str(r.get("TDoc Status", "") or r.get("Status", "") or r.get("status", "")).lower()
            return "withdrawn" not in st

        leaf_map: Dict[str, str] = {}
        superseded_set: Set[str] = set()
        predecessors_map: Dict[str, List[str]] = {}

        for tdoc_id in tdoc_map:
            if not _is_active(tdoc_id):
                continue
            curr = tdoc_id
            visited = {curr}
            while True:
                children = [c for c in forward_edges.get(curr, set()) if c not in visited and _is_active(c)]
                if not children:
                    break
                children.sort(key=str)
                next_node = children[-1]
                visited.add(next_node)
                curr = next_node

            leaf_map[tdoc_id] = curr
            if curr != tdoc_id:
                superseded_set.add(tdoc_id)
                preds = predecessors_map.setdefault(curr, [])
                if tdoc_id not in preds:
                    preds.append(tdoc_id)

        for preds in predecessors_map.values():
            preds.sort()

        return superseded_set, predecessors_map, leaf_map

    def _get_canonical_tdoc_folder(self, folder_name: str, tid: str) -> Optional[Path]:
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return None
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()
        return mtg_dir / base_tdoc

    def _find_local_binary_fast(self, folder_name: str, tid: str) -> Optional[Path]:
        """Checks extracted office binaries already on disk (zero unzip, zero network I/O)."""
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir or not mtg_dir.is_dir():
            return None
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()
        for folder in [mtg_dir / tid, mtg_dir / base_tdoc, mtg_dir, mtg_dir / "Docs", mtg_dir / "Revisions"]:
            if not folder.is_dir():
                continue
            for ext in (".docx", ".docm", ".pptx", ".doc", ".ppt", ".pdf"):
                for name in (tid, base_tdoc):
                    f = folder / f"{name}{ext}"
                    if f.is_file() and f.stat().st_size > 0:
                        return f
        return None

    def _backfill_canonical_cache(self, doc: Dict[str, Any], topic_md: Path) -> None:
        """Seeds canonical `{cache_dir}/{folder_name}/{base_tdoc}/{tid}.md` if missing."""
        tid = doc["tdoc_id"]
        folder_name = doc.get("folder_name", "")
        if not folder_name:
            return
        tdoc_folder = self._get_canonical_tdoc_folder(folder_name, tid)
        if not tdoc_folder:
            return
        safe_tid = sanitize_filename(tid)
        canonical_md = tdoc_folder / f"{safe_tid}.md"
        if Path(to_extended_path(canonical_md)).is_file():
            return
        src_media = topic_md.parent / f"{topic_md.stem}_media"
        _copy_markdown_with_media(
            src_md=topic_md,
            dst_md=canonical_md,
            src_media_dir=src_media if src_media.is_dir() else None,
            src_media_rel=f"{topic_md.stem}_media",
            dst_media_dir=tdoc_folder / "media",
            dst_media_rel="media",
        )

    def _restore_from_canonical_cache(
        self,
        doc: Dict[str, Any],
        local_bin: Optional[Path],
        expected_mode: str,
    ) -> bool:
        """
        Checks canonical meeting Markdown cache before downloading/converting.
        Copies valid `.md` and rewrites `media/` links -> `{stem}_media/` into `doc["abs_path"]`.
        """
        folder_name = doc.get("folder_name", "")
        if not folder_name:
            return False
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return False

        tid = doc["tdoc_id"]
        safe_tid = sanitize_filename(tid)
        tdoc_folder = self._get_canonical_tdoc_folder(folder_name, tid)
        abs_p = doc["abs_path"]

        cand_files: List[Tuple[Path, Path, str]] = []
        if tdoc_folder and tdoc_folder.is_dir():
            cand_files.append((tdoc_folder / f"{safe_tid}.md", tdoc_folder / "media", "media"))
            if safe_tid != tid:
                cand_files.append((tdoc_folder / f"{tid}.md", tdoc_folder / "media", "media"))
            cand_files.append((tdoc_folder / f"{tid}_LLM_v{CONVERTER_VERSION}.md", tdoc_folder / "media", "media"))
        tid_dir = mtg_dir / tid
        if tid_dir.is_dir() and tid_dir != tdoc_folder:
            cand_files.append((tid_dir / f"{safe_tid}.md", tid_dir / "media", "media"))

        for mdf, m_dir, m_rel in cand_files:
            if not Path(to_extended_path(mdf)).is_file():
                continue
            if is_conversion_stale(local_bin, mdf, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                continue
            ok = _copy_markdown_with_media(
                src_md=mdf,
                dst_md=abs_p,
                src_media_dir=m_dir,
                src_media_rel=m_rel,
                dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                dst_media_rel=f"{abs_p.stem}_media",
            )
            if ok:
                # Ensure standard canonical `{safe_tid}.md` also exists if restored from legacy `_LLM_v...`
                if tdoc_folder and mdf.name != f"{safe_tid}.md":
                    canonical_md = tdoc_folder / f"{safe_tid}.md"
                    if not Path(to_extended_path(canonical_md)).is_file():
                        _copy_markdown_with_media(mdf, canonical_md, m_dir, m_rel, tdoc_folder / "media", "media")
                return True
        return False

    def _migrate_legacy_export_file(
        self,
        doc: Dict[str, Any],
        target_dir: Path,
        manifest: ExportManifest,
        local_bin: Optional[Path],
        expected_mode: str,
    ) -> bool:
        """Migrates valid Markdown files from legacy `YYYY-QN/` export layout into `{WG}/{MeetingFolder}/` + canonical cache."""
        tid = doc["tdoc_id"]
        safe_tid = sanitize_filename(tid)
        legacy_candidates: List[Path] = []

        m_entry = manifest.tdocs.get(tid)
        if m_entry and m_entry.get("rel_path") and m_entry["rel_path"] != doc["rel_path"]:
            legacy_candidates.append(target_dir / m_entry["rel_path"])

        if doc.get("cycle"):
            old_dir = f"{sanitize_folder_name(doc['wg'])}_{sanitize_folder_name(doc['meeting_number'])}"
            legacy_candidates.append(target_dir / sanitize_folder_name(doc["cycle"]) / old_dir / doc["category"] / f"{safe_tid}.md")

        abs_p = doc["abs_path"]
        for leg_p in legacy_candidates:
            if not Path(to_extended_path(leg_p)).is_file():
                continue
            if is_conversion_stale(local_bin, leg_p, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                continue
            leg_media = leg_p.parent / f"{leg_p.stem}_media"
            ok = _copy_markdown_with_media(
                src_md=leg_p,
                dst_md=abs_p,
                src_media_dir=leg_media if leg_media.is_dir() else None,
                src_media_rel=f"{leg_p.stem}_media",
                dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                dst_media_rel=f"{abs_p.stem}_media",
            )
            if ok:
                self._backfill_canonical_cache(doc, abs_p)
                return True
        return False

    # ------------------------------------------------------------------
    # Download & Two-Phase Conversion Pipeline
    # ------------------------------------------------------------------

    def _find_or_download_tdoc_file(self, tid: str, folder_name: str, meeting_dict: Dict[str, Any]) -> Optional[Path]:
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return None

        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()

        candidate_folders = [
            mtg_dir / tid,
            mtg_dir / base_tdoc,
            mtg_dir,
            mtg_dir / "Docs",
            mtg_dir / "Revisions",
        ]

        # 1. Fast direct check if already extracted in meeting folder
        for folder in candidate_folders:
            if folder.is_dir():
                for ext in (".docx", ".docm", ".pptx", ".doc", ".ppt", ".pdf"):
                    direct_f = folder / f"{tid}{ext}"
                    if direct_f.is_file() and direct_f.stat().st_size > 0:
                        return direct_f
                    if base_tdoc != tid:
                        base_f = folder / f"{base_tdoc}{ext}"
                        if base_f.is_file() and base_f.stat().st_size > 0:
                            return base_f
                    for f in folder.glob(f"*{tid}*{ext}"):
                        if not f.name.startswith("~$") and f.is_file() and f.stat().st_size > 0:
                            return f

        # 2. Check if ZIP archive already exists in tdoc_dir or mtg_dir
        candidate_zips = [
            mtg_dir / tid / f"{tid}.zip",
            mtg_dir / base_tdoc / f"{base_tdoc}.zip",
            mtg_dir / f"{tid}.zip",
            mtg_dir / f"{base_tdoc}.zip",
            mtg_dir / "Docs" / f"{tid}.zip",
            mtg_dir / "Docs" / f"{base_tdoc}.zip",
            mtg_dir / "Revisions" / f"{tid}.zip",
            mtg_dir / "Revisions" / f"{base_tdoc}.zip",
        ]
        target_extract_dir = mtg_dir / base_tdoc
        for zip_path in candidate_zips:
            if zip_path.is_file() and zip_path.stat().st_size > 0:
                try:
                    import zipfile
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        for member in zf.namelist():
                            if member.lower().endswith(('.docx', '.docm', '.pptx', '.doc', '.ppt', '.pdf')) and not os.path.basename(member).startswith('~$'):
                                zf.extract(member, target_extract_dir)
                                return target_extract_dir / member
                except Exception:
                    pass

        # 3. If missing and network download is not enabled, return None
        if not self.allow_network_download or self.is_cancelled:
            return None

        from modules.meetings.core.tdocs.tdoc_url_resolver import normalize_3gpp_url

        candidate_urls = []
        is_rev = bool(re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE))
        docs_url = meeting_dict.get("docs_folder_url")
        url_key = meeting_dict.get("url_key")

        resolved_primary = resolve_tdoc_download_url(
            tid,
            docs_url=docs_url,
            main_url=url_key,
            meta=meeting_dict,
        )
        if resolved_primary:
            candidate_urls.append(resolved_primary.rsplit('/', 1)[0])

        bases = []
        if docs_url:
            norm_docs = normalize_3gpp_url(docs_url)
            m_base = re.sub(r'/[Dd]ocs/?$', '', norm_docs, flags=re.IGNORECASE)
            bases.append(m_base)
        if url_key:
            norm_key = normalize_3gpp_url(url_key)
            bases.append(norm_key)

        for b in bases:
            if is_rev:
                candidate_urls.append(f"{b}/INBOX/Revisions")
                candidate_urls.append(f"{b}/Inbox/Revisions")
                candidate_urls.append(f"{b}/Revisions")
                candidate_urls.append(f"{b}/Docs")
            else:
                candidate_urls.append(f"{b}/Docs")
                candidate_urls.append(f"{b}/INBOX")
                candidate_urls.append(f"{b}/Inbox")

        dedup_urls = list(dict.fromkeys(u.rstrip('/') for u in candidate_urls if u))

        targets_to_try = [tid]
        if is_rev and base_tdoc != tid:
            targets_to_try.append(base_tdoc)

        for target_id in targets_to_try:
            for base_url in dedup_urls:
                if self.is_cancelled:
                    return None
                try:
                    os.makedirs(target_extract_dir, exist_ok=True)
                    docs = TDocFileHandler.download_and_extract_tdoc(
                        target_filename=target_id,
                        base_url=base_url,
                        tdoc_dir=target_extract_dir,
                        timeout=12,
                        cancel_check=lambda: self.is_cancelled,
                    )
                    if docs:
                        logger.info(f"📥 [TopicDossier] Downloaded {target_id} from {base_url}")
                        return docs[0]
                except Exception as e:
                    logger.debug(f"[Download] Failed {target_id} from {base_url}: {e}")

        return None

    def _convert_docs_pipeline(
        self,
        docs_to_convert: List[Dict[str, Any]],
        target_dir: Path,
        base_pct: int = 30,
        pct_span: int = 50,
    ):
        """Two-phase parallel pre-download and sequential conversion with canonical cache write-back."""
        total = len(docs_to_convert)
        if total == 0:
            return

        # Phase 1: Polite Concurrent Pre-Download (max_workers=5)
        self.progress_msg.emit(f"Resolving source documents ({total} files)...")
        try:
            with ThreadPoolExecutor(max_workers=5) as executor:
                self._executor = executor
                future_to_doc = {
                    executor.submit(
                        self._find_or_download_tdoc_file,
                        doc["tdoc_id"],
                        doc["folder_name"],
                        doc.get("meeting_dict", {}),
                    ): doc
                    for doc in docs_to_convert
                }
                done_count = 0
                for fut in as_completed(future_to_doc):
                    if self.is_cancelled:
                        return
                    done_count += 1
                    self.progress_step.emit(base_pct + int((done_count / total) * (pct_span * 0.45)), 100)
                    try:
                        doc = future_to_doc[fut]
                        doc["source_file"] = fut.result()
                    except Exception as e:
                        logger.debug(f"[PreDownload] Failed item: {e}")
        finally:
            self._executor = None

        # Phase 2: Sequential Conversion into Canonical Cache + Write-Out to Topic Destination
        # Wrap the entire batch once in temporary_virtual_default_printer so Windows default printer
        # isn't toggled repeatedly across 1000+ conversions.
        with temporary_virtual_default_printer(log=logger):
            for idx, doc in enumerate(docs_to_convert):
                if self.is_cancelled:
                    return
                self._convert_or_stub_single_doc(doc, idx, total, base_pct, pct_span)

    def _convert_or_stub_single_doc(
        self,
        doc: Dict[str, Any],
        idx: int,
        total: int,
        base_pct: int,
        pct_span: int,
    ) -> None:
        """Converts one downloaded/cached TDoc via shared canonical cache or writes a metadata stub."""
        tid = doc["tdoc_id"]
        abs_p = doc["abs_path"]
        src_file = doc.get("source_file")
        mtg_dict = doc.get("meeting_dict", {})
        row = doc.get("row_data", {})
        expected_mode = resolve_revision_mode_for_type(doc.get("type", ""), default_mode="markup")

        self.progress_msg.emit(f"Converting ({idx + 1}/{total}): {tid}...")
        self.progress_step.emit(base_pct + int((0.45 + 0.55 * (idx / total)) * pct_span), 100)

        local_bin = Path(src_file) if src_file else None
        converted_ok = self._restore_from_canonical_cache(doc, local_bin=local_bin, expected_mode=expected_mode)

        if not converted_ok and src_file and os.path.exists(src_file):
            tdoc_url = _resolve_doc_server_url(tid, mtg_dict, row)
            merged_meta = dict(row or {})
            merged_meta.setdefault("docs_folder_url", mtg_dict.get("docs_folder_url"))
            merged_meta.setdefault("url_key", mtg_dict.get("url_key"))
            if doc.get("supersedes"):
                merged_meta["Supersedes"] = ", ".join(doc["supersedes"])

            tdoc_folder = self._get_canonical_tdoc_folder(doc.get("folder_name", ""), tid)
            try:
                if tdoc_folder:
                    # Convert into shared canonical cache first, then copy with rewritten media links
                    os.makedirs(to_extended_path(tdoc_folder), exist_ok=True)
                    canon_media = tdoc_folder / "media"
                    res_doc = convert_document_to_markdown(
                        doc_path=Path(src_file),
                        out_dir=tdoc_folder,
                        tdoc_id=tid,
                        meta=merged_meta,
                        revisions=expected_mode,
                        media_dir=canon_media,
                        media_rel="media",
                        tdoc_url=tdoc_url,
                        cancellation_check=lambda: self.is_cancelled,
                    )
                    canon_md = Path(res_doc.md_path) if res_doc and getattr(res_doc, "md_path", None) else (tdoc_folder / f"{sanitize_filename(tid)}.md")
                    converted_ok = _copy_markdown_with_media(
                        src_md=canon_md,
                        dst_md=abs_p,
                        src_media_dir=canon_media,
                        src_media_rel="media",
                        dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                        dst_media_rel=f"{abs_p.stem}_media",
                    )
                else:
                    # Fallback when no meeting cache folder exists (e.g. unit test mocks)
                    os.makedirs(to_extended_path(abs_p.parent), exist_ok=True)
                    media_dir = abs_p.parent / f"{abs_p.stem}_media"
                    convert_document_to_markdown(
                        doc_path=Path(src_file),
                        out_dir=abs_p.parent,
                        tdoc_id=tid,
                        meta=merged_meta,
                        revisions=expected_mode,
                        media_dir=media_dir,
                        media_rel=f"{abs_p.stem}_media",
                        tdoc_url=tdoc_url,
                        cancellation_check=lambda: self.is_cancelled,
                    )
                    converted_ok = abs_p.is_file() and abs_p.stat().st_size > 0
            except Exception as e:
                logger.warning(f"[Convert] Failed to convert {tid}: {e}")

        if not converted_ok and not abs_p.is_file():
            try:
                os.makedirs(to_extended_path(abs_p.parent), exist_ok=True)
                src_url = _resolve_doc_server_url(tid, mtg_dict, row)
                source_lines = f"- **Source URL:** [{src_url}]({src_url})" if src_url else ""
                stub_content = (
                    f"# {tid}: {doc.get('title', '')}\n\n"
                    f"- **Working Group:** {doc.get('wg', '')}\n"
                    f"- **Meeting:** {doc.get('meeting_number', '')}\n"
                    f"- **Agenda Item:** {doc.get('ai', '')}\n"
                    f"- **Work Item:** {doc.get('wi', '')}\n"
                    f"- **Type:** {doc.get('type', '')}\n"
                    f"- **Status:** {doc.get('status', '')}\n"
                    f"- **Source / Company:** {doc.get('source', '')}\n"
                    f"{source_lines}\n\n"
                    f"> [!NOTE]\n"
                    f"> *Document conversion was not available in local cache at export time.*\n\n"
                    f"### Abstract / Remarks\n"
                    f"{row.get('Abstract', '') or row.get('Secretary Remarks', 'No abstract provided.')}\n"
                )
                with open(to_extended_path(abs_p), "w", encoding="utf-8") as sf:
                    sf.write(stub_content)
            except Exception as stub_err:
                logger.warning(f"[Stub] Failed to create stub for {tid}: {stub_err}")
