"""Shared path normalization, media copy, checkpoint key, and shim-forwarding utilities."""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import sys
import urllib.parse
from pathlib import Path
from typing import Any, Dict, Optional, Set

from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.topic_digest_builder import rewrite_media_links
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    convert_document_to_markdown as _ORIG_CONVERT_DOCUMENT_TO_MARKDOWN,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.topic_decision import TopicDecision

logger = logging.getLogger(__name__)


def _normalize_tdoc_id(raw_tid: str) -> str:
    """Normalizes TDoc IDs to 3GPP convention: uppercase prefix/number + lowercase revision suffix ('R2-2600999r01')."""
    s = str(raw_tid or "").strip().upper()
    return re.sub(
        r'^([A-Z0-9]+-\d+)-?(?:R|REV)(\d{1,2}[A-Z]?)$',
        lambda m: f"{m.group(1)}r{m.group(2).lower()}",
        s,
    )


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem characters."""
    clean = re.sub(r'[\\/*?:"<>|#%]', "_", str(name or "")).strip()
    return re.sub(r'_+', '_', clean).strip("._ ") or "General"


def safe_relpath(target: Path, base: Path) -> str:
    """Computes clean relative path between two paths, safe for Windows extended paths."""
    try:
        t_str = str(target).replace("\\\\?\\", "").replace("//?/", "")
        b_str = str(base).replace("\\\\?\\", "").replace("//?/", "")
        t_abs = Path(t_str).resolve()
        b_abs = Path(b_str).resolve()
        return os.path.relpath(t_abs, b_abs).replace("\\", "/")
    except Exception:
        return str(target).replace("\\", "/")


def safe_markdown_rel_link(target: Path, base: Path) -> str:
    """Computes URL-encoded relative path for Markdown links."""
    rel = safe_relpath(target, base)
    clean_path = str(rel).replace("\\", "/")
    if "#" in clean_path:
        p_part, frag = clean_path.split("#", 1)
        return f"{urllib.parse.quote(p_part, safe='/')}#{urllib.parse.quote(frag, safe='')}"
    return urllib.parse.quote(clean_path, safe="/")


def get_default_topic_export_base() -> Path:
    """
    Returns the default base directory for cross-WG topic exports,
    located in the user's meeting cache directory rather than the code repository.
    """
    try:
        from modules.meetings.core.settings import MeetingsSettings
        settings = MeetingsSettings()
        if settings.cache_dir and Path(settings.cache_dir).is_dir():
            return Path(settings.cache_dir) / "Export"
    except Exception:
        pass
    from core.utils.paths import get_cache_root
    cache_root = get_cache_root()
    if cache_root and Path(cache_root).is_dir():
        return Path(cache_root) / "Export"
    return Path.home() / "3GPP_Delegate_Helper" / "cache" / "Export"


def _resolve_doc_server_url(
    tid: str,
    meeting_dict: Optional[Dict[str, Any]] = None,
    row_data: Optional[Dict[str, Any]] = None,
) -> str:
    """Builds the authoritative 3GPP HTTPS archive URL using meeting docs/main URLs."""
    mtg = meeting_dict or {}
    merged_meta: Dict[str, Any] = {}
    merged_meta.update(mtg)
    if row_data:
        merged_meta.update(row_data)
    return resolve_tdoc_download_url(
        tid,
        docs_url=mtg.get("docs_folder_url"),
        main_url=mtg.get("url_key"),
        meta=merged_meta,
    )


def _decision_checkpoint_key(meeting_folder_id: str, dec: TopicDecision) -> str:
    """Generates a deterministic manifest key for a TopicDecision across incremental sessions."""
    if dec.tdoc_id and not dec.is_text_agreement:
        return f"{meeting_folder_id}::{dec.decision_type}::{dec.tdoc_id}"
    digest = hashlib.sha1(dec.title_or_summary.strip().encode("utf-8")).hexdigest()[:10]
    return f"{meeting_folder_id}::{dec.decision_type}::{dec.tdoc_id or dec.ai or 'AI'}::{digest}"


def _copy_markdown_with_media(
    src_md: Path,
    dst_md: Path,
    src_media_dir: Optional[Path],
    src_media_rel: str,
    dst_media_dir: Path,
    dst_media_rel: str,
) -> bool:
    """
    Copies a Markdown file to `dst_md`, rewriting relative image/media links from
    `src_media_rel/` to `dst_media_rel/` and copying referenced media assets cleanly.
    """
    try:
        src_ext = Path(to_extended_path(src_md))
        dst_ext = Path(to_extended_path(dst_md))
        if not src_ext.is_file():
            return False

        raw_text = src_ext.read_text(encoding="utf-8", errors="replace")
        rewritten = rewrite_media_links(raw_text, src_media_rel, dst_media_rel)

        os.makedirs(dst_ext.parent, exist_ok=True)
        dst_ext.write_text(rewritten, encoding="utf-8")

        if src_media_dir and Path(to_extended_path(src_media_dir)).is_dir():
            src_m_ext = Path(to_extended_path(src_media_dir))
            # Determine which media filenames are referenced in rewritten markdown
            refs: Set[str] = set()
            for m in re.finditer(r'!\[[^\]]*\]\(([^)]+)\)|src=["\']([^"\']+)["\']', raw_text):
                ref = (m.group(1) or m.group(2) or "").strip().replace("\\", "/")
                if ref and not ref.startswith(("http://", "https://", "data:")):
                    refs.add(os.path.basename(ref))

            copied_any = False
            dst_m_ext = Path(to_extended_path(dst_media_dir))
            candidates = (
                [src_m_ext / name for name in refs if (src_m_ext / name).is_file()]
                if refs
                else [f for f in src_m_ext.glob("*.*") if f.is_file()]
            )
            if candidates:
                os.makedirs(dst_m_ext, exist_ok=True)
                for mf in candidates:
                    shutil.copy2(mf, dst_m_ext / mf.name)
                    copied_any = True
            elif not refs:
                # Even if markdown contained no media refs, avoid leaving empty folders
                pass
            _ = copied_any

        return dst_ext.is_file() and dst_ext.stat().st_size > 0
    except Exception as e:
        logger.debug(f"[CopyWithMedia] Failed copying {src_md} -> {dst_md}: {e}")
        return False


def call_convert_document_to_markdown(*args: Any, **kwargs: Any) -> Any:
    """Calls convert_document_to_markdown, honoring monkeypatches on the legacy topic_dossier_exporter shim."""
    for mod_name in (
        "modules.meetings.core.exporters.topic_dossier_exporter",
        "src.modules.meetings.core.exporters.topic_dossier_exporter",
    ):
        shim = sys.modules.get(mod_name)
        if shim is not None:
            shim_fn = getattr(shim, "convert_document_to_markdown", None)
            if shim_fn is not None and shim_fn is not _ORIG_CONVERT_DOCUMENT_TO_MARKDOWN:
                return shim_fn(*args, **kwargs)
    return _ORIG_CONVERT_DOCUMENT_TO_MARKDOWN(*args, **kwargs)
