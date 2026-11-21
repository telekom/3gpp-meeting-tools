"""Proactive Drafts caching, document download/open, and LLM workspace preparation."""

from __future__ import annotations

import hashlib
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from modules.emails.core.thread_markdown_exporter import (
    extract_drafts_document_urls,
    extract_thread_document_urls,
    is_drafts_document_url,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

logger = logging.getLogger(__name__)


def _build_candidate_document_urls(url: str) -> List[str]:
    """Builds candidate download URLs translating between 10.10.10.10 local server and 3gpp.org/ftp."""
    candidates = [url]
    # 10.10.10.10/ftp/SA/SA2/Inbox/... -> https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/SA2/Inbox/...
    m_local = re.match(r"^https?://10\.10\.10\.10/ftp/(?:SA|RAN|CT)/([A-Z0-9]+)/(.+)$", url, re.IGNORECASE)
    if m_local:
        wg, rest = m_local.group(1).upper(), m_local.group(2)
        candidates.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{wg}/{rest}")
    # https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_.../INBOX/... -> http://10.10.10.10/ftp/SA/SA2/Inbox/...
    m_remote = re.match(r"^https?://(?:www\.)?3gpp\.org/ftp/tsg_(sa|ran|ct)/[^/]+/[^/]+/(.+)$", url, re.IGNORECASE)
    if m_remote:
        tsg = m_remote.group(1).upper()
        rest = m_remote.group(2)
        if tsg == "SA":
            candidates.insert(0, f"http://10.10.10.10/ftp/SA/SA2/{rest}")
    return list(dict.fromkeys(candidates))


def get_drafts_cache_dir(agenda_dir: Path) -> Path:
    """Returns canonical DraftsCache directory under `<meeting_dir>/Agenda/DraftsCache`."""
    return Path(agenda_dir) / "DraftsCache"


def _infer_agenda_dir_from_cache_dir(cache_dir: Optional[Path]) -> Optional[Path]:
    if not cache_dir:
        return None
    p = Path(cache_dir)
    if p.name == "InboxAttachments":
        return p.parent
    if p.parent.name in ("ThreadAttachments", "TDocEmailCache"):
        return p.parent.parent
    if p.name == "DraftsCache":
        return p.parent
    return None


def resolve_draft_cache_target(url: str, drafts_root: Path) -> Path:
    """
    Maps a 10.10.10.10 / Drafts URL to a deterministic local file path inside `drafts_root`,
    preserving subfolders after `/Drafts/` to prevent same-filename collisions across KIs.
    """
    import urllib.parse

    drafts_root = Path(drafts_root)
    parsed = urllib.parse.urlsplit((url or "").strip())
    raw_segs = [urllib.parse.unquote(s) for s in parsed.path.split("/") if s]
    low_segs = [s.strip().lower() for s in raw_segs]
    rel_segs = raw_segs[low_segs.index("drafts") + 1 :] if "drafts" in low_segs else raw_segs[-1:]

    clean_segs = []
    for seg in rel_segs:
        cleaned = "".join(c for c in seg if c not in '<>:"/\\|?*').strip()
        if cleaned and cleaned not in (".", ".."):
            clean_segs.append(cleaned)

    if not clean_segs:
        clean_segs = ["draft_document.docx"]

    candidate = drafts_root.joinpath(*clean_segs)
    if len(str(candidate)) < 220:
        return candidate
    h = hashlib.sha1(url.encode("utf-8", errors="ignore")).hexdigest()[:8]
    return drafts_root / f"{h}_{clean_segs[-1]}"


def _download_document_candidates(
    url: str,
    target_path: Path,
    local_ip_unreachable: Optional[List[bool]] = None,
) -> bool:
    """Downloads a 3GPP / 10.10.10.10 URL to `target_path` with fast 10.10.10.10 connect timeout & failover."""
    from core.network.session import NetworkSession

    session = NetworkSession()
    target_path.parent.mkdir(parents=True, exist_ok=True)
    skip_local = bool(local_ip_unreachable and local_ip_unreachable[0])

    for cand_url in _build_candidate_document_urls(url):
        is_local = "://10.10.10.10/" in cand_url.lower()
        if is_local and skip_local:
            continue
        timeout = (1.5, 12) if is_local else (4, 15)
        try:
            resp = session.get(cand_url, timeout=timeout)
            if resp.status_code == 200 and resp.content:
                target_path.write_bytes(resp.content)
                return True
        except Exception as e:
            if is_local and local_ip_unreachable is not None:
                local_ip_unreachable[0] = True
                skip_local = True
            logger.debug(f"Candidate URL download failed ({cand_url}): {e}")
    return False


def cache_drafts_for_messages(
    messages: List[Dict[str, Any]],
    agenda_dir: Path,
    progress_cb: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[Path]:
    """
    Proactively downloads uncached 10.10.10.10 /Drafts/ document URLs referenced by emails into
    `<agenda_dir>/DraftsCache/`, immediately skipping files that already exist locally.
    """
    draft_urls = extract_drafts_document_urls(messages, prefer_clean_body=True)
    if not draft_urls:
        return []

    drafts_root = get_drafts_cache_dir(agenda_dir)
    pending: List[tuple] = []
    cached_paths: List[Path] = []
    for u in draft_urls:
        tgt = resolve_draft_cache_target(u, drafts_root)
        if tgt.exists() and tgt.stat().st_size > 0:
            cached_paths.append(tgt)
        else:
            pending.append((u, tgt))

    if not pending:
        return cached_paths

    local_unreachable = [False]
    total = len(pending)
    for idx, (u, tgt) in enumerate(pending, 1):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(f"Caching draft ({idx}/{total}): {tgt.name[:32]}")
        if _download_document_candidates(u, tgt, local_ip_unreachable=local_unreachable):
            cached_paths.append(tgt)
    return cached_paths


def _auto_cache_db_drafts(
    db: Any,
    messages: List[Dict[str, Any]],
    progress_cb: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[Path]:
    db_path = getattr(db, "db_path", None)
    if not db_path or str(db_path) == ":memory:":
        return []
    agenda_dir = Path(db_path).parent
    return cache_drafts_for_messages(messages, agenda_dir, progress_cb=progress_cb, cancel_check=cancel_check)


def cache_and_open_thread_document_url(
    url: str,
    thread_dir: Path,
    open_after: bool = True,
    drafts_root: Optional[Path] = None,
) -> Optional[Path]:
    """
    Downloads and caches a document linked in an email thread (e.g. 10.10.10.10 or 3gpp.org/ftp)
    into `DraftsCache/` (for Drafts URLs) or `thread_dir` and optionally opens it via sensitivity handler.
    """
    import urllib.parse
    import zipfile

    if not url:
        return None
    thread_dir = Path(thread_dir)
    thread_dir.mkdir(parents=True, exist_ok=True)

    raw_name = urllib.parse.unquote(urllib.parse.urlsplit(url).path.rsplit("/", 1)[-1]) or "document.docx"
    safe_name = "".join(c for c in raw_name if c not in '<>:"/\\|?*').strip() or "document.docx"
    flat_target = thread_dir / safe_name

    eff_drafts_root = drafts_root
    if eff_drafts_root is None:
        inferred_agenda = _infer_agenda_dir_from_cache_dir(thread_dir)
        if inferred_agenda is not None:
            eff_drafts_root = get_drafts_cache_dir(inferred_agenda)

    if is_drafts_document_url(url) and eff_drafts_root is not None:
        draft_target = resolve_draft_cache_target(url, eff_drafts_root)
        if draft_target.exists() and draft_target.stat().st_size > 0:
            target_path = draft_target
        elif flat_target.exists() and flat_target.stat().st_size > 0:
            target_path = flat_target
        else:
            target_path = draft_target
    else:
        target_path = flat_target

    if not (target_path.exists() and target_path.stat().st_size > 0):
        if not _download_document_candidates(url, target_path):
            return None

    open_target = target_path
    if target_path.suffix.lower() == ".zip" and zipfile.is_zipfile(target_path):
        try:
            with zipfile.ZipFile(target_path, "r") as zf:
                for member in zf.namelist():
                    if member.lower().endswith((".docx", ".doc", ".pptx", ".pdf", ".xlsx")) and not Path(member).name.startswith("~$"):
                        extracted = target_path.parent / Path(member).name
                        if not extracted.exists():
                            extracted.write_bytes(zf.read(member))
                        open_target = extracted
                        break
        except Exception:
            pass

    if open_after:
        try:
            from modules.word_tools.core.document_opener import open_document
            open_document(open_target.resolve())
        except Exception as e:
            logger.error(f"Could not open cached document {open_target}: {e}")
    return open_target


def cache_all_thread_documents_and_attachments(
    db: TrackedThreadsDatabase,
    thread_id: str,
    thread_dir: Path,
) -> List[Path]:
    """Caches all linked documents (10.10.10.10 / 3GPP FTP) and Outlook attachments for a thread into thread_dir."""
    messages = db.get_thread_messages(thread_id)
    cached: List[Path] = []
    thread_dir = Path(thread_dir)
    thread_dir.mkdir(parents=True, exist_ok=True)

    # 1. Linked URLs in message bodies
    for doc_url in extract_thread_document_urls(messages):
        p = cache_and_open_thread_document_url(doc_url, thread_dir, open_after=False)
        if p:
            cached.append(p)

    # 2. Outlook attachments
    from modules.emails.core.outlook_client import OutlookClient
    ns = OutlookClient.get_namespace()
    if ns:
        for msg in messages:
            eid = str(msg.get("id") or "").strip()
            atts_names = msg.get("attachments") or []
            if not eid or eid.startswith("MSG-") or not atts_names:
                continue
            try:
                item = ns.GetItemFromID(eid)
                atts = getattr(item, "Attachments", None)
                if not atts:
                    continue
                for i in range(1, atts.Count + 1):
                    att = atts.Item(i)
                    fname = getattr(att, "FileName", "") or ""
                    if not fname or fname.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                        continue
                    safe_name = "".join(c for c in fname if c not in '<>:"/\\|?*').strip()
                    if safe_name:
                        out_path = thread_dir / safe_name
                        if not out_path.exists():
                            att.SaveAsFile(str(out_path.resolve()))
                        cached.append(out_path)
            except Exception:
                continue
    return cached


def prepare_thread_llm_workspace(
    db: TrackedThreadsDatabase,
    thread_id: str,
    cache_dir: Path,
    deduplicate_quotes: bool = True,
    newest_first: bool = False,
    cache_linked_docs: bool = True,
) -> Dict[str, Any]:
    """
    Prepares a self-contained LLM workspace directory containing thread.md plus any
    cached linked documents/attachments converted to Markdown when possible.
    Returns {'workspace_dir': Path, 'thread_md': Path, 'doc_mds': List[Path]}.
    """
    from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown

    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    ws_dir = cache_dir / "llm_workspace"
    ws_dir.mkdir(parents=True, exist_ok=True)

    thread = db.get_thread(thread_id) or {"thread_id": thread_id}
    messages = db.get_thread_messages(thread_id)
    md_text = export_thread_to_markdown(
        thread,
        messages,
        deduplicate_quotes=deduplicate_quotes,
        interactive_links=False,
        newest_first=newest_first,
    )
    thread_md = ws_dir / f"{thread_id}.md"
    thread_md.write_text(md_text, encoding="utf-8")

    cached_files: List[Path] = []
    if cache_linked_docs:
        try:
            cached_files = cache_all_thread_documents_and_attachments(db, thread_id, cache_dir)
        except Exception as e:
            logger.debug(f"Could not cache linked docs during LLM workspace prep: {e}")

    # Also include any already-cached .docx/.docm/.md files in cache_dir
    for existing in cache_dir.iterdir():
        if existing.is_file() and existing not in cached_files:
            cached_files.append(existing)

    doc_mds: List[Path] = []
    for fpath in cached_files:
        if fpath.suffix.lower() in (".docx", ".docm"):
            out_md = ws_dir / f"{fpath.stem}.md"
            if not out_md.exists():
                try:
                    from modules.meetings.core.tdoc_markdown_converter import convert_document_to_markdown
                    convert_document_to_markdown(str(fpath), str(out_md))
                except Exception as e:
                    logger.debug(f"Markdown conversion skipped for {fpath.name}: {e}")
            if out_md.exists():
                doc_mds.append(out_md)
        elif fpath.suffix.lower() == ".md" and fpath.parent != ws_dir:
            doc_mds.append(fpath)

    return {"workspace_dir": ws_dir, "thread_md": thread_md, "doc_mds": doc_mds}
