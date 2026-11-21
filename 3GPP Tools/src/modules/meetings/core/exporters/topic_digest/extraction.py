# --- File: src/modules/meetings/core/exporters/topic_digest/extraction.py ---
"""Deterministic folder IDs, media rewriting, contributor matrix, and Markdown digest extraction."""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import sanitize_folder_name
from modules.meetings.core.topic_decision import TopicDecision

from .constants import _SUFFIX_IN_NUM_RE


def resolve_meeting_folder_id(mtg: Dict[str, Any]) -> str:
    """
    Produces a collision-free, human-readable meeting folder identifier:
    - 'TSGR3_133'      -> 'RAN3_133'
    - 'TSGR3_133-bis'  -> 'RAN3_133-bis'
    - 'TSGR2_133bis'   -> 'RAN2_133bis' (or 'RAN2_133-bis')
    - 'TSGS2_175_AHE'  -> 'SA2_175-AHE'
    """
    wg = sanitize_folder_name(str(mtg.get("wg_name", "")).strip().upper() or "WG")
    raw_num = str(mtg.get("meeting_number", "")).strip()
    folder_name = str(mtg.get("folder_name", "")).strip()

    clean_num = sanitize_folder_name(raw_num or "0")
    if raw_num and not _SUFFIX_IN_NUM_RE.search(raw_num) and folder_name:
        pat = re.compile(
            r'(?:^|[_-])' + re.escape(raw_num) + r'([_-]?(?:bis|ahe|adhoc|ah\d*|e\d*)\b)',
            re.IGNORECASE,
        )
        m = pat.search(folder_name)
        if m:
            suffix = m.group(1).lstrip("_-")
            suffix_clean = sanitize_folder_name(suffix)
            if suffix_clean:
                return f"{wg}_{clean_num}-{suffix_clean}"

    return f"{wg}_{clean_num}"


def rewrite_media_links(md_text: str, old_media_rel: str, new_media_rel: str) -> str:
    """Rewrites relative Markdown & HTML image/media references when copying cached .md files."""
    if not md_text or not old_media_rel or not new_media_rel or old_media_rel == new_media_rel:
        return md_text

    pattern_md = re.compile(r'(\]\()' + re.escape(old_media_rel) + r'(/[^\)\s]+)')
    updated = pattern_md.sub(rf'\g<1>{new_media_rel}\g<2>', md_text)

    pattern_html = re.compile(r'(src=["\'])' + re.escape(old_media_rel) + r'(/[^"\']+)')
    updated = pattern_html.sub(rf'\g<1>{new_media_rel}\g<2>', updated)
    return updated


def _rel_to_meeting_dir(rel_from_root: str, wg: str, meeting_folder_id: str) -> str:
    """Converts root-relative path '{wg}/{meeting_folder_id}/Cat/File.md' to meeting-relative 'Cat/File.md'."""
    if not rel_from_root:
        return ""
    norm = str(rel_from_root).replace("\\", "/")
    prefix = f"{sanitize_folder_name(wg)}/{meeting_folder_id}/"
    if norm.startswith(prefix):
        return norm[len(prefix):]
    return norm


def build_contributor_matrix(
    candidates: List[Dict[str, Any]],
    decisions: Optional[List[TopicDecision]] = None,
) -> List[Dict[str, Any]]:
    """
    Aggregates company participation within a meeting using CompanySanitizer:
    separates Lead Proponents vs. Confirmed Co-Signers (ignoring tentative '?' / '[...]' carry-overs).
    """
    company_stats: Dict[str, Dict[str, Any]] = {}

    def _ensure(comp: str) -> Dict[str, Any]:
        if comp not in company_stats:
            company_stats[comp] = {
                "company": comp,
                "lead_count": 0,
                "cosigned_count": 0,
                "approved_count": 0,
                "discussion_count": 0,
                "ais": set(),
                "tdocs": [],
            }
        return company_stats[comp]

    for doc in candidates:
        tid = doc.get("tdoc_id", "")
        src = str(doc.get("source", "") or "")
        ai = str(doc.get("ai", "") or "").strip()
        status_l = str(doc.get("status", "") or "").lower()
        is_approved = any(w in status_l for w in ("approved", "agreed", "endorsed"))

        primaries = CompanySanitizer.get_primary_contributors(src, include_tentative=False)
        confirmed, _tentative = CompanySanitizer.parse_contributors_detailed(src)

        if not primaries and confirmed:
            primaries = [confirmed[0]]

        lead_set = set(primaries)
        for p in primaries:
            st = _ensure(p)
            st["lead_count"] += 1
            if is_approved:
                st["approved_count"] += 1
            else:
                st["discussion_count"] += 1
            if ai:
                st["ais"].add(ai)
            if tid and tid not in st["tdocs"]:
                st["tdocs"].append(tid)

        for c in confirmed:
            if c not in lead_set:
                st = _ensure(c)
                st["cosigned_count"] += 1
                if ai:
                    st["ais"].add(ai)
                if tid and tid not in st["tdocs"]:
                    st["tdocs"].append(tid)

    rows = list(company_stats.values())
    rows.sort(
        key=lambda r: (
            -(r["lead_count"] + r["cosigned_count"]),
            -r["lead_count"],
            r["company"],
        )
    )
    return rows


def _strip_markdown_frontmatter(text: str) -> str:
    """Removes tdoc_converter HTML comments and top bullet metadata block before the document body."""
    cleaned = re.sub(r'^<!---\s*tdoc_converter:.*?-->\s*', '', text, flags=re.IGNORECASE | re.DOTALL)
    # Find the separator line '---' that tdoc2md places after the metadata header
    lines = cleaned.splitlines()
    for idx, line in enumerate(lines[:45]):
        if line.strip() == "---":
            return "\n".join(lines[idx + 1:]).strip()
    # Fallback: skip leading '# TDoc' and '- **Key:**' bullets
    start_idx = 0
    for idx, line in enumerate(lines[:35]):
        s = line.strip()
        if idx == 0 and s.startswith("# "):
            start_idx = idx + 1
        elif s.startswith("- **") or not s:
            start_idx = idx + 1
        else:
            break
    return "\n".join(lines[start_idx:]).strip()


def extract_tdoc_position_digest(
    md_path: Optional[Path],
    max_intro_chars: int = 1200,
    max_proposal_chars: int = 2000,
    max_changes_chars: int = 1200,
    max_fallback_chars: int = 2200,
) -> Dict[str, str]:
    """
    Deterministically extracts high-signal company position blocks from a converted TDoc Markdown file,
    stripping verbose specification boilerplate while preserving intro, proposals, conclusions, and key CR edits.
    """
    result = {"intro": "", "proposals": "", "changes": "", "fallback": ""}
    if not md_path:
        return result

    ext_path = Path(to_extended_path(md_path))
    if not ext_path.is_file():
        return result

    try:
        raw_text = ext_path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return result

    if "Document conversion was not available" in raw_text[:2048]:
        return result

    body = _strip_markdown_frontmatter(raw_text)
    if not body:
        return result

    # Strip embedded base64 data URIs or massive markup legend headers
    body = re.sub(r'!\[[^\]]*\]\(data:image/[^)]+\)', '[Image]', body)
    body = re.sub(r'>\s*\*\*Markup Legend:\*\*.*?(?:\n\n|\Z)', '', body, flags=re.IGNORECASE | re.DOTALL).strip()

    # Split body into (heading, text) sections
    sections: List[Tuple[str, str]] = []
    cur_heading = ""
    cur_lines: List[str] = []
    for line in body.splitlines():
        if re.match(r'^#{1,5}\s+\S', line) or re.match(r'^\*\*(?:Proposal|Observation|Conclusion|Way Forward|Summary|Reason for change)[^*]*\*\*\s*:?$', line.strip(), re.IGNORECASE):
            if cur_lines or cur_heading:
                sections.append((cur_heading, "\n".join(cur_lines).strip()))
            cur_heading = line.strip().lstrip("#").strip("* ").strip()
            cur_lines = []
        else:
            cur_lines.append(line)
    if cur_lines or cur_heading:
        sections.append((cur_heading, "\n".join(cur_lines).strip()))

    intro_parts: List[str] = []
    proposal_parts: List[str] = []
    intro_kw = ("introduction", "abstract", "background", "scope", "reason for change", "overview", "general")
    prop_kw = ("proposal", "observation", "conclusion", "way forward", "recommendation", "solution", "summary", "discussion", "agreement")

    for h, content in sections:
        if not content:
            continue
        hl = h.lower()
        if any(k in hl for k in prop_kw):
            snippet = f"**{h}:**\n{content}" if h else content
            proposal_parts.append(snippet)
        elif not hl or any(k in hl for k in intro_kw):
            snippet = f"**{h}:**\n{content}" if h else content
            intro_parts.append(snippet)

    # Fallback intro from first non-empty section if nothing matched intro_kw
    if not intro_parts and sections:
        first_h, first_c = sections[0]
        if first_c:
            intro_parts.append(f"**{first_h}:**\n{first_c}" if first_h else first_c)

    intro_text = "\n\n".join(intro_parts).strip()
    prop_text = "\n\n".join(proposal_parts).strip()

    # Extract [INSERTED: ...] change highlights (for CR/pCRs)
    change_lines: List[str] = []
    for line in body.splitlines():
        if "[INSERTED:" in line or "[DELETED:" in line:
            s = line.strip()
            if s and len(s) > 12:
                change_lines.append(s)
    changes_text = "\n".join(change_lines[:25]).strip()

    if len(intro_text) > max_intro_chars:
        intro_text = intro_text[:max_intro_chars].rstrip() + " ..."
    if len(prop_text) > max_proposal_chars:
        prop_text = prop_text[:max_proposal_chars].rstrip() + " ..."
    if len(changes_text) > max_changes_chars:
        changes_text = changes_text[:max_changes_chars].rstrip() + " ..."

    result["intro"] = intro_text
    result["proposals"] = prop_text
    result["changes"] = changes_text

    if not intro_text and not prop_text and not changes_text:
        clean_body = re.sub(r'\n{3,}', '\n\n', body).strip()
        result["fallback"] = (clean_body[:max_fallback_chars].rstrip() + " ...") if len(clean_body) > max_fallback_chars else clean_body

    return result
