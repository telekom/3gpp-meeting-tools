# --- File: src/modules/specifications/core/spec_clause_attribution.py ---
"""
3GPP Clause-to-TDoc Attribution Engine.

Identifies which specific Change Request (CR) or Working Group Pseudo-CR (p-CR)
introduced or modified a given specification clause across versions.

Features:
  - Mode A: TSG Plenary Normative TS CRs (row-per-CR with explicit CR numbers).
  - Mode B: Working Group Study TR p-CRs (handling placeholder numbering e.g. 6.x,
    the Key Issue Bridge, and ancestor clause inheritance).
  - Stage 1: Zero-download metadata and structural ranking (<50ms via SQLite).
  - Stage 2: Track-changes & placeholder-agnostic OpenXML text fingerprinting.
  - Background QThread execution for non-blocking UI responsiveness.
"""

import collections
import logging
import re
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_cache_root, get_project_root
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

logger = logging.getLogger(__name__)

# Stopwords for telecommunications title and diff keyword salience
BOILERPLATE_STOPWORDS: Set[str] = {
    "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are",
    "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but",
    "by", "can", "clause", "clauses", "cr", "did", "do", "does", "doing", "down", "during",
    "each", "editor", "editors", "few", "first", "for", "from", "further", "had", "has", "have",
    "having", "he", "her", "here", "hers", "herself", "him", "himself", "his", "how", "i", "if",
    "in", "into", "is", "it", "its", "itself", "just", "key", "may", "me", "more", "most", "my",
    "myself", "next", "no", "nor", "not", "note", "notes", "now", "of", "off", "on", "once",
    "only", "or", "other", "our", "ours", "ourselves", "out", "over", "own", "pcr", "s", "same",
    "shall", "she", "should", "so", "solution", "solutions", "some", "such", "t", "table", "than",
    "that", "the", "their", "theirs", "them", "themselves", "then", "there", "these", "they",
    "this", "those", "through", "to", "too", "tr", "ts", "under", "until", "up", "update",
    "updates", "very", "was", "we", "were", "what", "when", "where", "which", "while", "who",
    "whom", "why", "will", "with", "would", "you", "your", "yours", "yourself", "yourselves"
}


class SpecClauseAttributionEngine:
    """Multi-stage attribution engine linking 3GPP clause modifications to originating CRs/p-CRs."""

    def __init__(
        self,
        history_db: Optional[SpecHistoryDatabase] = None,
        search_db: Optional[Any] = None,
        report_gen: Optional[SpecTDocReportGenerator] = None,
    ):
        self.history_db = history_db or SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
        self.search_db = search_db
        self.report_gen = report_gen or SpecTDocReportGenerator(history_db=self.history_db)

    # -------------------------------------------------------------------------
    # Pre-Processing Helpers
    # -------------------------------------------------------------------------

    @staticmethod
    def clean_tokens(text: str) -> List[str]:
        """Extracts lowercase alphabetic/alphanumeric tokens, discarding common stopwords."""
        raw_words = re.findall(r"\b[A-Za-z0-9_-]{2,}\b", text.lower())
        return [w for w in raw_words if w not in BOILERPLATE_STOPWORDS and not w.isdigit()]

    @staticmethod
    def extract_key_issue_references(text: str) -> List[str]:
        """Detects Key Issue references (e.g. 'Key Issue #4', 'KI #4', 'KI5') in clause text."""
        matches = re.findall(r"(?:Key\s+Issue|KI)\s*#?\s*(\d+)", text, re.IGNORECASE)
        unique_nums = sorted(list(set(matches)), key=lambda x: int(x))
        return [f"KI #{num}" for num in unique_nums]

    @staticmethod
    def extract_salient_diff_keywords(diff_text: str, max_keywords: int = 15) -> List[str]:
        """Extracts high-entropy technical terms and acronyms from added (+) diff lines."""
        added_lines = [
            line[1:].strip() for line in diff_text.splitlines()
            if line.startswith("+") and not line.startswith("+++")
        ]
        joined = " ".join(added_lines)
        if not joined:
            return []

        # Find telecom tokens: hyphenated identifiers, all-caps acronyms, camelCase, and words >= 4 chars
        tokens = re.findall(r"\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)+\b|\b[A-Z]{2,}\b|\b[a-zA-Z0-9_]{4,}\b", joined)
        filtered = [
            t for t in tokens
            if t.lower() not in BOILERPLATE_STOPWORDS and not t.isdigit()
        ]
        # Count frequencies and return most frequent distinctive tokens
        counter = collections.Counter(filtered)
        return [tok for tok, _ in counter.most_common(max_keywords)]

    def resolve_ancestor_context(
        self,
        spec_number: str,
        version: str,
        clause_number: str
    ) -> Tuple[str, List[str]]:
        """
        Walks up the clause tree using 3gpp_spec_search.db to resolve ancestor titles.
        Returns (parent_root_title, list_of_ancestor_titles).
        """
        if not self.search_db or not hasattr(self.search_db, "get_connection"):
            return "", []

        ancestor_titles: List[str] = []
        curr_clause = clause_number.strip()

        try:
            with self.search_db.get_connection() as conn:
                while "." in curr_clause:
                    parts = curr_clause.rsplit(".", 1)
                    parent_clause = parts[0]
                    row = conn.execute(
                        "SELECT title, text FROM spec_clauses WHERE spec = ? AND version = ? AND clause = ? LIMIT 1",
                        (spec_number, version, parent_clause)
                    ).fetchone()
                    if row and row[0]:
                        t_str = row[0].strip()
                        if t_str and t_str not in ancestor_titles:
                            ancestor_titles.append(t_str)
                    curr_clause = parent_clause
        except Exception as e:
            logger.debug(f"Could not resolve ancestor context for {clause_number}: {e}")

        root_title = ancestor_titles[-1] if ancestor_titles else ""
        return root_title, ancestor_titles

    # -------------------------------------------------------------------------
    # Stage 1: Zero-Download Metadata & Semantic Ranking
    # -------------------------------------------------------------------------

    def attribute_clause(
        self,
        spec_number: str,
        clause_number: str,
        target_version: str,
        base_version: Optional[str] = None,
        clause_title: str = "",
        diff_text: str = "",
        clause_content: str = "",
        is_new_clause: Optional[bool] = None,
        progress_cb: Optional[Callable[[int, str], None]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Executes multi-stage attribution identifying which CR or p-CR modified or added the clause.
        Returns a list of candidate dictionaries ranked descending by confidence score.
        """
        clean_spec = spec_number.strip()
        clean_target_ver = target_version.lstrip("vV").strip()
        clean_base_ver = base_version.lstrip("vV").strip() if base_version else ""
        clean_clause = clause_number.strip()

        logger.info(
            f"🔎 [Attribution] Running clause attribution for TS {clean_spec} Cl. {clean_clause} "
            f"in v{clean_target_ver} (diff vs v{clean_base_ver or 'None'})..."
        )
        if progress_cb:
            progress_cb(10, f"Preparing attribution context for Clause {clean_clause} in v{clean_target_ver}...")

        # 1. Change Classification: NEW_CLAUSE vs. MODIFIED_CLAUSE
        if is_new_clause is None:
            if not clean_base_ver:
                is_new_clause = True
            elif self.search_db and hasattr(self.search_db, "get_connection"):
                try:
                    with self.search_db.get_connection() as conn:
                        row_base = conn.execute(
                            "SELECT 1 FROM spec_clauses WHERE spec = ? AND version = ? AND clause = ? LIMIT 1",
                            (clean_spec, clean_base_ver, clean_clause)
                        ).fetchone()
                        is_new_clause = (row_base is None)
                except Exception:
                    is_new_clause = False
            elif "@@ -0,0" in diff_text or "--- /dev/null" in diff_text:
                is_new_clause = True
            else:
                is_new_clause = False

        # 2. Ancestor Context & Key Issue Bridge
        root_title, ancestor_titles = self.resolve_ancestor_context(clean_spec, clean_target_ver, clean_clause)
        combined_title_context = " ".join([clause_title] + ancestor_titles)
        clean_title_tokens = set(self.clean_tokens(combined_title_context))

        ki_references = self.extract_key_issue_references(clause_content + " " + combined_title_context)
        salient_diff_keywords = self.extract_salient_diff_keywords(diff_text)

        # TR Solution & Key Issue Number extraction from clause hierarchy
        # e.g., if clause is 6.20 or 6.20.2, section_num = 20
        sol_num = ""
        ki_num = ""
        m_sol = re.match(r"^6\.(\d+)", clean_clause)
        if m_sol:
            sol_num = m_sol.group(1)
        m_ki = re.match(r"^5\.(\d+)", clean_clause)
        if m_ki:
            ki_num = m_ki.group(1)

        if progress_cb:
            progress_cb(25, f"Querying Change History records for v{clean_target_ver}...")

        # 3. Retrieve Candidate TDocs / CRs from 3gpp_spec_history.db
        history_rows = self.history_db.get_history_for_spec(clean_spec, version=clean_target_ver)
        if not history_rows:
            # Fallback check across all records if version string formatting differs
            all_rows = self.history_db.get_spec_history(clean_spec)
            history_rows = [r for r in all_rows if r.get("version") == clean_target_ver]

        if not history_rows:
            logger.warning(f"No history records found in 3gpp_spec_history.db for {clean_spec} v{clean_target_ver}")
            return []

        # Determine Mode A (Normative TS: individual CR rows) vs. Mode B (Study TR: omnibus p-CR list)
        is_mode_a = any(r.get("cr_number") and r.get("cr_number") != "-" for r in history_rows)

        candidate_entries: List[Dict[str, Any]] = []

        if is_mode_a:
            # Mode A: Each history row is an individual CR
            for r in history_rows:
                source_tdoc = r.get("source_tdoc", "")
                if not source_tdoc and r.get("tdocs"):
                    source_tdoc = r["tdocs"][0].get("raw_tdoc", "")
                candidate_entries.append({
                    "tdoc": source_tdoc or f"CR-{r.get('cr_number')}",
                    "cr_number": r.get("cr_number", ""),
                    "revision": r.get("revision", ""),
                    "category": r.get("category", ""),
                    "title": r.get("subject", ""),
                    "source": "",
                    "meeting": r.get("meeting", ""),
                    "agenda_item": "",
                    "ai_desc": "",
                    "abstract": "",
                    "is_pcr": False,
                    "rapporteur_notes": r.get("subject", ""),
                    "affected_clauses": r.get("affected_clauses", ""),
                })
        else:
            # Mode B: Study TR with omnibus p-CR list
            pcr_tokens: List[str] = []
            omnibus_notes = " ".join(r.get("subject", "") for r in history_rows)

            for r in history_rows:
                primary = r.get("source_tdoc", "")
                if primary and primary != "-" and primary not in pcr_tokens:
                    pcr_tokens.append(primary)
                for t in r.get("tdocs", []):
                    tok = t.get("raw_tdoc") or t.get("tdoc")
                    if tok and tok not in pcr_tokens:
                        pcr_tokens.append(tok)

            if progress_cb:
                progress_cb(40, f"Resolving metadata for {len(pcr_tokens)} candidate p-CRs...")

            # Resolve rich metadata using persistent SQLite cache / meeting agenda
            resolved_meta = self.report_gen.resolve_tdocs_metadata(
                tdoc_list=pcr_tokens,
                default_meeting=history_rows[0].get("meeting", "") if history_rows else ""
            )
            meta_by_tdoc = {m["tdoc"].upper(): m for m in resolved_meta}

            for token in pcr_tokens:
                m = meta_by_tdoc.get(token.upper()) or meta_by_tdoc.get(token.upper().split("R")[0]) or {}
                candidate_entries.append({
                    "tdoc": token,
                    "cr_number": "",
                    "revision": "",
                    "category": "",
                    "title": m.get("title", ""),
                    "source": m.get("source", ""),
                    "meeting": m.get("meeting", "") or (history_rows[0].get("meeting", "") if history_rows else ""),
                    "agenda_item": m.get("agenda_item", ""),
                    "ai_desc": m.get("ai_desc", ""),
                    "abstract": m.get("abstract", ""),
                    "remarks": m.get("remarks", ""),
                    "is_pcr": True,
                    "rapporteur_notes": omnibus_notes,
                })

        if progress_cb:
            progress_cb(60, f"Scoring and ranking {len(candidate_entries)} candidates...")

        # 4. Stage 1 Multi-Signal Scoring Matrix
        scored_candidates: List[Dict[str, Any]] = []

        for cand in candidate_entries:
            tdoc_id = cand["tdoc"]
            clean_tdoc_up = tdoc_id.upper()
            base_tdoc_up = clean_tdoc_up.split("R")[0]
            title = cand.get("title", "")
            abstract = cand.get("abstract", "")
            ai_desc = cand.get("ai_desc", "")
            agenda_item = cand.get("agenda_item", "")
            combined_meta_text = f"{title} {abstract} {ai_desc} {cand.get('remarks', '')}"
            cand_tokens = set(self.clean_tokens(combined_meta_text))

            score = 0
            evidence_badges: List[str] = []

            # --- Signal 1.1: Rapporteur Notes & Revision Lineage (Up to 35 pts) ---
            notes_text = cand.get("rapporteur_notes", "")
            if notes_text:
                # Check for explicit clause mention in vicinity of TDoc token
                pattern_clause_near_tdoc = rf"{re.escape(clean_tdoc_up)}[^\n]*?\b{re.escape(clean_clause)}\b"
                if re.search(pattern_clause_near_tdoc, notes_text, re.IGNORECASE):
                    score += 35
                    evidence_badges.append(f"🎯 Mentioned in Rapporteur Notes (Cl. {clean_clause})")
                elif sol_num and re.search(rf"{re.escape(clean_tdoc_up)}[^\n]*?\b6\.{sol_num}\b", notes_text, re.IGNORECASE):
                    score += 30
                    evidence_badges.append(f"🎯 Mentioned in Rapporteur Notes (Solution #{sol_num})")
                elif re.search(rf"\b{re.escape(clean_tdoc_up)}\b", notes_text, re.IGNORECASE):
                    # Check if notes explicitly adjust this TDoc
                    if any(kw in notes_text.lower() for kw in ["agreement", "adjusted", "implemented", "revision"]):
                        score += 15
                        evidence_badges.append("📝 Rapporteur Implementation Note")

            # --- Signal 1.2: Adaptive Structural Mapping (Up to 30 pts) ---
            aff_clauses = cand.get("affected_clauses", "")
            if aff_clauses:
                aff_list = [c.strip() for c in re.split(r"[,;\s]+", aff_clauses) if c.strip()]
                if clean_clause in aff_list or any(clean_clause == c or clean_clause.startswith(c + ".") for c in aff_list):
                    score += 30
                    evidence_badges.append(f"🎯 Listed in Affected Clauses ({clean_clause})")
            if not is_new_clause:
                # MODIFIED_CLAUSE: Explicit clause numbers are expected
                if re.search(rf"\b{re.escape(clean_clause)}\b", combined_meta_text, re.IGNORECASE):
                    score += 28
                    evidence_badges.append(f"🎯 Mentions Clause {clean_clause}")
                elif sol_num and re.search(rf"\b(?:Solution|Sol)\.?\s*#?\s*{sol_num}\b", combined_meta_text, re.IGNORECASE):
                    score += 26
                    evidence_badges.append(f"🎯 Mentions Solution #{sol_num}")
                elif ki_num and re.search(rf"\b(?:Key\s*Issue|KI)\.?\s*#?\s*{ki_num}\b", combined_meta_text, re.IGNORECASE):
                    score += 26
                    evidence_badges.append(f"🎯 Mentions Key Issue #{ki_num}")
            else:
                # NEW_CLAUSE: Placeholder numbering was used at submission; use Key Issue Bridge
                if ki_references:
                    matched_ki = False
                    for ki_ref in ki_references:
                        ki_val = ki_ref.replace("KI #", "").strip()
                        if re.search(rf"\b(?:Key\s*Issue|KI)\.?\s*#?\s*{ki_val}\b", combined_meta_text, re.IGNORECASE):
                            score += 25
                            evidence_badges.append(f"🔗 Resolves {ki_ref}")
                            matched_ki = True
                            break
                    if not matched_ki and agenda_item:
                        # Check if agenda item description matches Key Issue
                        for ki_ref in ki_references:
                            ki_val = ki_ref.replace("KI #", "").strip()
                            if ki_val in agenda_item or ki_val in ai_desc:
                                score += 20
                                evidence_badges.append(f"🔗 Agenda Item Matches {ki_ref}")
                                break

                # Boost for explicit "new solution" contributions
                if re.search(r"\b(?:new\s+solution|candidate\s+solution|add\s+solution)\b", combined_meta_text, re.IGNORECASE):
                    score += 8
                    evidence_badges.append("💡 New Solution Proposal")

            # --- Signal 1.3: Cleaned Semantic Title Overlap (Up to 25 pts) ---
            if clean_title_tokens and cand_tokens:
                intersection = clean_title_tokens.intersection(cand_tokens)
                if intersection:
                    jaccard = len(intersection) / len(clean_title_tokens.union(cand_tokens))
                    overlap_ratio = len(intersection) / len(clean_title_tokens)

                    if overlap_ratio >= 0.6 or jaccard >= 0.4:
                        score += 25
                        evidence_badges.append(f"🏷️ Title Overlap ({int(overlap_ratio * 100)}%)")
                    elif overlap_ratio >= 0.35:
                        score += 16
                        evidence_badges.append(f"🏷️ Partial Title Match ({int(overlap_ratio * 100)}%)")
                    elif len(intersection) >= 2:
                        score += 10
                        matched_words = ", ".join(list(intersection)[:2])
                        evidence_badges.append(f"🏷️ Shared Title Terms ({matched_words})")

            # --- Signal 1.4: High-Entropy Diff Keyword Salience (Up to 15 pts) ---
            if salient_diff_keywords:
                matched_diff_kws = [
                    kw for kw in salient_diff_keywords
                    if kw.lower() in combined_meta_text.lower()
                ]
                if len(matched_diff_kws) >= 3:
                    score += 15
                    evidence_badges.append(f"🔑 {len(matched_diff_kws)} Diff Keywords in Abstract")
                elif len(matched_diff_kws) >= 1:
                    score += 8
                    sample_kws = ", ".join(matched_diff_kws[:2])
                    evidence_badges.append(f"🔑 Diff Keyword ({sample_kws})")

            # Normalize composite score to [0, 100]
            final_score = min(100, score)

            # Assign human-readable confidence label
            if final_score >= 80:
                confidence_lbl = "Very High"
            elif final_score >= 55:
                confidence_lbl = "High"
            elif final_score >= 35:
                confidence_lbl = "Moderate"
            else:
                confidence_lbl = "Low"

            # Check for local document availability
            local_path = SpecTDocOpener.resolve_local_tdoc_file(tdoc_id)

            scored_candidates.append({
                "tdoc": tdoc_id,
                "cr_number": cand.get("cr_number", ""),
                "revision": cand.get("revision", ""),
                "category": cand.get("category", ""),
                "title": title or cand.get("title", ""),
                "source": cand.get("source", ""),
                "meeting": cand.get("meeting", ""),
                "agenda_item": agenda_item,
                "ai_desc": ai_desc,
                "abstract": abstract,
                "score": final_score,
                "confidence": confidence_lbl,
                "evidence_badges": evidence_badges,
                "is_new_clause": is_new_clause,
                "is_verified": False,
                "verified_match_pct": 0,
                "verified_snippet": "",
                "local_doc_path": str(local_path) if local_path and local_path.exists() else "",
            })

        # Rank descending by score
        scored_candidates.sort(key=lambda x: x["score"], reverse=True)

        # 5. Fast Automatic Verification on Top Locally-Cached Candidates
        if progress_cb:
            progress_cb(85, "Checking local candidate documents for content fingerprint match...")

        for idx, cand in enumerate(scored_candidates[:3]):
            p_str = cand.get("local_doc_path")
            if p_str and Path(p_str).exists() and diff_text:
                matched, pct, snip = self.verify_candidate_in_docx(
                    tdoc=cand["tdoc"],
                    diff_text=diff_text,
                    doc_path=Path(p_str),
                    cr_number=cand.get("cr_number", "")
                )
                if matched:
                    cand["is_verified"] = True
                    cand["verified_match_pct"] = pct
                    cand["verified_snippet"] = snip
                    cand["score"] = max(cand["score"], min(100, cand["score"] + 20))
                    cand["confidence"] = "Verified"
                    cand["evidence_badges"].insert(0, f"✅ Verified in DOCX ({pct}% Text Match)")

        # Re-sort to elevate verified candidates
        scored_candidates.sort(key=lambda x: (x["is_verified"], x["score"]), reverse=True)

        if progress_cb:
            progress_cb(100, f"Attribution complete: {len(scored_candidates)} candidates ranked.")

        top_hit = scored_candidates[0]["tdoc"] if scored_candidates else "None"
        logger.info(f"✅ [Attribution] Finished ranking {len(scored_candidates)} candidates for Cl. {clean_clause}. Top hit: {top_hit}")
        return scored_candidates

    # -------------------------------------------------------------------------
    # Stage 2: Track-Changes & Placeholder-Agnostic Content Fingerprinting
    # -------------------------------------------------------------------------

    @classmethod
    def verify_candidate_in_docx(
        cls,
        tdoc: str,
        diff_text: str,
        doc_path: Path,
        cr_number: str = ""
    ) -> Tuple[bool, int, str]:
        """
        Executes OpenXML Track-Changes-aware shingle containment check on a candidate Word document.
        Normalizes dynamic numbering ('6.x' / '6.20' -> '<NUM>') to achieve 100% ground truth.
        Returns (is_matched, match_percentage, matching_snippet).
        """
        if not doc_path.exists() or not diff_text.strip():
            return False, 0, ""

        # Extract added lines from diff
        added_lines = [
            line[1:].strip() for line in diff_text.splitlines()
            if line.startswith("+") and not line.startswith("+++") and len(line[1:].strip()) > 15
        ]
        if not added_lines:
            return False, 0, ""

        try:
            # Handle .docx (and Plenary CR pack .zip containing CR .docx)
            target_xml = ""
            if doc_path.suffix.lower() == ".docx":
                with zipfile.ZipFile(doc_path, "r") as zf:
                    if "word/document.xml" in zf.namelist():
                        target_xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
            elif doc_path.suffix.lower() == ".zip":
                with zipfile.ZipFile(doc_path, "r") as zf:
                    # If cr_number provided, find matching file e.g. *CR3104*.docx
                    candidates = [n for n in zf.namelist() if n.lower().endswith(".docx") and not n.startswith("~$")]
                    matched_file = None
                    if cr_number:
                        for c in candidates:
                            if cr_number in c:
                                matched_file = c
                                break
                    if not matched_file and candidates:
                        matched_file = candidates[0]

                    if matched_file:
                        with zipfile.ZipFile(zf.open(matched_file)) as inner_zf:
                            if "word/document.xml" in inner_zf.namelist():
                                target_xml = inner_zf.read("word/document.xml").decode("utf-8", errors="replace")

            if not target_xml:
                return False, 0, ""

            # Track-changes awareness: remove deleted text (<w:del> subtrees) when checking additions
            clean_xml = re.sub(r"<w:del\b[^>]*>.*?</w:del>", "", target_xml, flags=re.DOTALL)
            # Extract plain text runs
            text_runs = re.findall(r"<w:t\b[^>]*>(.*?)</w:t>", clean_xml, flags=re.DOTALL)
            doc_body_text = " ".join(text_runs)

            # Normalization: replace all clause/figure/table numbers with canonical <NUM>
            def _normalize(t_str: str) -> str:
                t_sub = re.sub(r"\b\d+(?:\.[0-9xX]+)+(?:-\d+)?\b", "<NUM>", t_str)
                t_sub = re.sub(r"\s+", " ", t_sub).strip().lower()
                return t_sub

            norm_doc_text = _normalize(doc_body_text)

            # 1. Exact Sentence Search (Testing 3 longest added lines)
            longest_lines = sorted(added_lines, key=len, reverse=True)[:3]
            for candidate_line in longest_lines:
                norm_line = _normalize(candidate_line)
                if len(norm_line) > 25 and norm_line in norm_doc_text:
                    snip = candidate_line[:120] + ("..." if len(candidate_line) > 120 else "")
                    return True, 98, snip

            # 2. 6-Word Shingle Containment Check
            def _make_shingles(s_str: str, k: int = 6) -> Set[Tuple[str, ...]]:
                words = s_str.split()
                if len(words) < k:
                    return {tuple(words)} if words else set()
                return {tuple(words[i:i + k]) for i in range(len(words) - k + 1)}

            sample_diff = _normalize(" ".join(added_lines[:15]))
            diff_shingles = _make_shingles(sample_diff, k=6)

            if diff_shingles:
                doc_shingles = _make_shingles(norm_doc_text, k=6)
                matches = diff_shingles.intersection(doc_shingles)
                containment = len(matches) / len(diff_shingles)
                match_pct = int(containment * 100)

                if match_pct >= 45:
                    snip = added_lines[0][:120] if added_lines else ""
                    return True, match_pct, snip

        except Exception as e:
            logger.debug(f"Error during Stage 2 fingerprinting for {tdoc}: {e}")

        return False, 0, ""


# -----------------------------------------------------------------------------
    def rank_origin_crs_for_query(
        self,
        spec_number: str,
        query: str,
        target_version: Optional[str] = None,
        affected_clauses_hint: Optional[Sequence[str]] = None,
        progress_cb: Optional[Callable[[int, str], None]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Ranks candidate Change Requests (CRs) across all releases to identify the origin CR
        that introduced the searched feature/term into the specification.

        Multi-signal scoring matrix:
          - Signal 1: Title/Subject query match & introductory keywords (up to 45 pts)
          - Signal 2: 3GPP Category weighting (Cat B = +25, Cat C = +15, Cat A = +10, Cat F = +5)
          - Signal 3: Affected clauses alignment with search results (non-exclusive, up to 30 pts)
          - Signal 4: Chronological primacy / baseline release introduction (up to 20 pts)
        """
        clean_spec = spec_number.strip()
        clean_query = query.strip()
        if not clean_query:
            return []

        logger.info(f"🎯 [Origin Ranking] Ranking origin CRs for TS {clean_spec} query '{clean_query}'...")
        if progress_cb:
            progress_cb(10, f"Querying Change History records for TS {clean_spec}...")

        all_records = self.history_db.get_spec_history(clean_spec)
        if not all_records:
            return []

        # Tokenize query
        query_lower = clean_query.lower()
        raw_query_words = re.findall(r"\b[A-Za-z0-9_-]{2,}\b", clean_query)
        salient_query_tokens = [
            w.lower() for w in raw_query_words
            if w.lower() not in BOILERPLATE_STOPWORDS and not w.isdigit()
        ]

        # Parse affected clauses hint (set of normalized clause numbers)
        target_clauses: Set[str] = set()
        if affected_clauses_hint:
            for item in affected_clauses_hint:
                for cl in re.findall(r"\b\d+(?:\.\d+)*\b", str(item)):
                    target_clauses.add(cl.strip())

        if progress_cb:
            progress_cb(30, f"Analyzing {len(all_records)} history entries...")

        # Introductory & definition keywords that indicate the origin introduction of a feature
        INTRO_KEYWORDS = {
            "definition", "define", "introduction", "intro", "introducing",
            "addition", "adding", "ki#", "key issue"
        }
        # Maintenance keywords that indicate subsequent corrections/alignments
        CORRECTION_KEYWORDS = {
            "correction", "clean-up", "cleanup", "alignment", "editorial",
            "clarification", "fixing", "fix"
        }

        # Filter candidate records that match query or target clauses
        candidates: List[Dict[str, Any]] = []
        for r in all_records:
            subj = r.get("subject", "")
            subj_lower = subj.lower()
            source_tdoc = r.get("source_tdoc", "")
            tdoc_tokens = [t.get("raw_tdoc", "") for t in r.get("tdocs", [])]
            aff = r.get("affected_clauses", "")

            # Relevance gate: CR must match query text, acronym, or target clause
            text_match = (query_lower in subj_lower) or any(tok in subj_lower for tok in salient_query_tokens if len(tok) >= 3)
            tdoc_match = (query_lower in source_tdoc.lower()) or any(query_lower in tok.lower() for tok in tdoc_tokens)
            clause_match = False
            if target_clauses and aff:
                aff_clauses = set(re.findall(r"\b\d+(?:\.\d+)*\b", aff))
                clause_match = bool(aff_clauses.intersection(target_clauses))

            if text_match or tdoc_match or clause_match:
                candidates.append(r)

        if not candidates:
            # Fallback: check all records if query has very short tokens
            candidates = all_records

        if progress_cb:
            progress_cb(60, f"Scoring {len(candidates)} candidate CRs...")

        # Determine earliest version among candidates for chronological primacy
        def ver_tuple(v_str: str) -> Tuple[int, ...]:
            parts = re.findall(r"\d+", v_str)
            return tuple(int(p) for p in parts) if parts else (999, 999, 999)

        min_version = None
        for c in candidates:
            v_str = c.get("version", "")
            if v_str:
                vt = ver_tuple(v_str)
                if min_version is None or vt < min_version:
                    min_version = vt

        scored: List[Dict[str, Any]] = []
        for r in candidates:
            subj = r.get("subject", "")
            subj_lower = subj.lower()
            cat = (r.get("category") or "").upper().strip()
            ver = r.get("version", "")
            meeting = r.get("meeting", "")
            date_str = r.get("date", "")
            cr_num = r.get("cr_number", "")
            source_tdoc = r.get("source_tdoc", "")
            aff = r.get("affected_clauses", "")

            score = 0
            evidence_badges: List[str] = []

            # --- Signal 1: Title / Subject Query Match & Origin Intent (Up to 45 pts) ---
            if re.search(rf"\b{re.escape(query_lower)}\b", subj_lower):
                score += 25
                evidence_badges.append(f"🏷️ Exact Query Match: '{clean_query}'")
            elif query_lower in subj_lower:
                score += 15
                evidence_badges.append(f"🏷️ Substring Match: '{clean_query}'")
            else:
                matched_kws = [k for k in salient_query_tokens if k in subj_lower]
                if len(matched_kws) >= 2:
                    score += 15
                    evidence_badges.append(f"🏷️ Shared Terms: {', '.join(matched_kws[:2])}")
                elif len(matched_kws) == 1:
                    score += 8
                    evidence_badges.append(f"🏷️ Term Match: {matched_kws[0]}")

            # Origin introduction keywords boost
            if any(k in subj_lower for k in INTRO_KEYWORDS):
                score += 15
                evidence_badges.append("💡 Feature Definition / Introduction")

            # Correction / Maintenance penalty (unless it's Cat B)
            if cat != "B" and any(k in subj_lower for k in CORRECTION_KEYWORDS):
                score -= 10
                evidence_badges.append("🔧 Maintenance / Correction")

            # --- Signal 2: 3GPP Category Weighting (Up to 25 pts) ---
            if cat == "B":
                score += 25
                evidence_badges.append("⭐ Category B (Addition of Feature)")
            elif cat == "C":
                score += 15
                evidence_badges.append("🔄 Category C (Functional Modification)")
            elif cat == "A":
                score += 10
                evidence_badges.append("🪞 Category A (Mirror CR)")
            elif cat in ("F", "D"):
                score += 5
                evidence_badges.append(f"🔧 Category {cat}")

            # --- Signal 3: Affected Clauses Alignment (Non-Exclusive, Up to 30 pts) ---
            if target_clauses:
                matched_aff = False
                if aff:
                    cand_clauses = set(re.findall(r"\b\d+(?:\.\d+)*\b", aff))
                    overlap = cand_clauses.intersection(target_clauses)
                    if overlap:
                        score += 25
                        evidence_badges.append(f"🎯 Affected Clauses Match: {', '.join(list(overlap)[:3])}")
                        matched_aff = True
                    else:
                        # Check prefix match (e.g. 5.15 prefix of 5.15.11)
                        if any(any(t == c or t.startswith(c + ".") or c.startswith(t + ".") for c in cand_clauses) for t in target_clauses):
                            score += 15
                            evidence_badges.append("🎯 Clause Prefix Match")
                            matched_aff = True

                # Also check if any target clause is mentioned in the subject line
                if not matched_aff:
                    for t in target_clauses:
                        if re.search(rf"\b{re.escape(t)}\b", subj):
                            score += 15
                            evidence_badges.append(f"🎯 Mentions Clause {t} in Title")
                            break

            # --- Signal 4: Chronological Primacy / Baseline Introduction (Up to 20 pts) ---
            vt = ver_tuple(ver)
            if min_version and vt == min_version:
                score += 20
                evidence_badges.append(f"⏳ Earliest Introduction Baseline (v{ver})")
            elif ver.endswith(".0.0") or ver.endswith(".1.0") or ver.startswith("0."):
                score += 15
                evidence_badges.append(f"📅 Major Release Baseline (v{ver})")
            else:
                score += 5

            # Clamp composite score to [0, 100]
            final_score = max(0, min(100, score))

            # Role Assignment
            if final_score >= 70 and (cat == "B" or any(k in subj_lower for k in INTRO_KEYWORDS)):
                role = "🎯 Foundational Origin CR"
            elif (cat in ("B", "C") or final_score >= 50) and cat not in ("F", "D") and not any(k in subj_lower for k in CORRECTION_KEYWORDS):
                role = "🏗️ Architecture Expansion"
            elif cat in ("F", "D") or any(k in subj_lower for k in CORRECTION_KEYWORDS):
                role = "🔧 Maintenance / Correction"
            else:
                role = "📝 Related CR"

            item = dict(r)
            item["score"] = final_score
            item["role"] = role
            item["evidence_badges"] = evidence_badges
            scored.append(item)

        # Sort descending by composite score, then by row order (earliest first on tie)
        scored.sort(key=lambda x: (x["score"], -x.get("row_order", 0)), reverse=True)

        if progress_cb:
            progress_cb(100, f"Ranked {len(scored)} candidate CRs.")

        return scored


# Background Worker Thread for Non-Blocking UI (SpecClauseAttributionThread)
# -----------------------------------------------------------------------------

class SpecClauseAttributionThread(QThread):
    """Background worker executing clause attribution without freezing the Qt event loop."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        spec_number: str,
        clause_number: str,
        target_version: str,
        base_version: Optional[str] = None,
        clause_title: str = "",
        diff_text: str = "",
        clause_content: str = "",
        is_new_clause: Optional[bool] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.spec_number = spec_number
        self.clause_number = clause_number
        self.target_version = target_version
        self.base_version = base_version
        self.clause_title = clause_title
        self.diff_text = diff_text
        self.clause_content = clause_content

    def run(self):
        try:
            results = self.engine.attribute_clause(
                spec_number=self.spec_number,
                clause_number=self.clause_number,
                target_version=self.target_version,
                base_version=self.base_version,
                clause_title=self.clause_title,
                diff_text=self.diff_text,
                clause_content=self.clause_content,
                is_new_clause=getattr(self, "is_new_clause", None),
                progress_cb=self.progress.emit,
            )
            self.finished.emit(results)
        except Exception as e:
            logger.error(f"Error during clause attribution thread: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecClauseDeepVerifyThread(QThread):
    """Background worker executing targeted Stage 2 content fingerprinting on top candidates."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        candidates: List[Dict[str, Any]],
        diff_text: str,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.candidates = [dict(c) for c in candidates]
        self.diff_text = diff_text

    def run(self):
        try:
            top_k = self.candidates[:3]
            total = len(top_k)
            for idx, cand in enumerate(top_k):
                if self.isInterruptionRequested():
                    return

                tdoc_id = cand.get("tdoc", "")
                self.progress.emit(
                    int((idx / max(1, total)) * 100),
                    f"Verifying candidate {idx + 1}/{total} ({tdoc_id}) in DOCX..."
                )
                local_path = SpecTDocOpener.resolve_local_tdoc_file(tdoc_id)
                if not local_path or not local_path.exists():
                    from modules.meetings.core.meetings_db import MeetingsDatabase
                    from modules.meetings.core.tdoc_url_resolver import normalize_3gpp_url, resolve_tdoc_download_url
                    from core.network.session import NetworkSession
                    from core.utils.paths import get_cache_root

                    data_db = get_project_root() / "db" / "3gpp_data.db"
                    meetings_db = MeetingsDatabase(data_db)
                    meeting = meetings_db.find_meeting_by_tdoc(tdoc_id)
                    download_url = ""
                    target_dir = get_cache_root() / "tdocs"
                    if meeting:
                        docs_url = meeting.get("docs_folder_url") or ""
                        folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "unknown")
                        target_dir = get_cache_root() / folder_name / "TDocs"
                        if docs_url:
                            norm_docs = normalize_3gpp_url(docs_url)
                            download_url = f"{norm_docs.rstrip('/')}/{tdoc_id}.zip"
                    if not download_url:
                        download_url = resolve_tdoc_download_url(tdoc_id)

                    if download_url:
                        try:
                            target_dir.mkdir(parents=True, exist_ok=True)
                            zip_dest = target_dir / f"{tdoc_id}.zip"
                            NetworkSession.download_file(
                                url=download_url,
                                dest_path=zip_dest,
                                cancel_cb=self.isInterruptionRequested,
                                atomic=True,
                            )
                            local_path = zip_dest
                        except Exception as dl_err:
                            logger.debug(f"Could not download {tdoc_id} for deep verify: {dl_err}")

                if local_path and local_path.exists():
                    matched, pct, snip = self.engine.verify_candidate_in_docx(
                        tdoc=tdoc_id,
                        diff_text=self.diff_text,
                        doc_path=local_path,
                        cr_number=cand.get("cr_number", "")
                    )
                    if matched:
                        cand["is_verified"] = True
                        cand["verified_match_pct"] = pct
                        cand["verified_snippet"] = snip
                        cand["score"] = max(cand["score"], min(100, cand["score"] + 20))
                        cand["confidence"] = "Verified"
                        badge = f"✅ Verified in DOCX ({pct}% Text Match)"
                        if badge not in cand.get("evidence_badges", []):
                            cand.setdefault("evidence_badges", []).insert(0, badge)
                        cand["local_doc_path"] = str(local_path)

            self.candidates.sort(key=lambda x: (x.get("is_verified", False), x.get("score", 0)), reverse=True)
            self.progress.emit(100, "Deep verification complete.")
            self.finished.emit(self.candidates)
        except Exception as e:
            logger.error(f"Error during deep verification: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecFeatureOriginRankingThread(QThread):
    """Background worker executing feature-level Origin CR ranking across specification releases."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        spec_number: str,
        query: str,
        target_version: Optional[str] = None,
        affected_clauses_hint: Optional[Sequence[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.spec_number = spec_number
        self.query = query
        self.target_version = target_version
        self.affected_clauses_hint = list(affected_clauses_hint) if affected_clauses_hint else None

    def run(self):
        try:
            results = self.engine.rank_origin_crs_for_query(
                spec_number=self.spec_number,
                query=self.query,
                target_version=self.target_version,
                affected_clauses_hint=self.affected_clauses_hint,
                progress_cb=self.progress.emit,
            )
            self.finished.emit(results)
        except Exception as e:
            logger.error(f"Error during feature origin ranking thread: {e}", exc_info=True)
            self.error.emit(str(e))
