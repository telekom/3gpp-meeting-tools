# --- File: src/modules/specifications/core/attribution/clause_matcher.py ---
"""Stage 1 multi-signal metadata and structural ranking for clause attribution."""

import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener

logger = logging.getLogger(__name__)


class ClauseMatcherMixin:
    """Mixin implementing Stage 1 clause-to-CR/p-CR ranking."""

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
                            (clean_spec, clean_base_ver, clean_clause),
                        ).fetchone()
                        is_new_clause = row_base is None
                except Exception:
                    is_new_clause = False
            elif "@@ -0,0" in diff_text or "--- /dev/null" in diff_text:
                is_new_clause = True
            else:
                is_new_clause = False

        # 2. Ancestor Context & Key Issue Bridge
        _root_title, ancestor_titles = self.resolve_ancestor_context(clean_spec, clean_target_ver, clean_clause)
        combined_title_context = " ".join([clause_title] + ancestor_titles)
        clean_title_tokens = set(self.clean_tokens(combined_title_context))

        ki_references = self.extract_key_issue_references(clause_content + " " + combined_title_context)
        salient_diff_keywords = self.extract_salient_diff_keywords(diff_text)

        # TR Solution & Key Issue Number extraction from clause hierarchy
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
            all_rows = self.history_db.get_spec_history(clean_spec)
            history_rows = [r for r in all_rows if r.get("version") == clean_target_ver]

        if not history_rows:
            logger.warning(f"No history records found in 3gpp_spec_history.db for {clean_spec} v{clean_target_ver}")
            return []

        # Determine Mode A (Normative TS: individual CR rows) vs. Mode B (Study TR: omnibus p-CR list)
        is_mode_a = any(r.get("cr_number") and r.get("cr_number") != "-" for r in history_rows)

        candidate_entries: List[Dict[str, Any]] = []

        if is_mode_a:
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

            resolved_meta = self.report_gen.resolve_tdocs_metadata(
                tdoc_list=pcr_tokens,
                default_meeting=history_rows[0].get("meeting", "") if history_rows else "",
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
                pattern_clause_near_tdoc = rf"{re.escape(clean_tdoc_up)}[^\n]*?\b{re.escape(clean_clause)}\b"
                if re.search(pattern_clause_near_tdoc, notes_text, re.IGNORECASE):
                    score += 35
                    evidence_badges.append(f"🎯 Mentioned in Rapporteur Notes (Cl. {clean_clause})")
                elif sol_num and re.search(
                    rf"{re.escape(clean_tdoc_up)}[^\n]*?\b6\.{sol_num}\b", notes_text, re.IGNORECASE
                ):
                    score += 30
                    evidence_badges.append(f"🎯 Mentioned in Rapporteur Notes (Solution #{sol_num})")
                elif re.search(rf"\b{re.escape(clean_tdoc_up)}\b", notes_text, re.IGNORECASE):
                    if any(kw in notes_text.lower() for kw in ["agreement", "adjusted", "implemented", "revision"]):
                        score += 15
                        evidence_badges.append("📝 Rapporteur Implementation Note")

            # --- Signal 1.2: Adaptive Structural Mapping (Up to 30 pts) ---
            aff_clauses = cand.get("affected_clauses", "")
            if aff_clauses:
                aff_list = [c.strip() for c in re.split(r"[,;\s]+", aff_clauses) if c.strip()]
                if clean_clause in aff_list or any(
                    clean_clause == c or clean_clause.startswith(c + ".") for c in aff_list
                ):
                    score += 30
                    evidence_badges.append(f"🎯 Listed in Affected Clauses ({clean_clause})")
            if not is_new_clause:
                if re.search(rf"\b{re.escape(clean_clause)}\b", combined_meta_text, re.IGNORECASE):
                    score += 28
                    evidence_badges.append(f"🎯 Mentions Clause {clean_clause}")
                elif sol_num and re.search(
                    rf"\b(?:Solution|Sol)\.?\s*#?\s*{sol_num}\b", combined_meta_text, re.IGNORECASE
                ):
                    score += 26
                    evidence_badges.append(f"🎯 Mentions Solution #{sol_num}")
                elif ki_num and re.search(
                    rf"\b(?:Key\s*Issue|KI)\.?\s*#?\s*{ki_num}\b", combined_meta_text, re.IGNORECASE
                ):
                    score += 26
                    evidence_badges.append(f"🎯 Mentions Key Issue #{ki_num}")
            else:
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
                        for ki_ref in ki_references:
                            ki_val = ki_ref.replace("KI #", "").strip()
                            if ki_val in agenda_item or ki_val in ai_desc:
                                score += 20
                                evidence_badges.append(f"🔗 Agenda Item Matches {ki_ref}")
                                break

                if re.search(
                    r"\b(?:new\s+solution|candidate\s+solution|add\s+solution)\b", combined_meta_text, re.IGNORECASE
                ):
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

            final_score = min(100, score)

            if final_score >= 80:
                confidence_lbl = "Very High"
            elif final_score >= 55:
                confidence_lbl = "High"
            elif final_score >= 35:
                confidence_lbl = "Moderate"
            else:
                confidence_lbl = "Low"

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

        scored_candidates.sort(key=lambda x: x["score"], reverse=True)

        # 5. Fast Automatic Verification on Top Locally-Cached Candidates
        if progress_cb:
            progress_cb(85, "Checking local candidate documents for content fingerprint match...")

        for cand in scored_candidates[:3]:
            p_str = cand.get("local_doc_path")
            if p_str and Path(p_str).exists() and diff_text:
                matched, pct, snip = self.verify_candidate_in_docx(
                    tdoc=cand["tdoc"],
                    diff_text=diff_text,
                    doc_path=Path(p_str),
                    cr_number=cand.get("cr_number", ""),
                )
                if matched:
                    cand["is_verified"] = True
                    cand["verified_match_pct"] = pct
                    cand["verified_snippet"] = snip
                    cand["score"] = max(cand["score"], min(100, cand["score"] + 20))
                    cand["confidence"] = "Verified"
                    cand["evidence_badges"].insert(0, f"✅ Verified in DOCX ({pct}% Text Match)")

        scored_candidates.sort(key=lambda x: (x["is_verified"], x["score"]), reverse=True)

        if progress_cb:
            progress_cb(100, f"Attribution complete: {len(scored_candidates)} candidates ranked.")

        top_hit = scored_candidates[0]["tdoc"] if scored_candidates else "None"
        logger.info(
            f"✅ [Attribution] Finished ranking {len(scored_candidates)} candidates for Cl. {clean_clause}. Top hit: {top_hit}"
        )
        return scored_candidates
