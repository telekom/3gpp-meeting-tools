# --- File: src/modules/specifications/core/attribution/origin_ranker.py ---
"""Cross-release Change Request (CR) origin ranking for specification features/terms."""

import logging
import re
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

from .text_helpers import BOILERPLATE_STOPWORDS

logger = logging.getLogger(__name__)


class OriginRankerMixin:
    """Mixin ranking Change Requests across releases to identify origin CRs."""

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

        query_lower = clean_query.lower()
        raw_query_words = re.findall(r"\b[A-Za-z0-9_-]{2,}\b", clean_query)
        salient_query_tokens = [
            w.lower()
            for w in raw_query_words
            if w.lower() not in BOILERPLATE_STOPWORDS and not w.isdigit()
        ]

        target_clauses: Set[str] = set()
        if affected_clauses_hint:
            for item in affected_clauses_hint:
                for cl in re.findall(r"\b\d+(?:\.\d+)*\b", str(item)):
                    target_clauses.add(cl.strip())

        if progress_cb:
            progress_cb(30, f"Analyzing {len(all_records)} history entries...")

        INTRO_KEYWORDS = {
            "definition",
            "define",
            "introduction",
            "intro",
            "introducing",
            "addition",
            "adding",
            "ki#",
            "key issue",
        }
        CORRECTION_KEYWORDS = {
            "correction",
            "clean-up",
            "cleanup",
            "alignment",
            "editorial",
            "clarification",
            "fixing",
            "fix",
        }

        candidates: List[Dict[str, Any]] = []
        for r in all_records:
            subj = r.get("subject", "")
            subj_lower = subj.lower()
            source_tdoc = r.get("source_tdoc", "")
            tdoc_tokens = [t.get("raw_tdoc", "") for t in r.get("tdocs", [])]
            aff = r.get("affected_clauses", "")

            text_match = (query_lower in subj_lower) or any(
                tok in subj_lower for tok in salient_query_tokens if len(tok) >= 3
            )
            tdoc_match = (query_lower in source_tdoc.lower()) or any(
                query_lower in tok.lower() for tok in tdoc_tokens
            )
            clause_match = False
            if target_clauses and aff:
                aff_clauses = set(re.findall(r"\b\d+(?:\.\d+)*\b", aff))
                clause_match = bool(aff_clauses.intersection(target_clauses))

            if text_match or tdoc_match or clause_match:
                candidates.append(r)

        if not candidates:
            candidates = all_records

        if progress_cb:
            progress_cb(60, f"Scoring {len(candidates)} candidate CRs...")

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

            if any(k in subj_lower for k in INTRO_KEYWORDS):
                score += 15
                evidence_badges.append("💡 Feature Definition / Introduction")

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
                    elif any(
                        any(t == c or t.startswith(c + ".") or c.startswith(t + ".") for c in cand_clauses)
                        for t in target_clauses
                    ):
                        score += 15
                        evidence_badges.append("🎯 Clause Prefix Match")
                        matched_aff = True

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

            final_score = max(0, min(100, score))

            if final_score >= 70 and (cat == "B" or any(k in subj_lower for k in INTRO_KEYWORDS)):
                role = "🎯 Foundational Origin CR"
            elif (
                (cat in ("B", "C") or final_score >= 50)
                and cat not in ("F", "D")
                and not any(k in subj_lower for k in CORRECTION_KEYWORDS)
            ):
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

        scored.sort(key=lambda x: (x["score"], -x.get("row_order", 0)), reverse=True)

        if progress_cb:
            progress_cb(100, f"Ranked {len(scored)} candidate CRs.")

        return scored
