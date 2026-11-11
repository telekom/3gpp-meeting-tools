# --- File: src/modules/meetings/core/ai_suggester.py ---
"""
Cross-WG Agenda Item (AI) Discovery and Suggestion Engine.

Scans cached meeting workbooks and database records across 3GPP Working Groups
(SA, RAN, CT) to automatically discover and rank relevant Agenda Items based on
topic keywords, Work Item tokens, and title correlations.
"""

from collections import Counter
from dataclasses import dataclass, field
import json
import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser

logger = logging.getLogger(__name__)

# Common administrative/generic Agenda Items that should never be auto-classified as "High" confidence
GENERIC_AI_CATEGORIES = {
    "corrections", "correction", "drafting", "drafting session", "drafting sessions",
    "miscellaneous", "approval of the agenda", "approval of agenda", "review of contributions",
    "incoming ls", "outgoing ls", "liaison statements", "liaison statement",
    "liaisons", "administrative", "work item management", "any other business",
    "aob", "opening of the meeting", "wrap-up", "closure of the meeting",
    "status report", "session report", "allocation of documents", "meeting schedule",
    "minutes", "approval of minutes", "review of actions",
}

# Working Group folder prefix patterns in 3GPP Delegate Helper cache
WG_FOLDER_PREFIXES = {
    "SA1": ["TSGS1", "TSGS1_", "SA1"],
    "SA2": ["TSGS2", "TSGS2_", "SA2"],
    "SA3": ["TSGS3", "TSGS3_", "SA3"],
    "SA4": ["TSGS4", "TSGS4_", "SA4"],
    "SA5": ["TSGS5", "TSGS5_", "SA5"],
    "SA6": ["TSGS6", "TSGS6_", "SA6"],
    "RAN1": ["TSGR1", "TSGR1_", "RAN1"],
    "RAN2": ["TSGR2", "TSGR2_", "RAN2"],
    "RAN3": ["TSGR3", "TSGR3_", "RAN3"],
    "RAN4": ["TSGR4", "TSGR4_", "RAN4"],
    "CT1": ["TSGC1", "TSGC1_", "CT1"],
    "CT3": ["TSGC3", "TSGC3_", "CT3"],
    "CT4": ["CT4", "CT4_", "TSGC4", "TSGC4_"],
    "CT6": ["TSGC6", "TSGC6_", "CT6"],
    "SA": ["TSG_SA", "TSGSA", "SP_"],
    "RAN": ["TSG_RAN", "TSGRAN", "RP_"],
    "CT": ["TSG_CT", "TSGCT", "CP_"],
}


@dataclass
class DiscoveredAi:
    """Represents a candidate Agenda Item discovered from meeting records."""
    wg: str
    ai_number: str
    ai_description: str
    confidence: str  # "High", "Medium", "Low"
    matching_reason: str
    tdoc_count: int
    matching_tdocs: int
    meetings: List[str] = field(default_factory=list)
    sample_titles: List[str] = field(default_factory=list)
    sample_wis: List[str] = field(default_factory=list)

    @property
    def suggested_pattern(self) -> str:
        """Returns pattern formatted with wildcard if numeric."""
        clean = self.ai_number.strip()
        if not clean.endswith("*") and not clean.endswith("?"):
            return f"{clean}*"
        return clean


class AgendaItemSuggester:
    """Scans and analyzes meeting records to recommend Agenda Items per Working Group."""

    def __init__(self, meetings_db: MeetingsDatabase):
        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()

    def _get_cache_dir(self) -> Path:
        cache_str = self.settings.cache_dir
        if cache_str and Path(cache_str).is_dir():
            return Path(cache_str)
        from core.utils.paths import get_cache_root
        return Path(get_cache_root())

    def infer_topic_from_known_ais(
        self,
        known_ai_patterns: Dict[str, List[str]],
        max_meetings: int = 6,
    ) -> Dict[str, Any]:
        """
        Inspects known Working Group Agenda Items to discover topic descriptions,
        salient keywords, and Work Item tokens.

        Example: {'SA2': ['20.6.21']} discovers:
        - Description: "Study on Architecture for 6G System (FS_6G_ARC) - 6G data framework in SA WG2"
        - Keywords: ['data framework', '6G data framework', 'data transfer']
        - WI Tokens: ['fs_6g_arc']
        """
        cache_dir = self._get_cache_dir()
        discovered_kws: List[str] = []
        discovered_wis: List[str] = []
        descriptions: List[str] = []
        collected_titles: List[str] = []
        source_summaries: List[str] = []

        for wg, pats in known_ai_patterns.items():
            norm_wg = wg.strip().upper()
            recent_mtgs = self._resolve_recent_meetings(norm_wg, cache_dir, max_meetings)
            for pat in pats:
                clean_pat = re.sub(r'[*?]', '', pat).strip().lower()
                matching_tdocs = []
                for m in recent_mtgs:
                    f_name = m.get("folder_name", "")
                    tdocs = self._load_meeting_tdocs(cache_dir / f_name)
                    for t in tdocs:
                        ai = str(t.get("Agenda Item", "") or t.get("AI", "")).strip().lower()
                        if ai == clean_pat or ai.startswith(clean_pat + "."):
                            matching_tdocs.append(t)
                    if len(matching_tdocs) >= 150:
                        break

                if not matching_tdocs:
                    continue

                for t in matching_tdocs:
                    desc = str(t.get("Agenda item description", "") or t.get("AI Description", "")).strip()
                    if desc and desc not in descriptions:
                        descriptions.append(desc)
                    wi = str(t.get("Related WIs", "") or t.get("wi", "")).strip()
                    if wi:
                        for w in re.split(r'[,; ]+', wi):
                            w_clean = re.sub(r'[^a-zA-Z0-9_]', '', w).lower()
                            if len(w_clean) >= 3 and w_clean not in discovered_wis:
                                discovered_wis.append(w_clean)
                    title = str(t.get("Title", "") or "").strip()
                    if title:
                        collected_titles.append(title)

                top_desc = descriptions[0] if descriptions else f"AI {pat}"
                source_summaries.append(f"{norm_wg} {pat} ('{top_desc}')")

        # Extract subtopics and WI tokens from descriptions
        for desc in descriptions:
            # WI from parentheses, e.g. "(FS_6G_ARC)"
            for pw in re.findall(r'\(([A-Za-z0-9_]+)\)', desc):
                pw_clean = pw.lower()
                if pw_clean not in discovered_wis:
                    discovered_wis.append(pw_clean)

            # Subtopic phrase after hyphen, e.g. " - 6G data framework in SA2"
            if " - " in desc:
                after_dash = desc.split(" - ", 1)[1].strip()
                clean_sub = re.sub(r'\s+in\s+([A-Z]{2,4}\s*(?:WG)?\d*)$', '', after_dash, flags=re.IGNORECASE).strip()
                if clean_sub and clean_sub.lower() not in [k.lower() for k in discovered_kws]:
                    discovered_kws.append(clean_sub)
                    if clean_sub.lower().startswith("6g "):
                        base_kw = clean_sub[3:].strip()
                        if base_kw and base_kw.lower() not in [k.lower() for k in discovered_kws]:
                            discovered_kws.append(base_kw)

        # Extract recurring topic-relevant phrases from collected titles
        # Only consider phrases that share semantic root tokens with the discovered subtopics
        subtopic_tokens = set()
        for kw in discovered_kws:
            for tok in re.findall(r'[a-zA-Z]{3,}', kw.lower()):
                subtopic_tokens.add(tok)

        title_text = " ".join(collected_titles).lower()
        candidate_phrases = [
            "data framework", "data transfer", "model transfer", "data collection",
            "ambient iot", "zero energy", "passive iot", "network energy savings",
            "integrated sensing", "isac", "slicing", "satellite", "non-terrestrial",
        ]
        for cand in candidate_phrases:
            cand_tokens = set(re.findall(r'[a-zA-Z]{3,}', cand.lower()))
            # Only add if candidate shares at least one significant token with discovered subtopic
            if subtopic_tokens and cand_tokens.intersection(subtopic_tokens):
                if cand in title_text and cand not in [k.lower() for k in discovered_kws]:
                    discovered_kws.append(cand)

        return {
            "keywords": discovered_kws,
            "wi_tokens": discovered_wis,
            "descriptions": descriptions,
            "source_summary": "; ".join(source_summaries),
        }

    def suggest_ais_for_topic(
        self,
        target_wgs: List[str],
        keywords: List[str],
        wi_tokens: Optional[List[str]] = None,
        max_meetings_per_wg: int = 6,
        progress_callback: Optional[Callable[[str, int, int], None]] = None,
    ) -> Dict[str, List[DiscoveredAi]]:
        """
        Scans recent meetings for the target WGs and returns ranked AI recommendations.

        :param target_wgs: List of Working Groups (e.g. ['SA2', 'RAN2', 'RAN3', 'SA5', 'CT4'])
        :param keywords: Search phrases (e.g. ['data framework', 'data transfer'])
        :param wi_tokens: Optional Work Item tokens (e.g. ['fs_6g_arc', '6g_df'])
        :param max_meetings_per_wg: Max recent meetings to analyze per WG
        :param progress_callback: Optional callback(current_wg_name, current_idx, total_wgs)
        :return: Dict mapping WG -> List of DiscoveredAi sorted by relevance
        """
        cache_dir = self._get_cache_dir()
        results: Dict[str, List[DiscoveredAi]] = {}

        clean_kws = [k.strip().lower() for k in (keywords or []) if k.strip()]
        clean_wis = [w.strip().lower() for w in (wi_tokens or []) if w.strip()]

        if not clean_kws and not clean_wis:
            return {wg: [] for wg in target_wgs}

        total_wgs = len(target_wgs)

        for w_idx, wg in enumerate(target_wgs, start=1):
            norm_wg = wg.strip().upper()
            results[norm_wg] = []

            if progress_callback:
                progress_callback(norm_wg, w_idx, total_wgs)

            # 1. Resolve candidate meetings for this WG
            candidate_mtgs = self._resolve_recent_meetings(norm_wg, cache_dir, max_meetings_per_wg)
            if not candidate_mtgs:
                continue

            # 2. Extract and aggregate distinct Agenda Items
            ai_data: Dict[str, Dict[str, Any]] = {}

            for m in candidate_mtgs:
                f_name = m.get("folder_name", "")
                mtg_dir = cache_dir / f_name
                tdocs = self._load_meeting_tdocs(mtg_dir)
                mtg_display = m.get("meeting_number") or f_name

                for row in tdocs:
                    ai = str(row.get("Agenda Item", "") or row.get("Agenda item", "") or row.get("AI", "")).strip()
                    if not ai:
                        continue
                    ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "")).strip()
                    title = str(row.get("Title", "") or "").strip()
                    wi = str(row.get("Related WIs", "") or row.get("wi", "")).strip()

                    if ai not in ai_data:
                        ai_data[ai] = {
                            "ai": ai,
                            "desc": ai_desc,
                            "tdoc_count": 0,
                            "match_count": 0,
                            "meetings": set(),
                            "titles": [],
                            "wis": set(),
                            "desc_match": False,
                            "title_matches": 0,
                            "wi_matches": 0,
                        }

                    entry = ai_data[ai]
                    entry["tdoc_count"] += 1
                    entry["meetings"].add(mtg_display)
                    if not entry["desc"] and ai_desc:
                        entry["desc"] = ai_desc
                    if wi:
                        entry["wis"].add(wi)

                    # Check AI description match
                    desc_lower = f"{ai_desc} {ai}".lower()
                    if any(kw in desc_lower for kw in clean_kws):
                        entry["desc_match"] = True

                    # Check Title match
                    title_lower = title.lower()
                    if any(kw in title_lower for kw in clean_kws):
                        entry["title_matches"] += 1
                        entry["match_count"] += 1
                        if len(entry["titles"]) < 4:
                            entry["titles"].append(title)

                    # Check WI match
                    wi_lower = wi.lower()
                    if any(wt in wi_lower for wt in clean_wis):
                        entry["wi_matches"] += 1
                        entry["match_count"] += 1

            # 3. Evaluate confidence and filter
            for ai, d in ai_data.items():
                desc_match = d["desc_match"]
                tot_cnt = d["tdoc_count"]
                title_cnt = d["title_matches"]
                wi_cnt = d["wi_matches"]
                ai_desc_str = d["desc"]

                # If keywords were provided, require keyword relevance (in desc or titles)
                if clean_kws and not desc_match and title_cnt == 0:
                    continue

                # If no keywords were provided (only WI tokens), require WI match
                if not clean_kws and wi_cnt == 0:
                    continue

                desc_norm = ai_desc_str.strip().lower()
                is_generic_ai = any(gen in desc_norm for gen in GENERIC_AI_CATEGORIES)

                reasons = []
                if desc_match:
                    reasons.append("AI description matches topic keywords")
                if title_cnt > 0:
                    reasons.append(f"TDoc title match ({title_cnt} TDocs)")
                if wi_cnt > 0:
                    reasons.append(f"Work Item match ({wi_cnt} TDocs)")
                if is_generic_ai:
                    reasons.append("Caution: Generic/administrative AI category")

                # Confidence calculation
                if is_generic_ai:
                    confidence = "Medium" if title_cnt >= 2 else "Low"
                elif desc_match:
                    confidence = "High"
                elif title_cnt >= 3 and (wi_cnt > 0 or not clean_wis):
                    confidence = "High"
                elif title_cnt >= 8:
                    confidence = "High"
                else:
                    confidence = "Medium"

                results[norm_wg].append(DiscoveredAi(
                    wg=norm_wg,
                    ai_number=ai,
                    ai_description=ai_desc_str,
                    confidence=confidence,
                    matching_reason="; ".join(reasons),
                    tdoc_count=tot_cnt,
                    matching_tdocs=title_cnt if clean_kws else wi_cnt,
                    meetings=sorted(list(d["meetings"])),
                    sample_titles=d["titles"][:3],
                    sample_wis=sorted(list(d["wis"]))[:3],
                ))

            # 4. Sort: High confidence first, non-generic first, then match count, then total count
            results[norm_wg].sort(
                key=lambda x: (
                    0 if x.confidence == "High" else 1,
                    1 if "Generic/administrative" in x.matching_reason else 0,
                    0 if "AI description matches" in x.matching_reason else 1,
                    -x.matching_tdocs,
                    -x.tdoc_count,
                )
            )

        return results

    def _resolve_recent_meetings(self, wg: str, cache_dir: Path, max_count: int) -> List[Dict[str, Any]]:
        """Resolves existing meeting directories for this WG in chronological reverse order."""
        # Query database first
        try:
            mtgs = self.meetings_db.search_meetings(wg_name=[wg])
            mtgs_sorted = sorted(
                [m for m in mtgs if m.get("start_date")],
                key=lambda x: str(x.get("start_date")),
                reverse=True,
            )
            candidate_mtgs = []
            for m in mtgs_sorted:
                f_name = m.get("folder_name", "")
                if (cache_dir / f_name).is_dir():
                    candidate_mtgs.append(m)
                if len(candidate_mtgs) >= max_count:
                    return candidate_mtgs
            if candidate_mtgs:
                return candidate_mtgs
        except Exception as e:
            logger.debug(f"[AiSuggester] DB search failed for {wg}: {e}")

        # Fallback: scan cache_dir directly with standard WG folder prefixes
        candidates = []
        prefixes = WG_FOLDER_PREFIXES.get(wg.upper(), [wg, f"TSG{wg}"])
        found_dirs = set()
        for pfx in prefixes:
            for d in cache_dir.glob(f"*{pfx}*"):
                if d.is_dir():
                    found_dirs.add(d)

        sorted_dirs = sorted(list(found_dirs), key=lambda d: d.name, reverse=True)
        for d in sorted_dirs[:max_count]:
            candidates.append({
                "folder_name": d.name,
                "meeting_number": d.name,
                "wg_name": wg,
            })
        return candidates

    def _load_meeting_tdocs(self, mtg_dir: Path) -> List[Dict[str, Any]]:
        """Loads TDoc list from pre-parsed JSON cache or Excel workbook."""
        agenda_dir = mtg_dir / "Agenda"
        if agenda_dir.is_dir():
            for jf in agenda_dir.glob("*.xlsx.json"):
                try:
                    with open(jf, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if data:
                            return data
                except Exception:
                    pass
            for ef in agenda_dir.glob("*.xlsx"):
                if not ef.name.startswith("~$"):
                    try:
                        data = TDocsParser.parse_tdocs_excel(str(ef))
                        if data:
                            return data
                    except Exception:
                        pass

        if mtg_dir.is_dir():
            for ef in mtg_dir.glob("*.xlsx"):
                if not ef.name.startswith("~$"):
                    try:
                        data = TDocsParser.parse_tdocs_excel(str(ef))
                        if data:
                            return data
                    except Exception:
                        pass

        return []
