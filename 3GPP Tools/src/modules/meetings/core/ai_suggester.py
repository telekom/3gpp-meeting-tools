# --- File: src/modules/meetings/core/ai_suggester.py ---
"""
Cross-WG Agenda Item (AI) Discovery and Suggestion Engine.

Scans cached meeting workbooks and database records across 3GPP Working Groups
(SA, RAN, CT) to automatically discover and rank relevant Agenda Items based on
topic keywords, Work Item tokens, and title correlations.
"""

import json
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser

logger = logging.getLogger(__name__)


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

    def suggest_ais_for_topic(
        self,
        target_wgs: List[str],
        keywords: List[str],
        wi_tokens: Optional[List[str]] = None,
        max_meetings_per_wg: int = 6,
    ) -> Dict[str, List[DiscoveredAi]]:
        """
        Scans recent meetings for the target WGs and returns ranked AI recommendations.

        :param target_wgs: List of Working Groups (e.g. ['SA2', 'RAN2', 'RAN3', 'SA5', 'CT4'])
        :param keywords: Search phrases (e.g. ['data framework', 'data transfer'])
        :param wi_tokens: Optional Work Item tokens (e.g. ['fs_6g_arc', '6g_df'])
        :param max_meetings_per_wg: Max recent meetings to analyze per WG
        :return: Dict mapping WG -> List of DiscoveredAi sorted by relevance
        """
        cache_dir = self._get_cache_dir()
        results: Dict[str, List[DiscoveredAi]] = {}

        clean_kws = [k.strip().lower() for k in (keywords or []) if k.strip()]
        clean_wis = [w.strip().lower() for w in (wi_tokens or []) if w.strip()]

        if not clean_kws and not clean_wis:
            return {wg: [] for wg in target_wgs}

        for wg in target_wgs:
            norm_wg = wg.strip().upper()
            results[norm_wg] = []

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

                # If keywords were provided, require keyword relevance (in desc or titles)
                if clean_kws and not desc_match and title_cnt == 0:
                    continue

                # If no keywords were provided (only WI tokens), require WI match
                if not clean_kws and wi_cnt == 0:
                    continue

                reasons = []
                if desc_match:
                    reasons.append("AI description matches topic keywords")
                if title_cnt > 0:
                    reasons.append(f"TDoc title match ({title_cnt} TDocs)")
                if wi_cnt > 0:
                    reasons.append(f"Work Item match ({wi_cnt} TDocs)")

                # Confidence calculation
                if desc_match:
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
                    ai_description=d["desc"],
                    confidence=confidence,
                    matching_reason="; ".join(reasons),
                    tdoc_count=tot_cnt,
                    matching_tdocs=title_cnt if clean_kws else wi_cnt,
                    meetings=sorted(list(d["meetings"])),
                    sample_titles=d["titles"][:3],
                    sample_wis=sorted(list(d["wis"]))[:3],
                ))

            # 4. Sort: High confidence first, then description match, then title matches, then total count
            results[norm_wg].sort(
                key=lambda x: (
                    0 if x.confidence == "High" else 1,
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
            # Filter valid start dates and sort reverse
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

        # Fallback: scan cache_dir directly for matching folders
        candidates = []
        patterns = [
            f"*{wg}*",
            f"*TSG{wg}*",
            f"*TSG{wg[0]}_{wg[1:]}*",
            f"*TSG{wg[0]}{wg[-1]}*",
            f"*{wg[0]}{wg[-1]}_*",
        ]
        found_dirs = set()
        for p in patterns:
            for d in cache_dir.glob(p):
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
                        return json.load(f)
                except Exception:
                    pass
            for ef in agenda_dir.glob("*.xlsx"):
                if not ef.name.startswith("~$"):
                    try:
                        return TDocsParser.parse_tdocs_excel(str(ef))
                    except Exception:
                        pass

        if mtg_dir.is_dir():
            for ef in mtg_dir.glob("*.xlsx"):
                if not ef.name.startswith("~$"):
                    try:
                        return TDocsParser.parse_tdocs_excel(str(ef))
                    except Exception:
                        pass

        return []
