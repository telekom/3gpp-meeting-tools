# --- File: src/modules/meetings/core/cohort_resolver.py ---
"""
Meeting Cohort Resolver.

Clusters multi-WG 3GPP meetings (SA, RAN, CT) across calendar dates into
synchronized quarterly Meeting Cohorts (e.g. 2025-Q4 Cycle), resolving staggered
schedules into unified cross-group milestones.
"""

import datetime
import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.topic_manager import MetaTopicDefinition

logger = logging.getLogger(__name__)

_DATE_PARSE_RE = re.compile(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')


def get_cohort_key(date_str: str) -> Optional[str]:
    """
    Computes standard quarterly cohort key (e.g., '2025-Q4') from an ISO or standard date string.
    """
    if not date_str:
        return None
    m = _DATE_PARSE_RE.search(str(date_str))
    if not m:
        return None
    year = int(m.group(1))
    month = int(m.group(2))
    if month < 1 or month > 12:
        return None
    quarter = ((month - 1) // 3) + 1
    return f"{year}-Q{quarter}"


def parse_cohort_key(key: str) -> Tuple[int, int]:
    """Parses 'YYYY-QN' into (year, quarter)."""
    m = re.match(r'^(\d{4})-Q([1-4])$', str(key or "").strip().upper())
    if not m:
        return (0, 0)
    return (int(m.group(1)), int(m.group(2)))


def compare_cohort_keys(k1: str, k2: str) -> int:
    """Returns -1 if k1 < k2, 0 if equal, 1 if k1 > k2."""
    y1, q1 = parse_cohort_key(k1)
    y2, q2 = parse_cohort_key(k2)
    if (y1, q1) < (y2, q2):
        return -1
    elif (y1, q1) > (y2, q2):
        return 1
    return 0


def format_cohort_label(key: str) -> str:
    """Returns a friendly label, e.g. '2025 Q4 Meeting Cycle'."""
    y, q = parse_cohort_key(key)
    if y == 0:
        return key
    return f"{y} Q{q} Cycle"


@dataclass
class MeetingCohort:
    key: str
    year: int
    quarter: int
    label: str
    meetings: List[Dict[str, Any]] = field(default_factory=list)
    wgs_present: Set[str] = field(default_factory=set)

    @property
    def start_date(self) -> str:
        dates = [m.get("start_date") for m in self.meetings if m.get("start_date")]
        return min(dates) if dates else ""

    @property
    def end_date(self) -> str:
        dates = [m.get("end_date") for m in self.meetings if m.get("end_date")]
        return max(dates) if dates else ""

    def get_meetings_for_wg(self, wg: str) -> List[Dict[str, Any]]:
        target = str(wg or "").strip().upper()
        return [m for m in self.meetings if str(m.get("wg_name", "")).strip().upper() == target]


class MeetingCohortResolver:
    """Resolves and clusters meetings across Working Groups into quarterly cycles."""

    @staticmethod
    def cluster_meetings_into_cohorts(meetings: List[Dict[str, Any]]) -> List[MeetingCohort]:
        """
        Groups arbitrary meeting records into chronologically ordered MeetingCohorts.
        """
        cohort_map: Dict[str, List[Dict[str, Any]]] = {}

        for m in meetings:
            d_str = m.get("start_date") or m.get("end_date") or ""
            key = get_cohort_key(d_str)
            if not key:
                # Fallback: try parsing folder_name if date missing
                key = "Unknown"
            if key not in cohort_map:
                cohort_map[key] = []
            cohort_map[key].append(m)

        # Build list of cohorts sorted chronologically
        cohorts: List[MeetingCohort] = []
        for key, mtgs in cohort_map.items():
            if key == "Unknown":
                continue
            y, q = parse_cohort_key(key)
            wgs = {str(m.get("wg_name", "")).strip().upper() for m in mtgs if m.get("wg_name")}
            # Sort internal meetings by WG name and start date
            sorted_mtgs = sorted(mtgs, key=lambda x: (str(x.get("wg_name", "")), str(x.get("start_date", ""))))
            cohorts.append(MeetingCohort(
                key=key,
                year=y,
                quarter=q,
                label=format_cohort_label(key),
                meetings=sorted_mtgs,
                wgs_present=wgs,
            ))

        cohorts.sort(key=lambda c: (c.year, c.quarter))
        return cohorts

    @classmethod
    def resolve_cohorts_for_topic(
        cls,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
    ) -> List[MeetingCohort]:
        """
        Queries meetings_db for target WGs and resolves the cohorts according to the topic's cycle mode.
        """
        target_wgs = [w.strip() for w in topic.target_wgs if w.strip()]
        # Query DB for all meetings matching target WGs
        raw_meetings = meetings_db.search_meetings(wg_name=target_wgs if target_wgs else None)
        all_cohorts = cls.cluster_meetings_into_cohorts(raw_meetings)

        if not all_cohorts:
            return []

        mode = topic.cycle_mode or "past_n_cycles"

        if mode == "past_n_cycles":
            # Filter cohorts up to current date so future meeting placeholders (e.g. 2027) don't displace active/recent ones
            now_iso = datetime.datetime.now().strftime("%Y-%m-%d")
            cur_cohort = get_cohort_key(now_iso) or "2099-Q4"
            valid_cohorts = [c for c in all_cohorts if compare_cohort_keys(c.key, cur_cohort) <= 0]
            if not valid_cohorts:
                valid_cohorts = all_cohorts
            n = max(1, int(topic.past_n_cycles or 4))
            return valid_cohorts[-n:]

        elif mode == "cycle_range":
            from_k = (topic.from_cycle or "").strip().upper()
            to_k = (topic.to_cycle or "").strip().upper()
            filtered = []
            for c in all_cohorts:
                if from_k and compare_cohort_keys(c.key, from_k) < 0:
                    continue
                if to_k and compare_cohort_keys(c.key, to_k) > 0:
                    continue
                filtered.append(c)
            return filtered

        elif mode == "date_range":
            d_from = (topic.date_from or "").strip()
            d_to = (topic.date_to or "").strip()
            filtered_meetings = [
                m for m in raw_meetings
                if (not d_from or str(m.get("start_date", "")) >= d_from)
                and (not d_to or str(m.get("end_date", "")) <= d_to)
            ]
            return cls.cluster_meetings_into_cohorts(filtered_meetings)

        return all_cohorts

    @classmethod
    def get_available_cycle_keys(cls, meetings_db: MeetingsDatabase) -> List[str]:
        """Returns sorted list of distinct cohort keys available in the database."""
        all_meetings = meetings_db.search_meetings()
        cohorts = cls.cluster_meetings_into_cohorts(all_meetings)
        return [c.key for c in cohorts]
