# --- File: src/modules/meetings/core/topic_decision.py ---
"""
Unified Technical Decision & Document Model for Cross-WG Topic Tracking.

Elevates RAN Chairman's Notes text agreements, RAN Endorsed Way Forwards,
and SA/CT Approved pCRs/CRs into identical first-class standing across Working Groups.
"""

import datetime
import logging
import re
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)


@dataclass
class TopicDecision:
    """Represents a substantive technical milestone, agreement, or approved contribution."""
    cohort_key: str                # e.g. "2025-Q4"
    wg: str                        # e.g. "RAN2", "SA2", "CT1"
    meeting_number: str            # e.g. "RAN2 #129", "SA2 #160"
    date: str                      # e.g. "2025-11-12"
    decision_type: str             # "Text Agreement", "Working Assumption", "Approved pCR", "Approved CR", "Endorsed WF", "Agreed LS", "Open Issue / FFS"
    title_or_summary: str          # Technical description or contribution title
    tdoc_id: Optional[str] = None  # None for standalone Chairman's notes text agreements
    source: str = ""               # "Chairman's Notes §8.7" or company source e.g. "Nokia, Qualcomm"
    target_spec: str = ""          # "TR 23.700-99", "TS 38.331", "Architecture Baseline"
    ai: str = ""                   # Agenda Item e.g. "8.7" or "5.14"
    wi: str = ""                   # Work Item / Study Item code
    status: str = "Agreed"         # "Agreed", "Approved", "Working Assumption", "Endorsed", "Open / FFS"
    dossier_rel_path: str = ""     # Relative link to .md file in dossier tree
    source_url: str = ""           # 3GPP Server download HTTPS URL
    notes: str = ""                # In-room discussion remarks or rapporteur action items

    @property
    def is_text_agreement(self) -> bool:
        return self.decision_type in ("Text Agreement", "Working Assumption", "Open Issue / FFS")

    @property
    def is_document(self) -> bool:
        return bool(self.tdoc_id)

    @property
    def badge_icon(self) -> str:
        dt = (self.decision_type or "").lower()
        if "text agreement" in dt or "approved" in dt:
            return "🟢"
        elif "working assumption" in dt:
            return "🟡"
        elif "endorsed wf" in dt or "conclusions" in dt:
            return "🔵"
        elif "ffs" in dt or "open" in dt:
            return "🟣"
        elif "ls" in dt:
            return "📬"
        return "📄"

    def to_markdown_row(self) -> str:
        """Formats the decision as a clean, standardized Markdown table row."""
        if self.dossier_rel_path:
            label = self.tdoc_id or "Agreement"
            doc_cell = f"[{label}]({self.dossier_rel_path})"
        else:
            doc_cell = self.tdoc_id or "—"
        spec_cell = self.target_spec or "—"
        summary_clean = re.sub(r'[\r\n\t]+', ' ', self.title_or_summary).replace("|", "\\|").strip()
        if len(summary_clean) > 200:
            summary_clean = summary_clean[:197] + "..."
        type_badge = f"{self.badge_icon} **{self.decision_type}**"
        return f"| {self.date or self.cohort_key} | **{self.wg}** | {type_badge} | {summary_clean} | {doc_cell} | {spec_cell} |"


def classify_tdoc_decision_type(tdoc_type: str, title: str, status: str) -> Optional[str]:
    """
    Classifies a TDoc into a substantive decision type, or None if it is a routine/unagreed contribution.
    """
    t_type = (tdoc_type or "").lower()
    t_title = (title or "").lower()
    t_stat = (status or "").lower()

    # Is it approved / agreed?
    is_approved = any(w in t_stat for w in ("approved", "agreed", "endorsed"))
    if not is_approved:
        # Check for Way Forward titles that might be noted with agreement
        if "way forward" in t_title or re.search(r'\bwf\b', t_title):
            if "noted" in t_stat or "handled" in t_stat or "agreed" in t_stat:
                return "Endorsed WF"
        return None

    if "pcr" in t_type or "pseudo" in t_type or "pcr" in t_title:
        return "Approved pCR"
    elif "cr" in t_type:
        return "Approved CR"
    elif "ls out" in t_type or "liaison" in t_type:
        return "Agreed LS"
    elif "wf" in t_type or "way forward" in t_title or re.search(r'\bwf\b', t_title):
        return "Endorsed WF"
    elif "work plan" in t_type or "sid" in t_type or "wid" in t_type:
        return "Approved pCR"
    elif "discussion" in t_type and is_approved:
        return "Endorsed WF"

    return "Approved pCR" if "study" in t_title else "Approved CR"


def normalize_chair_notes_decision_type(decision_str: str) -> str:
    """Normalizes decision strings from ChairNotesAgreementsParser."""
    d = (decision_str or "").lower()
    if "working assumption" in d or "wa" == d:
        return "Working Assumption"
    elif "ffs" in d or "open" in d or "to be continued" in d:
        return "Open Issue / FFS"
    elif "way forward" in d or "conclusion" in d:
        return "Endorsed WF"
    return "Text Agreement"
