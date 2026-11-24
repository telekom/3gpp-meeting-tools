"""AI/WI/keyword filtering, sorting keys, agenda CSV mapping, and URL candidate builders."""

from __future__ import annotations

import csv
import os
import re
from pathlib import Path
from typing import Dict, List, Tuple

from core.utils.paths import to_extended_path
from modules.meetings.ui.tdocs_models import TDocsTableModel


class AgreementFilterMixin:
    """Mixin providing AI/WI/keyword matching, sorting, agenda.csv parsing, and URL candidates."""

    @staticmethod
    def _normalize_ai(value: str) -> str:
        text = str(value or "").strip()
        text = re.sub(r"^AI\s*[:#-]?\s*", "", text, flags=re.IGNORECASE)
        return re.sub(r"\s+", "", text).lower().rstrip(".")

    @staticmethod
    def _ai_sort_key(ai_str: str) -> Tuple:
        if not ai_str:
            return (9999,)
        parts = re.findall(r"\d+", str(ai_str))
        if parts:
            return tuple(int(p) for p in parts)
        return (9998, str(ai_str))

    def _kind_enabled(self, kind: str) -> bool:
        low = kind.lower()
        if "ffs" in low or "open issue" in low:
            return self.include_ffs
        if "agreed tdoc" in low or "pcr" in low or "tp" in low:
            return self.include_agreed_tdocs
        return self.include_agreements

    def _matches_filter(self, record: dict) -> bool:
        # 1. WI Filter (supports multi-token base extraction just like TDocsTableModel)
        wi_hit = None
        if self.target_wis:
            rec_wi = str(record.get("Work Item", "")).strip()
            rec_ai_title = str(record.get("AI Title", "")).strip()
            rec_tdoc_title = str(record.get("TDoc Title", "")).strip()
            combined_wi_haystack = f"{rec_wi} {rec_ai_title} {rec_tdoc_title}".lower()
            rec_base_tokens = TDocsTableModel._extract_base_tokens(rec_wi)

            wi_hit = bool(rec_base_tokens & self.target_wi_base_tokens) or any(
                w.lower() in combined_wi_haystack for w in self.target_wis
            )

        # 2. AI Filter (exact match OR prefix match, e.g. `9.3` matches `9.3.1.2`)
        ai_hit = None
        if self.target_ais:
            rec_ai = self._normalize_ai(record.get("AI", ""))
            if not rec_ai:
                ai_hit = False
            else:
                ai_hit = any(
                    rec_ai == target_ai or rec_ai.startswith(f"{target_ai}.")
                    for target_ai in self.target_ais
                )

        active_facets = [h for h in (wi_hit, ai_hit) if h is not None]
        if active_facets:
            if self.wi_ai_or:
                if not any(active_facets):
                    return False
            elif not all(active_facets):
                return False

        # 3. Keyword / Topic Query Filter
        if self.keyword_query:
            haystack = " ".join([
                str(record.get("Agreement Text", "")),
                str(record.get("AI Title", "")),
                str(record.get("Work Item", "")),
                str(record.get("TDoc", "")),
                str(record.get("TDoc Title", "")),
                str(record.get("Source", "")),
                str(record.get("AI", "")),
            ]).lower()
            if "," in self.keyword_query or " OR " in self.keyword_query:
                or_terms = [
                    t.strip().lower()
                    for t in re.split(r",|\bOR\b", self.keyword_query)
                    if t.strip()
                ]
                if or_terms and not any(term in haystack for term in or_terms):
                    return False
            else:
                and_terms = [t.strip().lower() for t in self.keyword_query.split() if t.strip()]
                if and_terms and not all(term in haystack for term in and_terms):
                    return False

        return True

    @staticmethod
    def _extract_related_wis_from_row(row: dict) -> str:
        for key in ["Related WIs", "Related WI", "Related WI(s)", "Work Item", "Work Items", "WI", "WID"]:
            val = str(row.get(key, "")).strip()
            if val and val.lower() not in ["", "none", "unknown", "-"]:
                items = [w.strip() for w in re.split(r"[,;]+", val) if w.strip()]
                if items:
                    return ", ".join(items)
        desc = str(row.get("Agenda item description", "")).strip()
        if desc:
            m = re.search(r"\(([A-Za-z0-9_ -]{3,35})\)", desc)
            if m:
                cand = m.group(1).strip()
                if "_" in cand or any(c.isdigit() for c in cand):
                    return cand
        return ""

    @staticmethod
    def _load_agenda_csv_map(agenda_dir: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
        ai_wis: Dict[str, str] = {}
        ai_titles: Dict[str, str] = {}
        csv_path = agenda_dir / "agenda.csv"
        if not os.path.exists(to_extended_path(csv_path)):
            return ai_wis, ai_titles
        try:
            with open(to_extended_path(csv_path), "r", encoding="utf-8", errors="ignore") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    ai = str(row.get("ai_number", "")).strip()
                    desc = str(row.get("description", "")).strip()
                    acr = str(row.get("acronym", "")).strip()
                    if ai:
                        if acr:
                            ai_wis[ai] = acr
                        if desc:
                            ai_titles[ai] = desc
        except Exception:
            pass
        return ai_wis, ai_titles

    @staticmethod
    def _build_meeting_candidate_urls(mtg: dict) -> List[str]:
        urls = []
        raw_url = str(mtg.get("url_key", "")).strip()
        if raw_url:
            full = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            urls.append(full.rstrip("/"))
        docs_url = str(mtg.get("docs_folder_url", "")).strip()
        if docs_url:
            full_docs = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            base = re.sub(r"/Docs/?$", "", full_docs.rstrip("/"), flags=re.IGNORECASE)
            if base not in urls:
                urls.append(base)
        return urls
