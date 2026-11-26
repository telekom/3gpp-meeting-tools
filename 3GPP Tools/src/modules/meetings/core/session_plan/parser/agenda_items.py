# --- File: src/modules/meetings/core/session_plan/parser/agenda_items.py ---
"""Tokenization, room override extraction, KI# resolution, and AI range expansion."""

import re
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan.session_plan_models import (
    AgendaSlotItem,
    expand_numeric_ai_range,
)

from .constants import ADMIN_NOTE_RE, ROOM_OVERRIDE_RE


class AgendaItemsParserMixin:
    """Mixin extracting room overrides and AgendaSlotItems from cell text."""

    @classmethod
    def _extract_room_override_and_tokens(cls, raw_text: str) -> Tuple[str, List[str]]:
        """Separates physical room overrides (Breakout 1/2, Main Room) from agenda topics."""
        t = raw_text.replace("\xa0", " ")
        raw_tokens = re.split(r"\n+|\s{2,}|\s*(?<=\])\s+|\s*(?<=\))\s+(?=[A-Z0-9#])", t)

        room_override = ""
        clean_tokens = []

        for tok in raw_tokens:
            tok_s = tok.strip()
            if not tok_s:
                continue
            m_room = ROOM_OVERRIDE_RE.match(tok_s)
            if m_room:
                r_str = m_room.group(1).title()
                if "Main" in r_str:
                    room_override = "Main Room"
                elif "Breakout 1" in r_str or "Bmo" in r_str:
                    room_override = "Breakout 1"
                elif "Breakout 2" in r_str or "Percheron" in r_str:
                    room_override = "Breakout 2"
                else:
                    room_override = r_str
                rest = tok_s[m_room.end():].strip(" :)-–—")
                if rest:
                    clean_tokens.append(rest)
            else:
                clean_tokens.append(tok_s)

        # Join continuation lines: e.g. "Pre-Rel-19 maintenance" + "(5.x, 6.x, 7.x )" or split "(20.6.12 -" + "16)"
        merged: List[str] = []
        for tok_s in clean_tokens:
            if ADMIN_NOTE_RE.search(tok_s):
                continue
            if merged and merged[-1].count("(") > merged[-1].count(")"):
                merged[-1] = f"{merged[-1]} {tok_s}"
            elif merged and tok_s.startswith("(") and re.search(r"\d", tok_s) and not re.search(r"\(\d", merged[-1]):
                merged[-1] = f"{merged[-1]} {tok_s}"
            elif merged and (
                tok_s.startswith("(") or not re.search(r"\(|\b(?:KI#|FS_|AIML_|\d{1,2}\.(?:\d+|[xX]))", tok_s)
            ):
                merged[-1] = f"{merged[-1]} {tok_s}"
            else:
                merged.append(tok_s)

        return room_override, merged

    @staticmethod
    def _balance_parentheses(s: str) -> str:
        """Strips leftover unmatched parentheses and surrounding punctuation without unclosing labels."""
        s = s.strip(" -–—,")
        while s.endswith(")") and s.count(")") > s.count("("):
            s = s[:-1].strip()
        while s.startswith("(") and s.count("(") > s.count(")"):
            s = s[1:].strip()
        return s

    @classmethod
    def _extract_agenda_items(
        cls,
        raw_text: str,
        default_drafting: bool = False,
        default_tentative: bool = False,
        ki_map: Optional[Dict[str, str]] = None,
    ) -> Tuple[str, List[AgendaSlotItem]]:
        """
        Extracts room override and individual Agenda Items, expanding ranges like:
        - KI#18 (AI for Network) (20.6.18) [5+15]
        - KI#12 – KI#16 (20.6.12 - 16) [5,0,2,0,1] -> expands to 20.6.12 .. 20.6.16
        - Pre-Rel-19 maintenance (5.x, 6.x, 7.x) [9] -> expands to 5.*, 6.*, 7.*
        - Handles typos: K#1.2 -> KI#1.2, unclosed UIA_ARC (19.8.2 [4]
        - Resolves implicit KI numbers: KI#21 -> 20.6.21
        """
        items: List[AgendaSlotItem] = []
        if not raw_text or raw_text.strip() in ("-", "—", "Revisions", "Revisions (tbd)"):
            return "", items

        if ki_map is None:
            ki_map = {}

        room_override, tokens = cls._extract_room_override_and_tokens(raw_text)

        for tok in tokens:
            tok_clean = tok.strip()
            if not tok_clean or tok_clean in ("-", "—", "Revisions", "Revisions (tbd)"):
                continue
            if ADMIN_NOTE_RE.search(tok_clean):
                continue

            # Single AIs like "Opening (1), Agenda (2), Reports (3), Outgoing LSs (5.1), Work planning (30.x)"
            single_ai_matches = list(
                re.finditer(r"([A-Za-z][A-Za-z\s]*?)\s*\((\d{1,2}(?:\.\d+|\.[xX])*)\)", tok_clean)
            )
            if len(single_ai_matches) >= 2:
                for sm in single_ai_matches:
                    t_lbl = sm.group(1).strip(" ,")
                    a_num = re.sub(r"\.[xX]\b", ".x", sm.group(2).strip())
                    items.append(
                        AgendaSlotItem(
                            ai_number=a_num,
                            display_expr=f"{t_lbl} ({a_num})",
                            topic_label=t_lbl,
                            is_drafting=default_drafting or ("drafting" in tok_clean.lower()),
                            is_tentative=default_tentative or ("(?)" in tok_clean),
                        )
                    )
                continue

            # 1. Extract TDoc Count in brackets [5+15] or [7]
            tdoc_count = ""
            m_count = re.search(r"\[([\d\+\,\s]+)\]", tok_clean)
            if m_count:
                tdoc_count = m_count.group(1).strip()

            # 2. Extract AI Number: (20.6.18), (20.6.12 - 16), (20.6.12 – 16), (20.6.12 – 20.6.16), (30.x), (5.x, 6.x, 7.x )
            _D = r"[\u2010-\u2015\u2212\-]"
            _NUM_OR_WILD = r"\d{1,2}(?:\.(?:\d+(?:[a-z](?![a-z]))?|[xX]))+"
            _RANGE_SEG = rf"{_NUM_OR_WILD}(?:\s*{_D}\s*(?:{_NUM_OR_WILD}|\d+))?"

            paren_ai_matches = list(
                re.finditer(
                    rf"\(\s*({_RANGE_SEG}(?:\s*,\s*{_RANGE_SEG})*)\s*\)",
                    tok_clean,
                )
            )
            ai_str = ", ".join(m.group(1).strip() for m in paren_ai_matches) if paren_ai_matches else ""

            if not ai_str:
                # Look for standalone AI lists/ranges: e.g. "20.48, 20.49", "20.6.12 – 16", "30.x", "5.x, 6.x"
                standalone_matches = re.findall(
                    rf"\b({_RANGE_SEG}(?:,\s*{_RANGE_SEG})*)\b",
                    tok_clean,
                )
                if standalone_matches:
                    ai_str = ", ".join(m.strip() for m in standalone_matches if m.strip())

            # 3. Check KI# ranges & multi-KI lists if ai_str is empty: e.g. "KI#12 – KI#16" or "KI#3, KI#5, KI#11, KI#17?/KI#22?"
            if not ai_str:
                ki_resolved: List[str] = []
                covered_spans: List[Tuple[int, int]] = []
                for m_kr in re.finditer(
                    rf"\b(?:KI|K)\s*#?\s*(\d+)\s*{_D}\s*(?:(?:KI|K)\s*#?)?(\d+)\b", tok_clean, re.I
                ):
                    start_k, end_k = int(m_kr.group(1)), int(m_kr.group(2))
                    if start_k <= end_k and end_k - start_k <= 50:
                        ki_resolved.append(f"20.6.{start_k} - 20.6.{end_k}")
                        covered_spans.append((m_kr.start(), m_kr.end()))
                for m_ki in re.finditer(r"\b(?:KI|K)\s*#?\s*(\d+(?:\.\d+)?)\b", tok_clean, re.I):
                    if any(s <= m_ki.start() < e for s, e in covered_spans):
                        continue
                    k_num = m_ki.group(1)
                    if k_num in ki_map:
                        val = ki_map[k_num]
                    elif "." in k_num:
                        val = f"20.6.{k_num.split('.')[0]}"
                    else:
                        val = f"20.6.{k_num}"
                    if val not in ki_resolved:
                        ki_resolved.append(val)
                if ki_resolved:
                    ai_str = ", ".join(ki_resolved)
            else:
                # Populate ki_map from explicit AI bindings in this slot
                m_kr = re.search(rf"\b(?:KI|K)\s*#?\s*(\d+)\s*{_D}\s*(?:(?:KI|K)\s*#?)?(\d+)\b", tok_clean, re.I)
                if m_kr:
                    sk, ek = int(m_kr.group(1)), int(m_kr.group(2))
                    for kv in range(sk, min(ek, sk + 50) + 1):
                        ki_map.setdefault(str(kv), f"20.6.{kv}")
                else:
                    m_ki = re.search(r"\b(?:KI|K)\s*#?\s*(\d+(?:\.\d+)?)\b", tok_clean, re.I)
                    if m_ki:
                        ki_map[m_ki.group(1)] = ai_str

            # Normalize all Unicode dash variants & .X wildcards cleanly
            if ai_str:
                ai_str = re.sub(_D, "-", ai_str)
                ai_str = re.sub(r"\.[xX]\b", ".x", ai_str)

            # 4. Clean topic label
            clean_label = re.sub(r"\bK#(\d+)", r"KI#\1", tok_clean)
            clean_label = re.sub(r"\[.*?\]", "", clean_label)
            for pm in paren_ai_matches:
                clean_label = clean_label.replace(pm.group(0), "")
            if ai_str and not paren_ai_matches:
                clean_label = re.sub(r"\(\s*" + re.escape(ai_str) + r"\s*\)", "", clean_label, flags=re.I)
            clean_label = re.sub(r"\s+,", ",", clean_label)
            clean_label = re.sub(r"\s+", " ", clean_label)
            clean_label = cls._balance_parentheses(clean_label)

            # Expand ranges and comma-separated AIs uniformly
            is_tok_drafting = default_drafting or ("drafting" in tok_clean.lower())
            is_tok_tentative = default_tentative or ("(?)" in tok_clean)
            if ai_str:
                raw_parts = [p.strip() for p in ai_str.split(",") if p.strip()]
                expanded_ais: List[str] = []
                has_range = False
                for part in raw_parts:
                    if "-" in part or re.search(_D, part):
                        has_range = True
                        expanded_ais.extend(cls._expand_numeric_range(part))
                    else:
                        expanded_ais.append(part)

                counts = [c.strip() for c in tdoc_count.split(",")] if (has_range and "," in tdoc_count) else []
                for idx, single_ai in enumerate(expanded_ais):
                    c_val = counts[idx] if idx < len(counts) else tdoc_count
                    items.append(
                        AgendaSlotItem(
                            ai_number=single_ai,
                            display_expr=tok_clean,
                            topic_label=clean_label or (single_ai if len(expanded_ais) > 1 else tok_clean),
                            tdoc_count_str=c_val,
                            is_drafting=is_tok_drafting,
                            is_tentative=is_tok_tentative,
                        )
                    )
            else:
                if clean_label:
                    items.append(
                        AgendaSlotItem(
                            ai_number="",
                            display_expr=tok_clean,
                            topic_label=clean_label,
                            tdoc_count_str=tdoc_count,
                            is_drafting=is_tok_drafting,
                            is_tentative=is_tok_tentative,
                        )
                    )

        return room_override, items

    @staticmethod
    def _expand_numeric_range(range_str: str) -> List[str]:
        """Expands '20.6.12 - 16', '20.6.12 – 16', or '20.6.12 – 20.6.16' to concrete AI list."""
        return expand_numeric_ai_range(range_str)
