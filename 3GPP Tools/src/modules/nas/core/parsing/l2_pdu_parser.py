"""Parser for RAN2 Layer 2 specifications (SDAP, PDCP, RLC, MAC) of NR and LTE.

L2 PDU formats are drawn as embedded Visio/EMF figures that cannot be read from DOCX, so this parser
reads Clause 6 ("Protocol data units, formats, and parameters") heading by heading:

* every Clause 6 sub-clause becomes an `ie_definitions` record (prose, caption and table HTML, with a
  placeholder where a figure is drawn);
* leaf format clauses (Cl. 6.2.x; also Cl. 6.1.x for MAC) become "messages" (PDU formats);
* message rows (`ies`) come from the curated `l2_field_catalog` when the format is known, otherwise from the
  Clause 6.3 parameter definitions that are referenced in the format's text.
"""

from __future__ import annotations

import html
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from core.ui.theme import C
from modules.nas.core.parsing.l2_field_catalog import find_l2_format_fields
from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY,
    TAG_P,
    TAG_PPR,
    TAG_TBL,
    W_NS,
    extract_part_index,
)
from modules.nas.core.parsing.protocol_parser_utils import (
    _convert_table_to_html,
    _extract_p_text,
    extract_document_root,
)

L2_SPEC_NUMBERS = frozenset({"37.324", "38.323", "38.322", "38.321", "36.323", "36.322", "36.321"})
_MAC_SPECS = frozenset({"38.321", "36.321"})
_RLC_SPECS = frozenset({"38.322", "36.322"})

RE_HEADING_STYLE = re.compile(r"^(?:Heading|H)\s*\d$", re.IGNORECASE)
RE_HEADING_NUM = re.compile(r"^(\d+(?:\.\d+)*[a-zA-Z]?)(?:\s+(.*))?$")
RE_CAPTION = re.compile(r"^(?:Figure|Table)\s+\d", re.IGNORECASE)
RE_GENERIC_TITLE = re.compile(r"^(General|Scope)\b", re.IGNORECASE)
_GRAPHIC_TAGS = frozenset({"drawing", "pict", "object"})


@dataclass
class _Section:
    clause: str
    title: str
    depth: int
    html_parts: List[str] = field(default_factory=list)
    text_parts: List[str] = field(default_factory=list)
    figures: int = 0
    tables: int = 0

    def has_content(self) -> bool:
        return bool(self.text_parts or self.figures or self.tables)


class L2PduDocxParser:
    """Dedicated parser for TS 37.324 / 38.323 / 38.322 / 38.321 (and LTE TS 36.323 / 36.322 / 36.321)."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.323"):
        self.docx_paths = sorted([Path(p) for p in docx_paths], key=extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _paragraph_style(p_elem: Any) -> str:
        ppr = p_elem.find(TAG_PPR)
        if ppr is None:
            return ""
        ps = ppr.find(f"{W_NS}pStyle")
        return (ps.get(f"{W_NS}val") or "") if ps is not None else ""

    @staticmethod
    def _has_graphic(p_elem: Any) -> bool:
        return any(el.tag.rsplit("}", 1)[-1] in _GRAPHIC_TAGS for el in p_elem.iter() if isinstance(el.tag, str))

    @staticmethod
    def _normalize_key(name: str) -> str:
        return re.sub(r"[^a-zA-Z0-9]", "", name).lower()

    @staticmethod
    def _figure_note_html() -> str:
        return (
            f'<div style="margin: 6px 0; padding: 6px 10px; border: 1px dashed {C.BORDER_STRONG}; '
            f'background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; font-size: 11px;">'
            f"🖼️ PDU format figure (embedded vector graphic in the specification) — see the "
            f"<b>📦 Full Packet</b> view or open the document for the original drawing.</div>"
        )

    # ------------------------------------------------------------------ extraction
    def _collect_sections(self, progress_callback: Optional[Callable[[str, int], None]]) -> List[_Section]:
        sections: List[_Section] = []
        current: Optional[_Section] = None
        started = False
        finished = False

        for path in self.docx_paths:
            root = extract_document_root(path)
            body = root.find(TAG_BODY) if root is not None else None
            if body is None:
                self.logger.warning("Could not read word/document.xml from %s", path)
                continue
            for elem in list(body):
                if finished:
                    break
                if elem.tag == TAG_P:
                    text = _extract_p_text(elem).replace("\u200e", "").replace("\u200f", "").strip()
                    style = self._paragraph_style(elem)
                    if RE_HEADING_STYLE.match(style) and text:
                        m = RE_HEADING_NUM.match(text.strip())
                        num = m.group(1) if m else ""
                        title = (m.group(2) if m else text).strip()
                        if num == "6":
                            started, current = True, None
                            continue
                        if started:
                            if num.startswith("6.") and title:
                                current = _Section(num, title, num.count(".") + 1)
                                sections.append(current)
                            else:
                                finished, current = True, None
                            continue
                        continue
                    if current is None:
                        continue
                    has_graphic = self._has_graphic(elem)
                    if has_graphic:
                        current.figures += 1
                        current.html_parts.append(self._figure_note_html())
                    if not text:
                        continue
                    if style.upper().startswith("B"):
                        text = f"• {text}"
                    current.text_parts.append(text)
                    if RE_CAPTION.match(text):
                        current.html_parts.append(
                            f'<p style="margin: 3px 0; color: {C.TEXT_SECONDARY}; font-size: 11px;"><i>{html.escape(text)}</i></p>'
                        )
                    else:
                        current.html_parts.append(
                            f'<p style="margin: 4px 0; line-height: 1.4; color: {C.TEXT_BODY}; font-size: 11px;">{html.escape(text)}</p>'
                        )
                elif elem.tag == TAG_TBL and current is not None:
                    tbl_html = _convert_table_to_html(elem)
                    current.tables += 1
                    current.html_parts.append(tbl_html)
                    current.text_parts.append(re.sub(r"<[^>]+>", " ", tbl_html))

        if progress_callback:
            progress_callback(f"Found {len(sections)} Clause 6 sub-clauses in TS {self.spec_number}.", 55)
        return sections

    # ------------------------------------------------------------------ record builders
    def _definition_record(self, sec: _Section) -> Dict[str, Any]:
        header = (
            f'<h2 style="color: {C.ACCENT_SKY_DARK}; margin-top: 4px; margin-bottom: 6px; '
            f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; '
            f'padding-bottom: 4px;">{html.escape(sec.title)} (Clause {html.escape(sec.clause)})</h2>'
        )
        return {
            "clause": sec.clause,
            "name": sec.title,
            "ie_name": sec.title,
            "raw_description": header + "".join(sec.html_parts),
            "structure_table": "[]",
        }

    def _parameter_index(self, sections: List[_Section]) -> Dict[str, _Section]:
        params: Dict[str, _Section] = {}
        is_rlc = self.spec_number in _RLC_SPECS
        for sec in sections:
            in_param_clause = (
                (is_rlc and sec.clause.startswith("6.2.3.") and sec.depth == 4)
                or (sec.clause.startswith("6.3.") and sec.depth == 3)
            )
            if not in_param_clause or RE_GENERIC_TITLE.match(sec.title) or len(sec.title) > 60:
                continue
            params[self._normalize_key(sec.title)] = sec
            no_field = re.sub(r"\s+field$", "", sec.title, flags=re.IGNORECASE).strip()
            if no_field and no_field != sec.title:
                params.setdefault(self._normalize_key(no_field), sec)
            for m_abbr in re.finditer(r"\(([A-Za-z0-9/_-]+)\)", sec.title):
                abbr_key = self._normalize_key(m_abbr.group(1))
                if abbr_key:
                    params.setdefault(abbr_key, sec)
        return params

    def _is_format_section(self, sec: _Section, is_leaf: bool) -> bool:
        if not is_leaf or sec.depth < 3 or RE_GENERIC_TITLE.match(sec.title) or not sec.has_content():
            return False
        if self.spec_number in _MAC_SPECS:
            prefixes: Tuple[str, ...] = ("6.1.", "6.2.")
        elif self.spec_number in _RLC_SPECS:
            prefixes = ("6.2.2.",)
        else:
            prefixes = ("6.2.",)
        return sec.clause.startswith(prefixes)

    @staticmethod
    def _direction(title: str, spec_number: str) -> str:
        node = "E-UTRAN" if spec_number.startswith("36.") else "NG-RAN node"
        if re.search(r"sidelink|\bSL\b", title, re.IGNORECASE):
            return "UE ↔ UE"
        if re.search(r"\bDL\b|downlink", title, re.IGNORECASE):
            return f"{node} → UE"
        if re.search(r"\bUL\b|uplink", title, re.IGNORECASE):
            return f"UE → {node}"
        return f"UE ↔ {node}"

    def _message_ies(
        self, sec: _Section, params: Dict[str, _Section]
    ) -> Tuple[List[Dict[str, Any]], Optional[str]]:
        entry = find_l2_format_fields(self.spec_number, sec.title)
        rows: List[Dict[str, Any]] = []
        if entry:
            for fld in entry["fields"]:
                p_sec = params.get(self._normalize_key(fld["name"]))
                clause_ref = p_sec.clause if p_sec else sec.clause
                rows.append(self._ie_row(fld["name"], fld["length"], fld["presence"], fld["format"], clause_ref, fld["description"]))
            return rows, entry.get("direction")

        blob = " ".join(sec.text_parts)
        for p_sec in params.values():
            pattern = rf"(?<![A-Za-z0-9]){re.escape(p_sec.title)}(?![A-Za-z0-9])"
            if re.search(pattern, blob):
                rows.append(self._ie_row(p_sec.title, f"See Clause {p_sec.clause}", "M", "Parameter", p_sec.clause, ""))
        if not rows:
            rows.append(
                self._ie_row(
                    sec.title,
                    f"See Clause {sec.clause}",
                    "M",
                    "Octet / Bit fields",
                    sec.clause,
                    f"PDU format defined in TS {self.spec_number} Clause {sec.clause}.",
                )
            )
        return rows, None

    @staticmethod
    def _ie_row(name: str, length: str, presence: str, fmt: str, clause: str, desc: str) -> Dict[str, Any]:
        type_ref = f"{name} (Clause {clause})"
        return {
            "iei": "-",
            "name": name,
            "field": name,
            "information_element": name,
            "field_path": name,
            "depth": 0,
            "type": type_ref,
            "type_reference": type_ref,
            "presence": presence,
            "format": fmt,
            "length": length,
            "applicability": "Uu",
            "field_description": desc,
        }

    # ------------------------------------------------------------------ public API
    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        if progress_callback:
            progress_callback(f"Reading Clause 6 of TS {self.spec_number}...", 25)
        sections = self._collect_sections(progress_callback)
        params = self._parameter_index(sections)

        definitions = [self._definition_record(s) for s in sections if s.has_content()]
        messages: List[Dict[str, Any]] = []
        for idx, sec in enumerate(sections):
            nxt = sections[idx + 1] if idx + 1 < len(sections) else None
            is_leaf = not (nxt and nxt.clause.startswith(sec.clause + "."))
            if not self._is_format_section(sec, is_leaf):
                continue
            ies, cat_dir = self._message_ies(sec, params)
            prose = next((t for t in sec.text_parts if len(t) > 30 and not RE_CAPTION.match(t)), "")
            direction = cat_dir or self._direction(sec.title, self.spec_number)
            messages.append({
                "clause": sec.clause,
                "clause_ref": sec.clause,
                "name": sec.title,
                "message_name": sec.title,
                "table_caption": f"{sec.title} (Clause {sec.clause})",
                "description": prose or f"{sec.title} format (Clause {sec.clause}).",
                "direction": direction,
                "ies": ies,
                "fields": ies,
            })

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} PDU formats and {len(definitions)} definitions.", 90)
        self.logger.info(
            "Completed parsing TS %s: %d PDU format(s), %d definition(s)", self.spec_number, len(messages), len(definitions)
        )
        return messages, definitions
