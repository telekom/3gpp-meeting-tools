import html
import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY, TAG_P, TAG_TBL, TAG_TR, TAG_TC,
    RE_PART_INDEX, RE_CLAUSE_HEADER, RE_DESC_TABLE,
    RE_TYPE_DECL, RE_TYPE_KIND, RE_FIELD_LINE, RE_STRIP_KEYWORDS
)
from modules.nas.core.parsing.protocol_parser_utils import (
    extract_document_root, _extract_p_text, _extract_tc_text, _convert_table_to_html
)


class BaseAsn1DocxParser:
    """Base class for 3GPP ASN.1 specification parsing."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.331"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    def _scan_xml_documents(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[str, Dict[str, str], Dict[str, Dict[str, str]], Dict[str, str]]:
        """
        Pass 1: Scans Word XML trees, extracting ASN.1 code blocks between
        -- ASN1START and -- ASN1STOP, clause headings, and field description tables.
        """
        raw_asn1_blocks: List[str] = []
        field_desc_tables: Dict[str, str] = {}
        field_individual_descs: Dict[str, Dict[str, str]] = {}
        clause_map: Dict[str, str] = {}
        total_files = len(self.docx_paths)
        current_clause_num = "6.2"

        for file_idx, docx_path in enumerate(self.docx_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 40)
                progress_callback(f"Scanning {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            in_asn1 = False
            current_asn1_lines: List[str] = []
            last_p_text = ""
            current_heading_name = ""

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_p_text = p_text

                    match_clause = RE_CLAUSE_HEADER.match(p_text)
                    if match_clause:
                        current_clause_num = match_clause.group(1).strip()
                    elif p_text.startswith("–") or p_text.startswith("-"):
                        current_heading_name = p_text.lstrip("–- ").strip()
                        if current_heading_name:
                            clause_map[current_heading_name.lower()] = current_clause_num

                    if "-- ASN1START" in p_text:
                        in_asn1 = True
                        current_asn1_lines = []
                        continue
                    elif "-- ASN1STOP" in p_text:
                        in_asn1 = False
                        raw_asn1_blocks.append("\n".join(current_asn1_lines))
                        current_asn1_lines = []
                        continue

                    if in_asn1:
                        current_asn1_lines.append(p_text)

                elif elem.tag == TAG_TBL:
                    match_tbl = RE_DESC_TABLE.search(last_p_text)
                    target_name = match_tbl.group(1).strip() if match_tbl else current_heading_name

                    if target_name:
                        target_key = target_name.lower()
                        field_desc_tables[target_key] = _convert_table_to_html(elem)

                        rows = elem.findall(TAG_TR)
                        field_dict: Dict[str, str] = {}
                        for r in rows:
                            cells = r.findall(TAG_TC)
                            if len(cells) >= 2:
                                fname = _extract_tc_text(cells[0]).strip()
                                fdesc = _extract_tc_text(cells[1]).strip()
                                if fname and fdesc and "field descriptions" not in fname.lower():
                                    field_dict[fname.lower()] = fdesc
                        if field_dict:
                            field_individual_descs[target_key] = field_dict

        full_asn1_module = "\n".join(raw_asn1_blocks)
        if not full_asn1_module.strip():
            full_asn1_module = self._fallback_extract_asn1()

        return full_asn1_module, field_desc_tables, field_individual_descs, clause_map

    def _fallback_extract_asn1(self) -> str:
        collected = []
        for docx_path in self.docx_paths:
            root = extract_document_root(docx_path)
            if root is not None:
                body = root.find(TAG_BODY)
                if body is not None:
                    for p in body.findall(TAG_P):
                        t = _extract_p_text(p)
                        if any(kw in t for kw in ("::=", "SEQUENCE", "CHOICE", "PROTOCOL-IES")):
                            collected.append(t)
        return "\n".join(collected)

    def _extract_asn1_type_definitions(self, asn1_text: str) -> Dict[str, Dict[str, Any]]:
        """Parses raw ASN.1 text into structured type definitions with clean module boundary truncation."""
        type_defs: Dict[str, Dict[str, Any]] = {}
        clean_lines = [line.split("--")[0] for line in asn1_text.splitlines()]
        cleaned_text = "\n".join(clean_lines)

        matches = list(RE_TYPE_DECL.finditer(cleaned_text))
        for i, match in enumerate(matches):
            type_name = match.group(1).strip()
            start_pos = match.end()
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(cleaned_text)

            raw_block = cleaned_text[match.start():end_pos].strip()

            # Truncate module end keyword to prevent header spillovers
            end_kw_idx = raw_block.find("\nEND")
            if end_kw_idx != -1:
                raw_block = raw_block[:end_kw_idx].strip()

            def_body = cleaned_text[start_pos:end_pos].strip()

            kind_match = RE_TYPE_KIND.match(def_body)
            type_kind = kind_match.group(1).upper() if kind_match else "TYPE"
            fields: List[Dict[str, Any]] = []

            brace_start = def_body.find("{")
            if brace_start >= 0:
                depth = 0
                brace_end = -1
                for idx, char in enumerate(def_body[brace_start:], start=brace_start):
                    if char == "{":
                        depth += 1
                    elif char == "}":
                        depth -= 1
                        if depth == 0:
                            brace_end = idx
                            break

                if brace_end > brace_start:
                    inner_content = def_body[brace_start + 1:brace_end]
                    fields = self._parse_sequence_fields(inner_content, type_kind)

            type_defs[type_name] = {
                "kind": type_kind,
                "fields": fields,
                "raw_asn1": raw_block,
                "clause": type_name,
            }

        return type_defs

    def _parse_sequence_fields(self, inner_asn1: str, parent_kind: str) -> List[Dict[str, Any]]:
        """Extracts field records from inside a SEQUENCE or CHOICE definition block."""
        fields: List[Dict[str, Any]] = []
        cleaned = re.sub(r"\[\[|\]\]", " ", inner_asn1)

        tokens, current = [], []
        depth = p_depth = 0

        for char in cleaned:
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
            elif char == "(":
                p_depth += 1
            elif char == ")":
                p_depth -= 1

            if char == "," and depth == 0 and p_depth == 0:
                tokens.append("".join(current).strip())
                current = []
            else:
                current.append(char)

        if current:
            tokens.append("".join(current).strip())

        for tok in tokens:
            tok = " ".join(tok.split()).strip()
            if not tok or tok == "..." or tok.startswith("--"):
                continue

            field_match = RE_FIELD_LINE.match(tok)
            if field_match:
                fname = field_match.group(1).strip()
                rest = field_match.group(2).strip()

                presence = "O" if "OPTIONAL" in rest else ("M" if parent_kind == "SEQUENCE" else "O")
                ftype = RE_STRIP_KEYWORDS.sub("", rest).strip()

                fields.append({
                    "name": fname,
                    "type": ftype,
                    "presence": presence,
                    "format": parent_kind,
                    "iei": "",
                })

        return fields

    def _build_inspector_html(
            self,
            type_name: str,
            assigned_clause: str,
            raw_asn1: str,
            desc_table_html: str,
    ) -> str:
        """Constructs HTML presentation for the Inspector pane."""
        raw_asn1_html = (
            f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; '
            f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A; '
            f'white-space: pre-wrap;">{html.escape(raw_asn1)}</pre>'
        )

        if desc_table_html:
            desc_table_html = (
                f'<h3 style="font-size: 12px; font-weight: bold; color: #1E293B; margin-top: 10px; '
                f'margin-bottom: 4px;">{type_name} field descriptions</h3>{desc_table_html}'
            )

        return (
            f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 6px; '
            f'font-family: Segoe UI, sans-serif;">{type_name} (Clause {assigned_clause})</h2>'
            f'{raw_asn1_html}{desc_table_html}'
        )