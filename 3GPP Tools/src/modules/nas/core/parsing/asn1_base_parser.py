import html
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY, TAG_P, TAG_TBL, TAG_TR, TAG_TC,
    RE_PART_INDEX, RE_CLAUSE_HEADER, RE_DESC_TABLE,
    RE_TYPE_DECL, RE_TYPE_KIND, RE_FIELD_LINE, RE_STRIP_KEYWORDS,
    RE_COND_NEED, RE_MAJOR_BOUNDARY
)
from modules.nas.core.parsing.protocol_parser_utils import (
    extract_document_root, _extract_p_text, _extract_tc_text, _convert_table_to_html
)


class BaseAsn1DocxParser:
    """Base class for 3GPP ASN.1 specification parsing with Clause 9.2/9.3 and Clause 6.2/6.3 prose and table extraction."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.331"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    @staticmethod
    def _normalize_key(name: str) -> str:
        """Strips non-alphanumeric characters and converts to lowercase for resilient matching."""
        return re.sub(r'[^a-zA-Z0-9]', '', name).lower()

    def _scan_xml_documents(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[str, Dict[str, str], Dict[str, Dict[str, str]], Dict[str, str]]:
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
            current_section_def: Optional[Dict[str, Any]] = None

            def _finalize_section():
                nonlocal current_section_def
                if not current_section_def:
                    return
                s_name = current_section_def["name"]
                s_clause = current_section_def["clause"]
                s_prose = current_section_def["prose"]
                s_tables = current_section_def["tables"]

                if s_name and (s_prose or s_tables):
                    norm_k = self._normalize_key(s_name)
                    lower_k = s_name.lower()

                    clause_map[lower_k] = s_clause
                    if norm_k:
                        clause_map[norm_k] = s_clause

                    html_content = "".join(s_prose) + "".join(s_tables)
                    field_desc_tables[lower_k] = html_content
                    if norm_k:
                        field_desc_tables[norm_k] = html_content

                current_section_def = None

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_p_text = p_text

                    if "-- ASN1START" in p_text:
                        _finalize_section()
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
                        continue

                    match_clause = RE_CLAUSE_HEADER.match(p_text)
                    if match_clause:
                        _finalize_section()
                        current_clause_num = match_clause.group(1).strip()
                        clause_title = match_clause.group(2).strip()

                        if clause_title:
                            current_heading_name = clause_title
                            norm_t = self._normalize_key(clause_title)
                            clause_map[clause_title.lower()] = current_clause_num
                            if norm_t:
                                clause_map[norm_t] = current_clause_num

                            if current_clause_num.startswith(("9.2", "9.3", "6.2", "6.3", "6.6", "D.6")):
                                current_section_def = {
                                    "clause": current_clause_num,
                                    "name": clause_title,
                                    "prose": [],
                                    "tables": [],
                                }
                        continue

                    if p_text.startswith("–") or p_text.startswith("-"):
                        _finalize_section()
                        current_heading_name = p_text.lstrip("–- ").strip()
                        if current_heading_name:
                            norm_h = self._normalize_key(current_heading_name)
                            clause_map[current_heading_name.lower()] = current_clause_num
                            if norm_h:
                                clause_map[norm_h] = current_clause_num

                            current_section_def = {
                                "clause": current_clause_num,
                                "name": current_heading_name,
                                "prose": [],
                                "tables": [],
                            }
                        continue

                    if RE_MAJOR_BOUNDARY.match(p_text) and not p_text.startswith(("9.", "6.")):
                        _finalize_section()
                        continue

                    if current_section_def:
                        if p_text.startswith(("Direction:", "Direction :")):
                            current_section_def["prose"].append(
                                f'<p style="margin: 4px 0 6px 0; font-weight: bold; color: #0284C7; '
                                f'font-size: 11px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_section_def["prose"].append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">'
                                f'{html.escape(p_text)}</div>'
                            )
                        elif not p_text.startswith(("Table ", "Figure ")) and len(current_section_def["prose"]) < 8:
                            current_section_def["prose"].append(
                                f'<p style="margin: 3px 0; line-height: 1.4; color: #334155; '
                                f'font-size: 11px;">{html.escape(p_text)}</p>'
                            )

                elif elem.tag == TAG_TBL:
                    tbl_html = _convert_table_to_html(elem)

                    match_tbl = RE_DESC_TABLE.search(last_p_text)
                    rrc_target = match_tbl.group(1).strip() if match_tbl else ""
                    if rrc_target:
                        rk = rrc_target.lower()
                        field_desc_tables[rk] = tbl_html
                        field_desc_tables[self._normalize_key(rrc_target)] = tbl_html

                    if current_section_def and tbl_html:
                        current_section_def["tables"].append(tbl_html)

                    rows = elem.findall(TAG_TR)
                    if rows:
                        header_cells = [_extract_tc_text(tc).lower() for tc in rows[0].findall(TAG_TC)]
                        name_col, desc_col = 0, -1
                        for idx, htext in enumerate(header_cells):
                            if any(k in htext for k in ("semantics", "description", "explanation")):
                                desc_col = idx
                                break
                        if desc_col == -1 and len(header_cells) >= 2:
                            desc_col = 1

                        field_dict: Dict[str, str] = {}
                        for r in rows[1:]:
                            cells = r.findall(TAG_TC)
                            if len(cells) > max(name_col, desc_col) and desc_col != -1:
                                fname = _extract_tc_text(cells[name_col]).strip().lstrip("> ")
                                fdesc = _extract_tc_text(cells[desc_col]).strip()
                                if fname and fdesc and "field descriptions" not in fname.lower():
                                    field_dict[fname.lower()] = fdesc
                                    norm_fn = self._normalize_key(fname)
                                    if norm_fn:
                                        field_dict[norm_fn] = fdesc

                        if field_dict:
                            target_name = (current_section_def["name"] if current_section_def else current_heading_name)
                            if target_name:
                                field_individual_descs[target_name.lower()] = field_dict
                                field_individual_descs[self._normalize_key(target_name)] = field_dict

            _finalize_section()

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
        """Parses raw ASN.1 text into structured type definitions, retaining full comments and inline structures."""
        type_defs: Dict[str, Dict[str, Any]] = {}
        matches = list(RE_TYPE_DECL.finditer(asn1_text))

        for i, match in enumerate(matches):
            # Ensure ::= is not inside an ASN.1 comment
            line_start = asn1_text.rfind("\n", 0, match.start())
            line_start = 0 if line_start == -1 else line_start + 1
            prefix_line = asn1_text[line_start:match.start()]
            if "--" in prefix_line:
                continue

            type_name = match.group(1).strip()
            start_pos = match.end()
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(asn1_text)

            raw_block = asn1_text[match.start():end_pos].strip()
            end_kw_idx = raw_block.find("\nEND")
            if end_kw_idx != -1:
                raw_block = raw_block[:end_kw_idx].strip()

            def_body = asn1_text[start_pos:end_pos].strip()
            kind_match = RE_TYPE_KIND.match(def_body)
            type_kind = kind_match.group(1).upper() if kind_match else "TYPE"
            fields: List[Dict[str, Any]] = []

            brace_start = self._find_first_structural_brace(def_body)
            if brace_start >= 0:
                brace_end = self._find_matching_brace(def_body, brace_start)
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

    @staticmethod
    def _find_first_structural_brace(text: str) -> int:
        in_comment = False
        i = 0
        n = len(text)
        while i < n:
            c = text[i]
            if not in_comment and c == '-' and i + 1 < n and text[i+1] == '-':
                in_comment = True
                i += 2
                continue
            elif in_comment:
                if c == '\n':
                    in_comment = False
                elif c == '-' and i + 1 < n and text[i+1] == '-':
                    in_comment = False
                    i += 2
                    continue
                i += 1
                continue
            if c == '{':
                return i
            i += 1
        return -1

    @staticmethod
    def _find_matching_brace(text: str, start_idx: int) -> int:
        depth = 0
        in_comment = False
        i = start_idx
        n = len(text)
        while i < n:
            c = text[i]
            if not in_comment and c == '-' and i + 1 < n and text[i+1] == '-':
                in_comment = True
                i += 2
                continue
            elif in_comment:
                if c == '\n':
                    in_comment = False
                elif c == '-' and i + 1 < n and text[i+1] == '-':
                    in_comment = False
                    i += 2
                    continue
                i += 1
                continue

            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    return i
            i += 1
        return -1

    def _parse_sequence_fields(self, inner_asn1: str, parent_kind: str) -> List[Dict[str, Any]]:
        """Extracts field records while fully preserving inline SEQUENCE/CHOICE bodies and comments."""
        fields: List[Dict[str, Any]] = []
        cleaned = re.sub(r"\[\[|\]\]", " ", inner_asn1)
        tokens = self._split_asn1_tokens(cleaned)

        for tok in tokens:
            tok = tok.strip()
            if not tok or tok == "..." or tok.startswith("--"):
                continue

            field_match = RE_FIELD_LINE.match(tok)
            if not field_match:
                continue

            fname = field_match.group(1).strip()
            rest = field_match.group(2).strip()

            # Handle inline SEQUENCE / CHOICE blocks without truncating inner keywords
            if rest.startswith(("SEQUENCE", "CHOICE")) and "{" in rest:
                brace_start = self._find_first_structural_brace(rest)
                brace_end = self._find_matching_brace(rest, brace_start) if brace_start >= 0 else -1

                if brace_start >= 0 and brace_end > brace_start:
                    raw_type_body = rest[:brace_end + 1].strip()
                    tail = rest[brace_end + 1:].strip()
                    inner_body = rest[brace_start + 1:brace_end]

                    # Parse presence and condition from trailing tail
                    need_match = RE_COND_NEED.search(tail) or RE_COND_NEED.search(tok)
                    if need_match:
                        code = need_match.group(1).strip()
                        presence = f"C ({code})" if "Cond" in code else f"O ({code})"
                    elif "OPTIONAL" in tail.upper():
                        presence = "O"
                    else:
                        presence = "M" if parent_kind == "SEQUENCE" else "O"

                    child_kind = "SEQUENCE" if rest.startswith("SEQUENCE") else "CHOICE"
                    nested_fields = self._parse_sequence_fields(inner_body, child_kind)

                    fields.append({
                        "name": fname,
                        "type": raw_type_body,
                        "presence": presence,
                        "format": parent_kind,
                        "raw_field": tok,
                        "nested_fields": nested_fields,
                        "iei": "",
                    })
                    continue

            # Standard simple or parameterized fields
            need_match = RE_COND_NEED.search(rest)
            if need_match:
                code = need_match.group(1).strip()
                presence = f"C ({code})" if "Cond" in code else f"O ({code})"
            elif "OPTIONAL" in rest.upper():
                presence = "O"
            else:
                presence = "M" if parent_kind == "SEQUENCE" else "O"

            # Strip outer keywords from simple type
            clean_type_candidate = RE_STRIP_KEYWORDS.sub("", rest.split("--")[0]).strip()

            fields.append({
                "name": fname,
                "type": clean_type_candidate,
                "presence": presence,
                "format": parent_kind,
                "raw_field": tok,
                "nested_fields": [],
                "iei": "",
            })

        return fields

    def _split_asn1_tokens(self, text: str) -> List[str]:
        """Splits comma-separated ASN.1 fields at depth 0, attaching trailing comments to the field token."""
        tokens: List[str] = []
        current: List[str] = []
        depth = p_depth = b_depth = 0
        in_comment = False
        i = 0
        n = len(text)

        while i < n:
            c = text[i]

            if not in_comment and c == '-' and i + 1 < n and text[i+1] == '-':
                in_comment = True
                current.append(c)
                current.append(text[i+1])
                i += 2
                continue
            elif in_comment:
                current.append(c)
                if c == '\n':
                    in_comment = False
                elif c == '-' and i + 1 < n and text[i+1] == '-':
                    current.append(text[i+1])
                    i += 2
                    in_comment = False
                    continue
                i += 1
                continue

            if c == '{':
                depth += 1
            elif c == '}':
                depth = max(0, depth - 1)
            elif c == '(':
                p_depth += 1
            elif c == ')':
                p_depth = max(0, p_depth - 1)
            elif c == '[':
                b_depth += 1
            elif c == ']':
                b_depth = max(0, b_depth - 1)

            if c == ',' and depth == 0 and p_depth == 0 and b_depth == 0:
                current.append(c)
                # Consume any trailing whitespace and single-line comment on the same line
                j = i + 1
                while j < n and text[j] in (' ', '\t'):
                    current.append(text[j])
                    j += 1
                if j + 1 < n and text[j] == '-' and text[j+1] == '-':
                    while j < n and text[j] != '\n':
                        current.append(text[j])
                        j += 1
                tokens.append("".join(current).strip())
                current = []
                i = j
                continue

            current.append(c)
            i += 1

        if current:
            tail = "".join(current).strip()
            if tail:
                tokens.append(tail)

        return tokens

    def _build_inspector_html(
            self,
            type_name: str,
            assigned_clause: str,
            raw_asn1: str,
            desc_table_html: str,
    ) -> str:
        """Constructs HTML presentation combining ASN.1 definitions and 3GPP tabular specifications."""
        clause_str = f" (Clause {assigned_clause})" if assigned_clause and assigned_clause != type_name else ""

        header_html = (
            f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 8px; '
            f'font-family: Segoe UI, sans-serif; font-size: 15px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
            f'{type_name}{clause_str}</h2>'
        )

        raw_asn1_html = ""
        if raw_asn1 and raw_asn1.strip():
            raw_asn1_html = (
                f'<div style="margin-bottom: 10px;">'
                f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 3px;">ASN.1 Definition:</div>'
                f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; '
                f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A; '
                f'white-space: pre-wrap; margin: 0;">{html.escape(raw_asn1)}</pre>'
                f'</div>'
            )

        spec_section_html = ""
        if desc_table_html and desc_table_html.strip():
            spec_section_html = (
                f'<div style="margin-top: 8px;">'
                f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 4px;">'
                f'Specification Details & Tabular Definition:</div>'
                f'{desc_table_html}'
                f'</div>'
            )

        return f'{header_html}{raw_asn1_html}{spec_section_html}'