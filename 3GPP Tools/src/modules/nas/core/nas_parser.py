import html
import json
import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Tuple, Union, Set
import zipfile

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.specifications.utils.utils import file_version_to_version

# WordprocessingML XML Namespaces
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TAG_BODY = f"{W_NS}body"
TAG_P = f"{W_NS}p"
TAG_TBL = f"{W_NS}tbl"
TAG_TR = f"{W_NS}tr"
TAG_TC = f"{W_NS}tc"
TAG_T = f"{W_NS}t"
TAG_TAB = f"{W_NS}tab"
TAG_BR = f"{W_NS}br"
TAG_CR = f"{W_NS}cr"
TAG_HYPHEN = f"{W_NS}noBreakHyphen"
TAG_TCPR = f"{W_NS}tcPr"
TAG_GRIDSPAN = f"{W_NS}gridSpan"
TAG_VMERGE = f"{W_NS}vMerge"
TAG_PPR = f"{W_NS}pPr"
TAG_JC = f"{W_NS}jc"


def _extract_p_text(p_elem) -> str:
    """Extracts clean text from a <w:p> node, preserving spaces, tabs, and line breaks."""
    text_pieces = []
    for node in p_elem.iter():
        tag = node.tag
        if tag == TAG_T:
            if node.text:
                text_pieces.append(node.text)
        elif tag == TAG_TAB:
            text_pieces.append(" ")
        elif tag in (TAG_BR, TAG_CR):
            text_pieces.append(" ")
        elif tag == TAG_HYPHEN:
            text_pieces.append("-")

    raw = "".join(text_pieces)
    raw = raw.replace("\u00a0", " ").replace("\xa0", " ")
    return " ".join(raw.split())


def _extract_tc_text(tc_elem) -> str:
    """Extracts text from a table cell <w:tc>, joining multiple paragraphs with spaces."""
    p_texts = []
    for p in tc_elem.findall(TAG_P):
        pt = _extract_p_text(p)
        if pt:
            p_texts.append(pt)
    return " ".join(p_texts).strip()


def _convert_table_to_html(tbl_elem, is_figure_diagram: bool = False) -> str:
    """Converts a Word XML table into a styled HTML table supporting colspan and vertical alignment."""
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return ""

    table_style = (
        "border-collapse: collapse; margin: 8px 0; border: 1px solid #CBD5E1; "
        "font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; width: 100%;"
    )
    html_parts = [f'<table border="1" cellspacing="0" cellpadding="4" style="{table_style}">']

    for r_idx, row in enumerate(rows):
        html_parts.append("<tr>")
        cells = row.findall(TAG_TC)

        for cell in cells:
            tcPr = cell.find(TAG_TCPR)
            colspan = 1
            is_vmerge_continue = False

            if tcPr is not None:
                gs = tcPr.find(TAG_GRIDSPAN)
                if gs is not None:
                    val = gs.get(f"{W_NS}val")
                    if val and val.isdigit():
                        colspan = int(val)

                vm = tcPr.find(TAG_VMERGE)
                if vm is not None:
                    val = vm.get(f"{W_NS}val")
                    if val != "restart":
                        is_vmerge_continue = True

            if is_vmerge_continue:
                continue

            cell_text = html.escape(_extract_tc_text(cell))
            tag = "th" if r_idx == 0 else "td"
            style_bits = ["border: 1px solid #E2E8F0;", "padding: 4px 6px;"]

            if r_idx == 0:
                style_bits.append("background-color: #F1F5F9; font-weight: bold; color: #1E293B;")

            colspan_attr = f' colspan="{colspan}"' if colspan > 1 else ""
            style_str = " ".join(style_bits)
            html_parts.append(f'<{tag}{colspan_attr} style="{style_str}">{cell_text}</{tag}>')

        html_parts.append("</tr>")

    html_parts.append("</table>")
    return "".join(html_parts)


# =========================================================================
# --- RRC / NGAP ASN.1 EXTRACTOR & PARSER ---
# =========================================================================
class RRCAsn1DocxParser:
    """
    Extracts ASN.1 Message definitions, Information Elements, Sequence Fields,
    and accompanying Field Description tables from 3GPP TS 38.331 / TS 36.331 / TS 38.413.
    """

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.331"):
        self.docx_paths = sorted(docx_paths, key=lambda p: self._extract_part_index(p.name))
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(filename: str) -> int:
        match = re.search(r"_(\d+)_", filename)
        return int(match.group(1)) if match else 0

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        raw_asn1_blocks: List[str] = []
        field_desc_tables: Dict[str, str] = {}  # {ie_or_msg_name: html_table}
        total_files = len(self.docx_paths)

        # 1. First pass: Collect all ASN.1 text and Field Description HTML tables
        for file_idx, docx_path in enumerate(self.docx_paths):
            if progress_callback:
                progress_callback(f"Scanning {docx_path.name} ({file_idx + 1}/{total_files})...", 10 + int(file_idx / total_files * 40))

            try:
                with zipfile.ZipFile(docx_path, "r") as zf:
                    if "word/document.xml" not in zf.namelist():
                        continue
                    xml_bytes = zf.read("word/document.xml")
            except Exception as e:
                self.logger.warning(f"Failed to read {docx_path.name}: {e}")
                continue

            root = ET.fromstring(xml_bytes)
            body = root.find(TAG_BODY)
            if body is None:
                continue

            in_asn1 = False
            current_asn1_lines = []
            last_p_text = ""

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_p_text = p_text

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
                    # Check for Field Descriptions table (e.g., 'RRCReconfiguration field descriptions')
                    match_tbl = re.search(r"([A-Za-z0-9\-_]+)\s+field\s+descriptions", last_p_text, re.IGNORECASE)
                    if match_tbl:
                        target_name = match_tbl.group(1).strip()
                        field_desc_tables[target_name.lower()] = _convert_table_to_html(elem)

        # Combine all raw ASN.1 into a unified module string
        full_asn1_module = "\n".join(raw_asn1_blocks)
        if not full_asn1_module.strip():
            # Fallback for specs that do not use explicit ASN1START markers
            full_asn1_module = self._fallback_extract_asn1()

        if progress_callback:
            progress_callback("Parsing ASN.1 structures and building evolution records...", 60)

        # 2. Extract Type Definitions from ASN.1
        type_defs = self._extract_asn1_type_definitions(full_asn1_module)

        # 3. Classify Messages vs. Information Elements
        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        # Identify top-level RRC Messages / SIBs
        is_rrc = "331" in self.spec_number
        is_ngap = "413" in self.spec_number

        for type_name, type_info in type_defs.items():
            is_message = False
            clean_name = type_name.strip()

            if is_rrc:
                # Common RRC PDU / SIB naming patterns
                if (
                    clean_name.startswith("RRC")
                    or clean_name.startswith("SIB")
                    or clean_name.startswith("SystemInformation")
                    or clean_name.endswith("Request")
                    or clean_name.endswith("Response")
                    or clean_name.endswith("Command")
                    or clean_name.endswith("Complete")
                    or clean_name.endswith("Failure")
                ) and not clean_name.endswith("-IEs") and not clean_name.endswith("-v1530-IEs"):
                    is_message = True

            elif is_ngap:
                if clean_name.endswith("Request") or clean_name.endswith("Response") or clean_name.endswith("Acknowledge") or clean_name.endswith("UnsuccessfulOutcome"):
                    is_message = True

            # Prepare HTML definition for Inspector
            raw_asn1_html = f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A;">{html.escape(type_info["raw_asn1"])}</pre>'
            desc_table_html = field_desc_tables.get(clean_name.lower(), "")
            if desc_table_html:
                desc_table_html = f'<h3 style="font-size: 12px; font-weight: bold; color: #1E293B; margin-top: 10px; margin-bottom: 4px;">{clean_name} field descriptions</h3>' + desc_table_html

            full_inspector_html = (
                f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 6px; font-family: Segoe UI, sans-serif;">{clean_name} (Clause {type_info.get("clause", "6.2/6.3")})</h2>'
                + raw_asn1_html
                + desc_table_html
            )

            ie_definitions.append({
                "clause": type_info.get("clause", clean_name),
                "ie_name": clean_name,
                "raw_description": full_inspector_html,
                "structure_table": json.dumps([]),
            })

            # If it qualifies as a message / PDU, unroll its fields for the Evolution Matrix
            if is_message:
                unrolled_fields = self._unroll_message_fields(clean_name, type_defs)
                if unrolled_fields:
                    messages.append({
                        "clause": type_info.get("clause", "6.2"),
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} ASN.1 PDU Definition",
                        "ies": unrolled_fields,
                    })

        return messages, ie_definitions

    def _fallback_extract_asn1(self) -> str:
        """Collects paragraphs matching ASN.1 syntax when ASN1START markers are omitted."""
        collected = []
        for docx_path in self.docx_paths:
            try:
                with zipfile.ZipFile(docx_path, "r") as zf:
                    if "word/document.xml" not in zf.namelist():
                        continue
                    root = ET.fromstring(zf.read("word/document.xml"))
                    body = root.find(TAG_BODY)
                    if body is not None:
                        for p in body.findall(TAG_P):
                            t = _extract_p_text(p)
                            if "::=" in t or "SEQUENCE" in t or "CHOICE" in t or "OPTIONAL" in t:
                                collected.append(t)
            except Exception:
                pass
        return "\n".join(collected)

    def _extract_asn1_type_definitions(self, asn1_text: str) -> Dict[str, Dict[str, Any]]:
        """Parses raw ASN.1 text into structured type definitions."""
        type_defs: Dict[str, Dict[str, Any]] = {}
        # Match type declarations like 'TypeName ::= SEQUENCE { ... }'
        pattern = re.compile(
            r"([A-Za-z0-9\-]+)\s*::=\s*(SEQUENCE|CHOICE|ENUMERATED|BIT STRING|OCTET STRING|INTEGER)?\s*(\{[^}]*\})?",
            re.MULTILINE | re.DOTALL,
        )

        for match in pattern.finditer(asn1_text):
            type_name = match.group(1).strip()
            type_kind = match.group(2) or "TYPE"
            body = match.group(3) or ""
            raw_block = match.group(0).strip()

            fields = []
            if body and body.startswith("{") and body.endswith("}"):
                inner = body[1:-1].strip()
                # Split comma-separated field definitions
                raw_fields = re.split(r",\s*(?![^{]*\})", inner)
                for rf in raw_fields:
                    f_clean = " ".join(rf.split()).strip()
                    if not f_clean or f_clean.startswith("--"):
                        continue
                    # Match 'fieldName FieldType [OPTIONAL | DEFAULT ...]'
                    f_match = re.match(r"([A-Za-z0-9\-]+)\s+([A-Za-z0-9\-\(\)\s]+?)(?:\s+(OPTIONAL|MANDATORY|DEFAULT\s+[^,\s]+))?(?:--.*)?$", f_clean)
                    if f_match:
                        fname = f_match.group(1).strip()
                        ftype = f_match.group(2).strip()
                        pres = "OPTIONAL" if f_match.group(3) and "OPTIONAL" in f_match.group(3) else ("M" if not f_match.group(3) else "DEFAULT")
                        fields.append({
                            "name": fname,
                            "type": ftype,
                            "presence": pres,
                            "format": type_kind,
                        })

            type_defs[type_name] = {
                "kind": type_kind,
                "fields": fields,
                "raw_asn1": raw_block,
                "clause": type_name,
            }

        return type_defs

    def _unroll_message_fields(
        self, root_name: str, type_defs: Dict[str, Dict[str, Any]], max_depth: int = 3
    ) -> List[Dict[str, Any]]:
        """Recursively unrolls nested SEQUENCE / CHOICE fields into an indexed hierarchy for the Evolution Matrix."""
        result: List[Dict[str, Any]] = []

        def recurse(current_type: str, path_prefix: str, depth: int, visited: Set[str]):
            if depth > max_depth or current_type in visited:
                return

            t_info = type_defs.get(current_type)
            if not t_info or not t_info["fields"]:
                return

            new_visited = visited | {current_type}

            for f in t_info["fields"]:
                fname = f["name"]
                ftype = f["type"]
                presence = f["presence"]
                fmt = f["format"]

                full_path = f"{path_prefix}.{fname}" if path_prefix else fname
                clean_type_lookup = re.sub(r"[\(\)\{\}].*$", "", ftype).strip()

                result.append({
                    "iei": "",
                    "information_element": fname,
                    "field_path": full_path,
                    "depth": depth,
                    "type_reference": ftype,
                    "presence": "O" if presence == "OPTIONAL" else ("M" if presence == "M" else "O"),
                    "format": fmt,
                    "length": "-",
                })

                # Recursively expand nested sequences / critical extensions
                if clean_type_lookup in type_defs and depth < max_depth:
                    recurse(clean_type_lookup, full_path, depth + 1, new_visited)

        recurse(root_name, "", 0, set())

        # If direct unrolling returned empty (e.g. root wraps criticalExtensions), look for `<root>-IEs`
        if not result and f"{root_name}-IEs" in type_defs:
            recurse(f"{root_name}-IEs", "", 0, set())

        return result


# =========================================================================
# --- UNIFIED PARSER DISPATCHER ---
# =========================================================================
class NASDocxParser:
    """Direct XML parser for 3GPP Specifications, supporting both NAS and RRC/NGAP ASN.1."""

    def __init__(self, docx_paths: Union[Path, str, List[Union[Path, str]]]):
        if isinstance(docx_paths, (str, Path)):
            self.docx_paths = [Path(docx_paths)]
        else:
            self.docx_paths = [Path(p) for p in docx_paths]

        self.docx_paths.sort(key=lambda p: self._extract_part_index(p.name))
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(filename: str) -> int:
        match = re.search(r"_(\d+)_", filename)
        return int(match.group(1)) if match else 0

    def extract_spec_number(self) -> str:
        if not self.docx_paths:
            return "24.501"
        stem = self.docx_paths[0].stem.replace(".", "").replace("-", "").replace("_", "")
        for known in ["38331", "36331", "38413", "24301", "24501"]:
            if known in stem:
                return f"{known[:2]}.{known[2:]}"
        match = re.search(r"(24|36|38)[._]?(301|501|331|413)", self.docx_paths[0].stem)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        return "24.501"

    def extract_version_from_filename(self) -> str:
        if not self.docx_paths:
            return ""
        stem = self.docx_paths[0].stem
        match = re.search(r"-([a-zA-Z0-9]{3})(?:_\d+.*)?$", stem)
        if match:
            parsed_ver = file_version_to_version(match.group(1))
            if parsed_ver:
                return parsed_ver
        return stem

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        spec_num = self.extract_spec_number()

        # Route ASN.1 specifications to RRCAsn1DocxParser
        if any(s in spec_num for s in ["38.331", "36.331", "38.413"]):
            rrc_parser = RRCAsn1DocxParser(self.docx_paths, spec_number=spec_num)
            return rrc_parser.parse(progress_callback=progress_callback)

        # Standard NAS Clause 8/9 Parser
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        caption_pattern = re.compile(
            r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
            re.IGNORECASE,
        )
        ie_heading_pattern = re.compile(
            r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$"
        )
        major_boundary_pattern = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")

        total_files = len(valid_paths)

        for file_idx, docx_path in enumerate(valid_paths):
            file_weight = 1.0 / total_files
            base_file_progress = 10 + int((file_idx / total_files) * 80)

            if progress_callback:
                progress_callback(f"Reading {docx_path.name} ({file_idx + 1}/{total_files})...", base_file_progress)

            try:
                with zipfile.ZipFile(docx_path, "r") as zf:
                    if "word/document.xml" not in zf.namelist():
                        continue
                    xml_bytes = zf.read("word/document.xml")
            except Exception as e:
                self.logger.warning(f"Failed to read word/document.xml from {docx_path.name}: {e}")
                continue

            root = ET.fromstring(xml_bytes)
            body = root.find(TAG_BODY)
            if body is None:
                continue

            last_caption_info: Optional[Tuple[str, str, str]] = None
            last_paragraph_text: str = ""
            current_ie_def: Optional[Dict[str, Any]] = None

            body_elements = list(body)
            total_elements = len(body_elements)

            for idx, elem in enumerate(body_elements):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if p_text:
                        last_paragraph_text = p_text

                        if p_text.startswith("Table 8.") or p_text.startswith("Table D."):
                            match_cap = caption_pattern.search(p_text)
                            if match_cap:
                                clause = match_cap.group(1).strip()
                                name = match_cap.group(2).strip()
                                name = re.sub(r"(?i)\s+message\s+content", "", name).strip()
                                last_caption_info = (clause, name, p_text)
                        else:
                            last_caption_info = None

                        match_ie = ie_heading_pattern.match(p_text)
                        if match_ie:
                            if current_ie_def:
                                self._finalize_ie_def(current_ie_def, ie_definitions)

                            cl = match_ie.group(1).strip()
                            ie_name = match_ie.group(2).strip()
                            current_ie_def = {
                                "clause": cl,
                                "ie_name": ie_name,
                                "html_content": [
                                    f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; font-family: Segoe UI, sans-serif;">{ie_name} (Clause {cl})</h2>'
                                ],
                            }
                        elif major_boundary_pattern.match(p_text) and not p_text.startswith("9."):
                            if current_ie_def:
                                self._finalize_ie_def(current_ie_def, ie_definitions)
                                current_ie_def = None
                        elif current_ie_def:
                            if p_text.startswith("Figure 9.") or p_text.startswith("Figure D."):
                                current_ie_def["html_content"].append(
                                    f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{p_text}</p>'
                                )
                            elif p_text.startswith("Table 9.") or p_text.startswith("Table D."):
                                current_ie_def["html_content"].append(
                                    f'<p style="font-weight: bold; color: #1A237E; margin-top: 12px; margin-bottom: 4px;">{p_text}</p>'
                                )
                            elif p_text.startswith("NOTE"):
                                current_ie_def["html_content"].append(
                                    f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{p_text}</div>'
                                )
                            else:
                                current_ie_def["html_content"].append(
                                    f'<p style="margin: 4px 0; line-height: 1.4; color: #263238;">{p_text}</p>'
                                )

                elif elem.tag == TAG_TBL:
                    if last_caption_info:
                        clause, name, full_caption = last_caption_info
                        ies = self._parse_clause_8_table(elem)
                        if ies:
                            messages.append({
                                "clause": clause,
                                "message_name": name,
                                "table_caption": full_caption,
                                "ies": ies,
                            })
                        last_caption_info = None

                    elif current_ie_def:
                        is_diagram = "figure" in last_paragraph_text.lower() or any(
                            "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC)
                        )
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_diagram)
                        if tbl_html:
                            current_ie_def["html_content"].append(tbl_html)

            if current_ie_def:
                self._finalize_ie_def(current_ie_def, ie_definitions)

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} messages and {len(ie_definitions)} definitions.", 95)

        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        html_str = "".join(current_ie_def["html_content"])
        ie_definitions.append({
            "clause": current_ie_def["clause"],
            "ie_name": current_ie_def["ie_name"],
            "raw_description": html_str,
            "structure_table": json.dumps([]),
        })

    def _parse_clause_8_table(self, tbl_elem) -> List[Dict[str, str]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_row_idx = -1
        for r_idx in range(min(3, len(rows))):
            cells_text = [_extract_tc_text(tc).lower() for tc in rows[r_idx].findall(TAG_TC)]
            joined_header = " ".join(cells_text)
            if "information element" in joined_header or "iei" in cells_text:
                header_row_idx = r_idx
                break

        if header_row_idx == -1:
            return []

        ies = []
        for row in rows[header_row_idx + 1:]:
            cells = [_extract_tc_text(tc) for tc in row.findall(TAG_TC)]
            if len(cells) < 6:
                continue

            iei = cells[0]
            ie_name = cells[1]
            type_ref = cells[2]
            presence = cells[3]
            fmt = cells[4]
            length = cells[5]

            clean_name = ie_name.lower().replace(" ", "")
            if not clean_name or "informationelement" in clean_name:
                continue

            ies.append({
                "iei": iei,
                "information_element": ie_name,
                "field_path": ie_name,
                "depth": 0,
                "type_reference": type_ref,
                "presence": presence,
                "format": fmt,
                "length": length,
            })

        return ies