import html
import json
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY, TAG_P, TAG_TBL, TAG_TR, TAG_TC
)
from modules.nas.core.parsing.protocol_parser_utils import (
    extract_document_root, _extract_p_text, _extract_tc_text, _convert_table_to_html
)

RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_PFCP_CLAUSE_HEADER = re.compile(r"^((?:7|8)(?:\.[0-9A-Za-z]+)+)\s*(.*)$")
RE_PFCP_TABLE_CAPTION = re.compile(r"^Table\s+([78]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+)$", re.IGNORECASE)

# Canonical 3GPP TS 29.244 top-level message clauses
TOP_LEVEL_PFCP_CLAUSES = {
    "7.4.2.1", "7.4.2.2", "7.4.3.1", "7.4.3.2", "7.4.4.1", "7.4.4.2", "7.4.4.3",
    "7.4.4.4", "7.4.4.5", "7.4.4.6", "7.4.4.7", "7.4.5.1.1", "7.4.5.2.1",
    "7.4.6.1", "7.4.6.2", "7.4.7.1", "7.4.7.2", "7.5.2.1", "7.5.3.1", "7.5.4.1",
    "7.5.5.1", "7.5.6", "7.5.7.1", "7.5.8.1", "7.5.9.1"
}

# Known scalar identifier fields that should never be recursively expanded as containers
SCALAR_ID_FIELDS = {
    "pdrid", "farid", "urrid", "qerid", "barid", "srrid", "marid", "nodeid",
    "trafficendpointid", "groupid", "failedruleid", "mbsunicastparametersid",
    "headerhandlingcontrolruleid", "headerhandlingcontrolid", "reportingendpointid",
    "n6delaymeasurementcontrolinformationid"
}


class PFCPDocxParser:
    """Dedicated parser for 3GPP TS 29.244 (PFCP) specifications."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "29.244"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    @staticmethod
    def _normalize_key(name: str) -> str:
        return re.sub(r"[^a-zA-Z0-9]", "", name).lower()

    @staticmethod
    def _clean_clause_number(clause_ref: str) -> str:
        return clause_ref.split("-")[0].strip()

    @staticmethod
    def _get_tc_grid_span(tc_elem: Any) -> int:
        for elem in tc_elem.iter():
            if elem.tag.endswith("gridSpan"):
                for attr_name, attr_val in elem.attrib.items():
                    if attr_name.endswith("val") and attr_val.isdigit():
                        return int(attr_val)
        return 1

    def _expand_row_cells(self, row: Any) -> List[str]:
        expanded: List[str] = []
        for tc in row.findall(TAG_TC):
            text = _extract_tc_text(tc).strip()
            span = self._get_tc_grid_span(tc)
            expanded.extend([text] * span)
        return expanded

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            self.logger.error("No valid specification docx files found: %s", self.docx_paths)
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        total_files = len(valid_paths)
        self.logger.info("Starting PFCP ingestion for TS %s across %d document part(s)", self.spec_number, total_files)

        raw_tables: List[Tuple[str, str, str, Any]] = []
        ie_definitions_dict: Dict[str, Dict[str, Any]] = {}
        type_registry: Dict[str, Dict[str, str]] = {}

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(f"Scanning PFCP {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

            self.logger.info("Parsing document part %d/%d: %s", file_idx + 1, total_files, docx_path.name)
            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            last_caption_info: Optional[Tuple[str, str, str]] = None
            current_ie_clause = ""
            current_ie_name = ""
            current_ie_html: List[str] = []

            def _finalize_current_ie():
                nonlocal current_ie_clause, current_ie_name, current_ie_html
                if current_ie_clause and current_ie_name and current_ie_html:
                    norm_k = self._normalize_key(current_ie_name)
                    if norm_k not in ie_definitions_dict:
                        ie_definitions_dict[norm_k] = {
                            "clause": current_ie_clause,
                            "name": current_ie_name,
                            "ie_name": current_ie_name,
                            "raw_description": "".join(current_ie_html),
                            "structure_table": json.dumps([]),
                        }
                current_ie_clause = ""
                current_ie_name = ""
                current_ie_html = []

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue

                    match_header = RE_PFCP_CLAUSE_HEADER.match(p_text)
                    if match_header:
                        cl_num = match_header.group(1).strip()
                        cl_title = match_header.group(2).strip()

                        if cl_num.startswith("8.2.") and cl_title:
                            _finalize_current_ie()
                            current_ie_clause = cl_num
                            current_ie_name = cl_title
                            current_ie_html = [
                                f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                                f'{html.escape(cl_title)} (Clause {cl_num})</h2>'
                            ]
                        elif not cl_num.startswith("8.2."):
                            _finalize_current_ie()

                    match_cap = RE_PFCP_TABLE_CAPTION.match(p_text)
                    if match_cap:
                        clause_ref = match_cap.group(1).strip()
                        cap_name = match_cap.group(2).strip()
                        last_caption_info = (clause_ref, cap_name, p_text)
                    elif not p_text.startswith("Table "):
                        last_caption_info = None

                    if current_ie_clause:
                        if p_text.startswith("Figure 8."):
                            current_ie_html.append(
                                f'<p style="font-weight: bold; color: #475569; margin-top: 8px; margin-bottom: 3px; font-size: 11px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith("Table 8."):
                            current_ie_html.append(
                                f'<p style="font-weight: bold; color: #0369A1; margin-top: 10px; margin-bottom: 3px; font-size: 11px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_html.append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">'
                                f'{html.escape(p_text)}</div>'
                            )
                        elif not p_text.startswith(("Figure", "Table")) and len(current_ie_html) < 20:
                            current_ie_html.append(
                                f'<p style="margin: 3px 0; line-height: 1.4; color: #334155; font-size: 11px;">{html.escape(p_text)}</p>'
                            )

                elif elem.tag == TAG_TBL:
                    if last_caption_info:
                        clause_ref, cap_name, full_cap = last_caption_info
                        raw_tables.append((clause_ref, cap_name, full_cap, elem))
                        last_caption_info = None

                    if current_ie_clause:
                        is_bit_diagram = any("bit" in _extract_tc_text(c).lower() or "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC))
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_bit_diagram)
                        if tbl_html:
                            current_ie_html.append(tbl_html)

            _finalize_current_ie()

        self.logger.info("Scanned XML structures: Found %d raw captioned tables and %d IE definitions in Clause 8.2",
                         len(raw_tables), len(ie_definitions_dict))

        if progress_callback:
            progress_callback("Analyzing Table 8.1.2-1 Type registry and Clause 7 tables...", 60)

        # 1. Parse Table 8.1.2-1 (IE Type Master Registry)
        for clause_ref, cap_name, _, tbl_elem in raw_tables:
            if "8.1.2-1" in clause_ref or "information element types" in cap_name.lower():
                self._parse_type_registry(tbl_elem, type_registry)
                self.logger.info("Parsed Table 8.1.2-1 type registry: Loaded %d IE types", len(type_registry))
                break

        # 2. Parse Clause 7 tables into messages and grouped IEs
        top_messages_raw: List[Dict[str, Any]] = []
        grouped_ies: Dict[str, Dict[str, Any]] = {}

        for clause_ref, cap_name, full_cap, tbl_elem in raw_tables:
            if not clause_ref.startswith("7."):
                continue

            parsed_ies = self._parse_clause_7_table(tbl_elem, type_registry, clause_ref, cap_name)
            if not parsed_ies:
                continue

            clean_name = self._sanitize_table_name(cap_name)
            base_clause = self._clean_clause_number(clause_ref)
            is_top_msg = self._is_top_level_message(clean_name, clause_ref, base_clause)

            table_record = {
                "clause": base_clause,
                "clause_ref": clause_ref,
                "name": clean_name,
                "message_name": clean_name,
                "table_caption": full_cap,
                "ies": parsed_ies,
                "fields": parsed_ies,
            }

            if is_top_msg:
                top_messages_raw.append(table_record)
                self.logger.info("Found top-level message [%s] (%s) with %d direct IEs", base_clause, clean_name, len(parsed_ies))
            else:
                norm_key = self._normalize_key(clean_name)
                grouped_ies[norm_key] = table_record
                grouped_ies[clause_ref] = table_record
                if base_clause not in grouped_ies:
                    grouped_ies[base_clause] = table_record
                self.logger.debug("Registered grouped IE container [%s] -> %s (%d fields)", clause_ref, norm_key, len(parsed_ies))

        self.logger.info("Clause 7 parsing complete: Identified %d top-level messages and %d grouped IE groups",
                         len(top_messages_raw), len(grouped_ies))

        if progress_callback:
            progress_callback("Unrolling nested Grouped Information Elements across PFCP messages...", 80)

        # 3. Unroll Grouped IEs hierarchically into top-level messages
        final_messages: List[Dict[str, Any]] = []
        for msg in top_messages_raw:
            unrolled_ies = self._unroll_pfcp_message_ies(msg["ies"], grouped_ies)
            self.logger.info("Unrolled message '%s': %d total fields (original direct: %d)",
                            msg["name"], len(unrolled_ies), len(msg["ies"]))
            final_messages.append({
                "clause": msg["clause"],
                "name": msg["name"],
                "message_name": msg["message_name"],
                "table_caption": msg["table_caption"],
                "ies": unrolled_ies,
                "fields": unrolled_ies,
            })

        ie_definitions = list(ie_definitions_dict.values())
        self.logger.info("Completed parsing TS %s: Returning %d final messages and %d IE definitions",
                         self.spec_number, len(final_messages), len(ie_definitions))
        return final_messages, ie_definitions

    def _parse_type_registry(self, tbl_elem: Any, registry: Dict[str, Dict[str, str]]):
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return

        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) >= 2:
                type_val = cells[0]
                ie_name = cells[1]
                clause_ref = cells[2] if len(cells) > 2 else ""
                appl = cells[3] if len(cells) > 3 else ""

                if type_val.isdigit() and ie_name:
                    norm_k = self._normalize_key(ie_name)
                    registry[norm_k] = {
                        "type_id": type_val,
                        "clause": clause_ref,
                        "applicability": appl,
                    }

    def _sanitize_table_name(self, caption: str) -> str:
        s = caption.strip().lstrip(": \t")
        # Strip leading "Information Elements in (a/an/the)?"
        s = re.sub(r"^Information\s+Elements\s+in\s+(?:an?\s+|the\s+)?", "", s, flags=re.IGNORECASE)
        # Strip trailing sub-table qualifiers
        s = re.sub(r"\s+IE\s+(?:in|within)\s+(?:the\s+)?.+$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+(?:in|within)\s+(?:the\s+)?.+$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+message\s+content$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+message$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+IE$", "", s, flags=re.IGNORECASE)
        # Strip remaining leading article "a " or "an " or "the "
        s = re.sub(r"^(?:an?|the)\s+", "", s, flags=re.IGNORECASE)
        return s.strip().lstrip(": \t")

    @staticmethod
    def _is_top_level_message(clean_name: str, clause_ref: str, base_clause: str) -> bool:
        if not clause_ref.endswith("-1"):
            return False

        lower_name = clean_name.lower()
        if any(lower_name.endswith(kw) for kw in ("request", "response", "reject", "indication")):
            return True
        if base_clause in TOP_LEVEL_PFCP_CLAUSES and not any(
            k in lower_name for k in ("information", "parameter", "filter", "rule", "pdr", "far", "urr", "qer", "bar", "srr", "mar")
        ):
            return True
        return False

    def _parse_clause_7_table(
        self,
        tbl_elem: Any,
        type_registry: Dict[str, Dict[str, str]],
        clause_ref: str = "",
        cap_name: str = "",
    ) -> List[Dict[str, Any]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_idx = -1
        name_col = -1
        pres_col = -1
        comment_col = -1
        type_col = -1
        appl_cols: Dict[int, str] = {}

        # Step 1: Locate the primary header row using grid-expanded cells
        for r_idx in range(min(4, len(rows))):
            expanded = self._expand_row_cells(rows[r_idx])
            joined = " ".join(c.lower() for c in expanded)
            if "information element" in joined:
                header_idx = r_idx
                for idx, c in enumerate(expanded):
                    c_low = c.lower()
                    if "information element" in c_low and name_col == -1:
                        name_col = idx
                    elif c_low in ("p", "presence") and pres_col == -1:
                        pres_col = idx
                    elif any(k in c_low for k in ("condition", "comment")) and comment_col == -1:
                        comment_col = idx
                    elif any(k in c_low for k in ("ie type", "type", "reference")) and type_col == -1:
                        type_col = idx
                    elif "clause" in c_low and type_col == -1:
                        type_col = idx
                break

        if header_idx == -1 or name_col == -1:
            return []

        # Step 2: Detect applicability subheader row (e.g. Sxa | Sxb | Sxc | N4 | N4mb)
        data_start_idx = header_idx + 1
        if header_idx + 1 < len(rows):
            next_expanded = self._expand_row_cells(rows[header_idx + 1])
            sub_ifaces: Dict[int, str] = {}
            for idx, c in enumerate(next_expanded):
                c_clean = c.strip()
                if c_clean.lower() in ("sxa", "sxb", "sxc", "sxa/sxb", "sxb'", "sxa'", "n4", "n4mb"):
                    sub_ifaces[idx] = c_clean
            if sub_ifaces:
                appl_cols = sub_ifaces
                data_start_idx = header_idx + 2

        # Step 3: IE Type / Reference is the final column in 3GPP Clause 7 tables
        header_width = len(self._expand_row_cells(rows[header_idx]))
        if type_col == -1 or type_col in appl_cols or (appl_cols and type_col < max(appl_cols.keys())):
            type_col = header_width - 1

        ies: List[Dict[str, Any]] = []
        for r_offset in range(data_start_idx, len(rows)):
            row = rows[r_offset]
            raw_cells = self._expand_row_cells(row)
            if len(raw_cells) <= name_col:
                continue

            ie_name = raw_cells[name_col].strip()

            is_note = (
                ie_name.upper().startswith("NOTE")
                or any(c.strip().upper().startswith("NOTE") for c in raw_cells if c.strip())
            )
            is_header_kw = any(kw in ie_name.lower() for kw in ("information element", "octet 1", "octets"))
            is_filler_text = ie_name.lower().startswith("same ies and requirements")

            if not ie_name or is_header_kw or is_note or is_filler_text:
                continue

            presence = raw_cells[pres_col].strip() if pres_col != -1 and len(raw_cells) > pres_col else "O"
            if presence:
                presence = presence.split()[0]
            if not presence:
                presence = "O"

            type_ref = raw_cells[type_col].strip() if type_col != -1 and len(raw_cells) > type_col else ""

            if appl_cols:
                matched_ifaces = [
                    iface for c_idx, iface in appl_cols.items()
                    if len(raw_cells) > c_idx and raw_cells[c_idx].strip().upper() == "X"
                ]
                appl = ", ".join(matched_ifaces)
            else:
                appl = raw_cells[comment_col + 1].strip() if comment_col != -1 and len(raw_cells) > comment_col + 1 and (comment_col + 1) != type_col else ""

            norm_name = self._normalize_key(ie_name)
            reg_info = type_registry.get(norm_name, {})
            iei_val = reg_info.get("type_id", "")

            if not appl and reg_info.get("applicability"):
                appl = reg_info["applicability"]

            if not type_ref or type_ref == ie_name or type_ref in ("-", "X"):
                if reg_info.get("clause"):
                    type_ref = f"{ie_name} ({reg_info['clause']})"
                else:
                    type_ref = ie_name

            is_grouped = (
                "within" in type_ref.lower()
                or "create" in ie_name.lower()
                or "update" in ie_name.lower()
                or "information elements in" in type_ref.lower()
            )

            ies.append({
                "iei": iei_val,
                "name": ie_name,
                "field": ie_name,
                "information_element": ie_name,
                "field_path": ie_name,
                "depth": 0,
                "type": type_ref,
                "type_reference": type_ref,
                "presence": presence,
                "format": "Grouped" if is_grouped else "IE",
                "length": "-",
                "applicability": appl,
            })

        return ies

    def _unroll_pfcp_message_ies(
        self,
        base_ies: List[Dict[str, Any]],
        grouped_ies: Dict[str, Dict[str, Any]],
        max_depth: int = 4,
    ) -> List[Dict[str, Any]]:
        unrolled: List[Dict[str, Any]] = []

        def recurse(current_ie: Dict[str, Any], path_prefix: str, depth: int, visited: Set[str]):
            name = current_ie["information_element"]
            current_path = f"{path_prefix}.{name}" if path_prefix else name
            appl = current_ie.get("applicability", "")

            norm_name = self._normalize_key(name)
            type_raw = str(current_ie.get("type_reference", "") or "")

            # Guard: ID fields and primitive Clause 8.2 IEs are always leaves, never containers
            is_scalar_id = norm_name in SCALAR_ID_FIELDS or (norm_name.endswith("id") and norm_name not in ("applicationid", "area_session_id"))
            is_clause_8 = bool(re.search(r"\b8\.2(?:\.\d+)*\b", type_raw))

            target_group = None
            if not is_scalar_id and not is_clause_8:
                # 1. Exact normalized name match
                target_group = grouped_ies.get(norm_name)

                # 2. Match by clause reference extracted from type_reference (e.g. "7.5.2.2")
                if not target_group:
                    clause_match = re.search(r"\b(7\.\d+(?:\.\d+)*(?:-\d+)?)\b", type_raw)
                    if clause_match:
                        target_group = grouped_ies.get(clause_match.group(1))

                # 3. Match by normalized type reference
                if not target_group:
                    type_norm = self._normalize_key(type_raw)
                    target_group = grouped_ies.get(type_norm)

            is_grouped = bool(target_group) and depth < max_depth and norm_name not in visited

            record = dict(current_ie)
            record["name"] = name
            record["field"] = name
            record["field_path"] = current_path
            record["depth"] = depth
            if is_grouped:
                record["format"] = "Grouped"
            unrolled.append(record)

            if is_grouped and target_group:
                new_visited = visited | {norm_name}
                for child_ie in target_group.get("ies", []):
                    child_copy = dict(child_ie)
                    if not child_copy.get("applicability") and appl:
                        child_copy["applicability"] = appl
                    recurse(child_copy, current_path, depth + 1, new_visited)

        for top_ie in base_ies:
            recurse(top_ie, "", 0, set())

        return unrolled