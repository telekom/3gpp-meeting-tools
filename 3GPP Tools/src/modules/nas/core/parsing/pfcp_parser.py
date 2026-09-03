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
RE_CLEAN_GROUPED_NAME = re.compile(r"^(?:Information\s+Elements\s+in\s+)?(.+?)(?:\s+IE(?:\s+within\s+.+)?|\s+within\s+.+)?$", re.IGNORECASE)

TOP_LEVEL_PFCP_CLAUSES = {
    "7.4.2.1", "7.4.2.2", "7.4.3.1", "7.4.3.2", "7.4.4.1", "7.4.4.2", "7.4.4.3",
    "7.4.4.4", "7.4.4.5", "7.4.4.6", "7.4.4.7", "7.4.5.1", "7.4.5.2", "7.4.6.1",
    "7.4.6.2", "7.4.7.1", "7.4.7.2", "7.5.2.1", "7.5.3.1", "7.5.4.1", "7.5.5.1",
    "7.5.7.1", "7.5.8.1", "7.5.9.1", "7.5.9.2"
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
        """Strips table suffixes (e.g. '7.5.2.1-1' -> '7.5.2.1')."""
        return clause_ref.split("-")[0].strip()

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            self.logger.error("No valid specification docx files found: %s", self.docx_paths)
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        total_files = len(valid_paths)
        self.logger.info("Starting PFCP ingestion for TS %s across %d document part(s)", self.spec_number, total_files)

        raw_tables: List[Tuple[str, str, str, Any]] = []  # (clause, caption_title, full_caption, tbl_elem)
        ie_definitions_dict: Dict[str, Dict[str, Any]] = {}
        type_registry: Dict[str, Dict[str, str]] = {}  # norm_name -> {type_id, clause, applicability}

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(f"Scanning PFCP {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

            self.logger.info("Parsing document part %d/%d: %s", file_idx + 1, total_files, docx_path.name)
            root = extract_document_root(docx_path)
            if root is None:
                self.logger.warning("Failed to extract XML document root from %s", docx_path.name)
                continue

            body = root.find(TAG_BODY)
            if body is None:
                self.logger.warning("Document body missing in %s", docx_path.name)
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
                        self.logger.debug("Indexed IE definition: %s (Clause %s)", current_ie_name, current_ie_clause)
                current_ie_clause = ""
                current_ie_name = ""
                current_ie_html = []

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue

                    # Match Clause Headings (e.g. 8.2.1 Cause)
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

                    # Match Table Captions (e.g. Table 7.5.2.1-1: Information Elements in PFCP Session Establishment Request)
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

        if not type_registry:
            self.logger.warning("Table 8.1.2-1 not found or empty. IE type resolution will fallback to names.")

        # 2. Parse Clause 7 tables into messages and grouped IEs
        top_messages_raw: List[Dict[str, Any]] = []
        grouped_ies: Dict[str, Dict[str, Any]] = {}

        for clause_ref, cap_name, full_cap, tbl_elem in raw_tables:
            if not clause_ref.startswith("7."):
                continue

            parsed_ies = self._parse_clause_7_table(tbl_elem, type_registry)
            if not parsed_ies:
                self.logger.debug("Skipping Clause 7 table with no detectable IE rows: %s (%s)", clause_ref, cap_name)
                continue

            clean_name = self._sanitize_table_name(cap_name)
            base_clause = self._clean_clause_number(clause_ref)
            is_top_msg = self._is_top_level_message(clean_name, base_clause)

            table_record = {
                "clause": base_clause,
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
                self.logger.debug("Registered grouped IE container [%s] -> %s (%d fields)", base_clause, norm_key, len(parsed_ies))
                # Register short form without action prefixes for flexible resolution
                stripped_k = re.sub(r"^(create|update|remove|created|updated)", "", norm_key)
                if stripped_k and stripped_k not in grouped_ies:
                    grouped_ies[stripped_k] = table_record

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
        s = caption.strip()
        match = RE_CLEAN_GROUPED_NAME.match(s)
        if match:
            s = match.group(1).strip()
        s = re.sub(r"\s+message\s+content$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"^Information\s+Elements\s+in\s+(?:a\s+)?", "", s, flags=re.IGNORECASE)
        return s.strip()

    @staticmethod
    def _is_top_level_message(name: str, base_clause: str) -> bool:
        lower_name = name.lower()
        if any(lower_name.endswith(kw) for kw in ("request", "response", "reject", "indication")):
            if not any(sub in lower_name for sub in ("within", " ie in", " ie within")):
                return True
        if base_clause in TOP_LEVEL_PFCP_CLAUSES:
            return True
        return False

    def _parse_clause_7_table(
        self, tbl_elem: Any, type_registry: Dict[str, Dict[str, str]]
    ) -> List[Dict[str, Any]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_idx = -1
        name_col, pres_col, comment_col, appl_col, type_col = -1, -1, -1, -1, -1

        for r_idx in range(min(3, len(rows))):
            cells = [_extract_tc_text(tc).lower() for tc in rows[r_idx].findall(TAG_TC)]
            joined = " ".join(cells)
            if "information element" in joined:
                header_idx = r_idx
                for idx, c in enumerate(cells):
                    if "information element" in c:
                        name_col = idx
                    elif c in ("p", "presence"):
                        pres_col = idx
                    elif any(k in c for k in ("condition", "comment")):
                        comment_col = idx
                    elif any(k in c for k in ("appl", "applicability", "interface")):
                        appl_col = idx
                    elif any(k in c for k in ("ie type", "type", "clause", "reference")):
                        type_col = idx
                break

        if header_idx == -1 or name_col == -1:
            return []

        ies: List[Dict[str, Any]] = []
        for row in rows[header_idx + 1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) <= name_col:
                continue

            ie_name = cells[name_col]
            if not ie_name or any(kw in ie_name.lower() for kw in ("information element", "octet 1", "octets")):
                continue

            presence = cells[pres_col] if pres_col != -1 and len(cells) > pres_col else "O"
            type_ref = cells[type_col] if type_col != -1 and len(cells) > type_col else ie_name
            appl = cells[appl_col] if appl_col != -1 and len(cells) > appl_col else ""

            norm_name = self._normalize_key(ie_name)
            reg_info = type_registry.get(norm_name, {})
            iei_val = reg_info.get("type_id", "")

            if not appl and reg_info.get("applicability"):
                appl = reg_info["applicability"]

            if not type_ref or type_ref == ie_name:
                if reg_info.get("clause"):
                    type_ref = f"{ie_name} ({reg_info['clause']})"

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
                "format": "Grouped" if "within" in type_ref.lower() or "create" in ie_name.lower() else "IE",
                "length": "-",
                "applicability": appl,
            })

        return ies

    def _unroll_pfcp_message_ies(
        self,
        base_ies: List[Dict[str, Any]],
        grouped_ies: Dict[str, Dict[str, Any]],
        max_depth: int = 3,
    ) -> List[Dict[str, Any]]:
        unrolled: List[Dict[str, Any]] = []

        def recurse(current_ie: Dict[str, Any], path_prefix: str, depth: int, visited: Set[str]):
            name = current_ie["information_element"]
            current_path = f"{path_prefix}.{name}" if path_prefix else name
            appl = current_ie.get("applicability", "")

            norm_name = self._normalize_key(name)
            target_group = (
                grouped_ies.get(norm_name)
                or grouped_ies.get(self._normalize_key(current_ie.get("type_reference", "")))
                or grouped_ies.get(re.sub(r"^(create|update|remove)", "", norm_name))
            )

            is_grouped = bool(target_group) and depth < max_depth and norm_name not in visited

            record = dict(current_ie)
            record["name"] = name
            record["field"] = name
            record["field_path"] = current_path
            record["depth"] = depth
            if is_grouped:
                record["format"] = "Grouped"
            unrolled.append(record)

            if is_grouped:
                new_visited = visited | {norm_name}
                for child_ie in target_group["ies"]:
                    child_copy = dict(child_ie)
                    if not child_copy.get("applicability") and appl:
                        child_copy["applicability"] = appl
                    recurse(child_copy, current_path, depth + 1, new_visited)

        for top_ie in base_ies:
            recurse(top_ie, "", 0, set())

        return unrolled