import html
import json
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY, TAG_P, TAG_TBL, TAG_TC, TAG_TR
)
from modules.nas.core.parsing.protocol_parser_utils import (
    extract_document_root, _extract_p_text, _extract_tc_text, _convert_table_to_html
)

RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_UP_CLAUSE_HEADER = re.compile(r"^((?:5|6)\.(?:5\.2|5\.3)(?:\.[0-9A-Za-z]+)*)\s*(.*)$")
RE_UP_FIGURE_CAPTION = re.compile(r"^Figure\s+((?:5|6)\.5\.2\.\d+-\d+)\s*[:\.]\s*(.+)$", re.IGNORECASE)

# Mapping of known field names to their conditional presence triggers in TS 38.415
CONDITIONAL_FIELD_TRIGGERS = {
    "ppi": "C (PPP=1)",
    "pagingpolicyindicator": "C (PPP=1)",
    "dlsendingtimestamp": "C (QMP=1)",
    "dlsendingtimestamprepeated": "C (QMP=1)",
    "dlreceivedtimestamp": "C (QMP=1)",
    "ulsendingtimestamp": "C (QMP=1)",
    "dlqfisequencenumber": "C (SNP=1)",
    "ulqfisequencenumber": "C (SNP=1)",
    "dlmbsqfisequencenumber": "C (MSNP=1)",
    "dldelayresult": "C (DL Delay Ind.=1)",
    "uldelayresult": "C (UL Delay Ind.=1)",
    "n3n9delayresult": "C (N3/N9 Delay Ind.=1)",
    "d1ulpdcpdelayresultind": "C (New IE Flag 0=1)",
    "ulcongestioninformation": "C (New IE Flag 1=1)",
    "dlcongestioninformation": "C (New IE Flag 2=1)",
    "pssize": "C (PSSI=1)",
    "pdusetsize": "C (PSSI=1)",
    "padding": "O",
}


class PDUSessionUPDocxParser:
    """Dedicated parser for 3GPP TS 38.415 (PDU Session & PDU Set User Plane Protocols)."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.415"):
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
    def _get_tc_grid_span(tc_elem: Any) -> int:
        for elem in tc_elem.iter():
            if elem.tag.endswith("gridSpan"):
                for attr_name, attr_val in elem.attrib.items():
                    if attr_name.endswith("val") and attr_val.isdigit():
                        return int(attr_val)
        return 1

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            self.logger.error("No valid specification docx files found: %s", self.docx_paths)
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        total_files = len(valid_paths)
        raw_tables: List[Tuple[str, str, str, Any]] = []
        ie_definitions_dict: Dict[str, Dict[str, Any]] = {}
        frame_tables: List[Tuple[str, str, Any]] = []

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(f"Scanning TS 38.415 {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            current_clause_num = ""
            current_clause_title = ""
            current_ie_html: List[str] = []
            pending_frame_clause = ""
            pending_frame_name = ""

            def _finalize_ie():
                nonlocal current_clause_num, current_clause_title, current_ie_html
                if current_clause_num and current_clause_title and current_ie_html:
                    norm_k = self._normalize_key(current_clause_title)
                    if norm_k not in ie_definitions_dict:
                        ie_definitions_dict[norm_k] = {
                            "clause": current_clause_num,
                            "name": current_clause_title,
                            "ie_name": current_clause_title,
                            "raw_description": "".join(current_ie_html),
                            "structure_table": json.dumps([]),
                        }
                    # Also map acronym/short token (e.g. QFI, RQI, PPI, SNP)
                    paren_match = re.search(r"\(([^)]+)\)", current_clause_title)
                    if paren_match:
                        short_k = self._normalize_key(paren_match.group(1))
                        if short_k and short_k not in ie_definitions_dict:
                            ie_definitions_dict[short_k] = ie_definitions_dict[norm_k]

                current_clause_num = ""
                current_clause_title = ""
                current_ie_html = []

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue

                    match_hdr = RE_UP_CLAUSE_HEADER.match(p_text)
                    if match_hdr:
                        cl_num = match_hdr.group(1).strip()
                        cl_title = match_hdr.group(2).strip()

                        # Frame Format Clauses (Clause 5.5.2.x and 6.5.2.x)
                        if ".5.2." in cl_num and cl_title:
                            _finalize_ie()
                            pending_frame_clause = cl_num
                            pending_frame_name = cl_title
                            continue

                        # IE Coding Definition Clauses (Clause 5.5.3.x and 6.5.3.x)
                        if ".5.3." in cl_num and cl_title:
                            _finalize_ie()
                            current_clause_num = cl_num
                            current_clause_title = cl_title
                            current_ie_html = [
                                f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                                f'{html.escape(cl_title)} (Clause {cl_num})</h2>'
                            ]
                            continue
                        else:
                            _finalize_ie()

                    # Check for Figure caption identifying a preceding frame table
                    match_fig = RE_UP_FIGURE_CAPTION.match(p_text)
                    if match_fig:
                        fig_num = match_fig.group(1).strip()
                        fig_title = match_fig.group(2).strip()
                        if frame_tables and frame_tables[-1][1] == "":
                            # Associate previous unnamed frame table with this caption
                            last_idx = len(frame_tables) - 1
                            prev_clause, _, tbl = frame_tables[last_idx]
                            frame_tables[last_idx] = (prev_clause, fig_title, tbl)

                    if current_clause_num:
                        if p_text.startswith(("Description:", "Value range:", "Field length:", "Field Length:")):
                            bold_prefix, _, rest = p_text.partition(":")
                            current_ie_html.append(
                                f'<p style="margin: 4px 0; line-height: 1.4; color: #1E293B; font-size: 11px;">'
                                f'<b>{html.escape(bold_prefix)}:</b> {html.escape(rest.strip())}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_html.append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">'
                                f'{html.escape(p_text)}</div>'
                            )
                        elif len(current_ie_html) < 20:
                            current_ie_html.append(
                                f'<p style="margin: 3px 0; line-height: 1.4; color: #334155; font-size: 11px;">{html.escape(p_text)}</p>'
                            )

                elif elem.tag == TAG_TBL:
                    if pending_frame_clause:
                        frame_tables.append((pending_frame_clause, pending_frame_name, elem))
                        pending_frame_clause = ""
                        pending_frame_name = ""

                    if current_clause_num:
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=True)
                        if tbl_html:
                            current_ie_html.append(tbl_html)

            _finalize_ie()

        if progress_callback:
            progress_callback("Extracting bit-level frame formats for PDU Session & PDU Set User Plane...", 65)

        # Build message records from frame tables
        messages: List[Dict[str, Any]] = []
        for clause_num, frame_name, tbl_elem in frame_tables:
            clean_msg_name = self._clean_frame_name(frame_name, clause_num)
            ies = self._parse_frame_table(tbl_elem, clause_num, ie_definitions_dict)

            if ies:
                # Also store the complete HTML table figure in ie_definitions under the message name
                tbl_figure_html = _convert_table_to_html(tbl_elem, is_figure_diagram=True)
                ie_definitions_dict[self._normalize_key(clean_msg_name)] = {
                    "clause": clause_num,
                    "name": clean_msg_name,
                    "ie_name": clean_msg_name,
                    "raw_description": (
                        f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
                        f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                        f'{clean_msg_name} (Clause {clause_num})</h2>'
                        f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-top: 8px; margin-bottom: 4px;">'
                        f'Frame Format & Bit Grid Layout:</div>{tbl_figure_html}'
                    ),
                    "structure_table": json.dumps([]),
                }

                messages.append({
                    "clause": clause_num,
                    "clause_ref": clause_num,
                    "name": clean_msg_name,
                    "message_name": clean_msg_name,
                    "table_caption": f"{clean_msg_name} Frame Format (Clause {clause_num})",
                    "ies": ies,
                    "fields": ies,
                })

        ie_definitions = list(ie_definitions_dict.values())
        self.logger.info("Completed parsing TS %s: Loaded %d messages and %d definitions",
                         self.spec_number, len(messages), len(ie_definitions))
        return messages, ie_definitions

    def _clean_frame_name(self, raw_name: str, clause_num: str) -> str:
        s = raw_name.strip()
        s = re.sub(r"(?i)\s*format$", "", s).strip()
        if not s:
            if "5.5.2.1" in clause_num:
                return "DL PDU SESSION INFORMATION (PDU Type 0)"
            elif "5.5.2.2" in clause_num:
                return "UL PDU SESSION INFORMATION (PDU Type 1)"
            elif "6.5.2.1" in clause_num:
                return "DL PDU SET INFORMATION (PDU Type 0)"
            return f"User Plane Frame (Clause {clause_num})"
        return s

    def _parse_frame_table(
        self, tbl_elem: Any, clause_num: str, ie_defs: Dict[str, Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """Extracts individual bit fields and octet fields from TS 38.415 frame tables."""
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        # Determine default reference points based on section
        is_pdu_set = clause_num.startswith("6.")
        default_appl = "NG-U, Xn-U, F1-U, N9" if is_pdu_set else "NG-U, Xn-U, N9"

        header_idx = -1
        for r_idx in range(min(3, len(rows))):
            row_text = " ".join(_extract_tc_text(c).lower() for c in rows[r_idx].findall(TAG_TC))
            if any(k in row_text for k in ("bit", "octet", "number of octet")):
                header_idx = r_idx
                break

        data_start_idx = header_idx + 1 if header_idx != -1 else 1
        ies: List[Dict[str, Any]] = []

        for r_offset in range(data_start_idx, len(rows)):
            row = rows[r_offset]
            cells = row.findall(TAG_TC)
            if not cells:
                continue

            # The final cell contains the Octet count/position (e.g., '1', '0 or 1', '0 or 8', '0-3')
            last_cell_text = _extract_tc_text(cells[-1]).strip()
            data_cells = cells[:-1] if len(cells) > 1 else cells

            for c in data_cells:
                raw_text = _extract_tc_text(c).strip()
                if not raw_text or raw_text.lower() in ("bits", "number of octets"):
                    continue

                span = self._get_tc_grid_span(c)
                clean_name = self._clean_field_name(raw_text)
                norm_k = self._normalize_key(clean_name)

                # Determine bit or octet length
                if span < 8 and (len(data_cells) > 1 or not any(k in last_cell_text for k in ("octet", "or"))):
                    length_desc = f"{span} bit" if span == 1 else f"{span} bits"
                    fmt_desc = "Flag" if span == 1 else "Bit field"
                else:
                    length_desc = f"{last_cell_text} octets" if last_cell_text.isdigit() else last_cell_text
                    if "time" in norm_k:
                        fmt_desc = "Timestamp"
                    elif any(k in norm_k for k in ("sn", "sequencenumber", "result", "size", "importance")):
                        fmt_desc = "Integer"
                    else:
                        fmt_desc = "Octet field"

                # Determine presence requirement
                if last_cell_text.startswith("0") or "0 or" in last_cell_text or "0-" in last_cell_text:
                    presence = CONDITIONAL_FIELD_TRIGGERS.get(norm_k, "C")
                else:
                    presence = "M"

                # Resolve Clause 5.5.3 / 6.5.3 definition link
                ref_info = ie_defs.get(norm_k)
                if not ref_info:
                    # Try matching without acronym parentheses
                    base_token = re.sub(r"\(.*?\)", "", clean_name).strip()
                    ref_info = ie_defs.get(self._normalize_key(base_token))

                clause_ref = f"Clause {ref_info['clause']}" if ref_info else f"Clause {clause_num}"
                type_ref = f"{clean_name} ({clause_ref})"

                ies.append({
                    "iei": "-",
                    "name": clean_name,
                    "field": clean_name,
                    "information_element": clean_name,
                    "field_path": clean_name,
                    "depth": 0,
                    "type": type_ref,
                    "type_reference": type_ref,
                    "presence": presence,
                    "format": fmt_desc,
                    "length": length_desc,
                    "applicability": default_appl,
                })

        return ies

    @staticmethod
    def _clean_field_name(raw: str) -> str:
        s = raw.strip()
        # Normalize PDU Type (=0) -> PDU Type
        s = re.sub(r"\(=[01]\)", "", s).strip()
        # Clean extra spaces in Delay labels
        s = re.sub(r"\s+", " ", s).strip()
        return s