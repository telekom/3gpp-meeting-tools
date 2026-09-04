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
# Match exact frame format clauses (5.5.2.x and 6.5.2.x) and IE coding clauses (5.5.3.x and 6.5.3.x)
RE_UP_FRAME_HEADER = re.compile(r"^((?:5|6)\.5\.2\.\d+)\s*(.*)$")
RE_UP_IE_HEADER = re.compile(r"^((?:5|6)\.5\.3\.\d+)\s*(.*)$")

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
        self.logger.info("Starting TS %s parsing across %d file(s)", self.spec_number, total_files)

        ie_definitions_dict: Dict[str, Dict[str, Any]] = {}
        frame_tables: List[Tuple[str, str, Any]] = []

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(f"Scanning TS 38.415 {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

            self.logger.info("[File %d/%d] Reading %s", file_idx + 1, total_files, docx_path.name)
            root = extract_document_root(docx_path)
            if root is None:
                self.logger.warning("Could not extract word/document.xml from %s", docx_path.name)
                continue

            body = root.find(TAG_BODY)
            if body is None:
                self.logger.warning("No body node found in %s", docx_path.name)
                continue

            in_body = False
            current_clause_num = ""
            current_clause_title = ""
            current_ie_html: List[str] = []
            pending_frame_clause = ""
            pending_frame_name = ""
            pending_p_count = 0

            def _finalize_ie():
                nonlocal current_clause_num, current_clause_title, current_ie_html
                if current_clause_num and current_clause_title and current_ie_html:
                    norm_k = self._normalize_key(current_clause_title)
                    if norm_k not in ie_definitions_dict:
                        self.logger.debug("Captured IE definition: [%s] '%s'", current_clause_num, current_clause_title)
                        ie_definitions_dict[norm_k] = {
                            "clause": current_clause_num,
                            "name": current_clause_title,
                            "ie_name": current_clause_title,
                            "raw_description": "".join(current_ie_html),
                            "structure_table": json.dumps([]),
                        }
                    # Map short acronym token (e.g. QFI, RQI, PPI, SNP, MSNP, PSSN, PSI)
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

                    # Step 1: Ignore Table of Contents and front matter.
                    # Strict match ensures TOC lines like "1 Scope 6" are ignored until the actual heading appears.
                    if not in_body:
                        if re.match(r"^1\.?\s+Scope$", p_text.strip(), re.IGNORECASE):
                            in_body = True
                            self.logger.info("Found Scope boundary ('%s'). Transitioned to specification body.", p_text.strip())
                        continue

                    p_text_clean = re.sub(r"\s+\d+$", "", p_text).strip()

                    # Step 2: Check for Frame Format Clauses (5.5.2.1, 5.5.2.2, 6.5.2.1)
                    match_frame = RE_UP_FRAME_HEADER.match(p_text_clean)
                    if match_frame:
                        _finalize_ie()
                        cl_num = match_frame.group(1).strip()
                        cl_title = self._clean_frame_name(match_frame.group(2), cl_num)

                        if pending_frame_clause:
                            self.logger.warning(
                                "Overwriting unfulfilled frame candidate [%s] '%s' with new candidate [%s] '%s'",
                                pending_frame_clause, pending_frame_name, cl_num, cl_title
                            )

                        pending_frame_clause = cl_num
                        pending_frame_name = cl_title
                        pending_p_count = 0
                        self.logger.info("Detected frame heading candidate: [%s] '%s'", cl_num, cl_title)
                        continue

                    # Step 3: Check for IE Definition Clauses (5.5.3.x and 6.5.3.x)
                    match_ie = RE_UP_IE_HEADER.match(p_text_clean)
                    if match_ie:
                        _finalize_ie()
                        current_clause_num = match_ie.group(1).strip()
                        current_clause_title = match_ie.group(2).strip()
                        self.logger.debug("Entering IE coding section: [%s] '%s'", current_clause_num, current_clause_title)
                        current_ie_html = [
                            f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
                            f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                            f'{html.escape(current_clause_title)} (Clause {current_clause_num})</h2>'
                        ]
                        if pending_frame_clause:
                            self.logger.warning(
                                "Entered IE coding clause [%s] while frame [%s] '%s' was still pending a table.",
                                current_clause_num, pending_frame_clause, pending_frame_name
                            )
                        pending_frame_clause = ""
                        pending_frame_name = ""
                        continue
                    elif re.match(r"^(?:[1-6]\.[0-9]|Annex)", p_text_clean):
                        _finalize_ie()

                    # Paragraph distance tracking between frame heading and table
                    if pending_frame_clause:
                        pending_p_count += 1
                        self.logger.debug(
                            "Paragraph %d after frame [%s]: '%s'",
                            pending_p_count, pending_frame_clause, p_text_clean[:60]
                        )
                        if pending_p_count > 15:
                            self.logger.warning(
                                "Discarded pending frame candidate [%s] '%s' after %d paragraphs without finding a table.",
                                pending_frame_clause, pending_frame_name, pending_p_count
                            )
                            pending_frame_clause = ""
                            pending_frame_name = ""

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
                    num_rows = len(elem.findall(TAG_TR))
                    self.logger.debug(
                        "Encountered table with %d rows. Active context -> pending_frame: [%s] '%s', current_ie: [%s]",
                        num_rows, pending_frame_clause, pending_frame_name, current_clause_num
                    )

                    if pending_frame_clause:
                        is_valid = self._is_valid_frame_table(elem)
                        if is_valid:
                            self.logger.info(
                                "Assigned table (%d rows) to frame format [%s] '%s'",
                                num_rows, pending_frame_clause, pending_frame_name
                            )
                            frame_tables.append((pending_frame_clause, pending_frame_name, elem))
                        else:
                            self.logger.warning(
                                "Table (%d rows) following frame heading [%s] rejected by _is_valid_frame_table.",
                                num_rows, pending_frame_clause
                            )
                        pending_frame_clause = ""
                        pending_frame_name = ""

                    if current_clause_num:
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=True)
                        if tbl_html:
                            current_ie_html.append(tbl_html)

            _finalize_ie()
            if not in_body:
                self.logger.error("Failed to detect '1 Scope' in %s. Ingestion was skipped!", docx_path.name)

        if progress_callback:
            progress_callback("Extracting bit-level frame formats for PDU Session & PDU Set User Plane...", 65)

        self.logger.info("Found %d raw frame table(s) across specifications", len(frame_tables))
        messages: List[Dict[str, Any]] = []

        for clause_num, frame_name, tbl_elem in frame_tables:
            clean_msg_name = self._clean_frame_name(frame_name, clause_num)
            self.logger.info("Parsing fields for frame: [%s] '%s'", clause_num, clean_msg_name)
            ies = self._parse_frame_table(tbl_elem, clause_num, ie_definitions_dict)

            if ies:
                self.logger.info("Extracted %d field(s) for message '%s'", len(ies), clean_msg_name)
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
            else:
                self.logger.warning("Frame [%s] '%s' yielded 0 fields after parsing!", clause_num, clean_msg_name)

        ie_definitions = list(ie_definitions_dict.values())
        self.logger.info(
            "Completed parsing TS %s: Successfully generated %d message(s) and %d definition(s)",
            self.spec_number, len(messages), len(ie_definitions)
        )
        for idx, m in enumerate(messages, 1):
            self.logger.info("  [%d] Message: '%s' (Clause %s) with %d fields", idx, m["message_name"], m["clause"], len(m["ies"]))

        return messages, ie_definitions

    def _is_valid_frame_table(self, tbl_elem: Any) -> bool:
        """Verifies that a Word table contains a genuine bit/octet frame layout."""
        rows = tbl_elem.findall(TAG_TR)
        if len(rows) < 2:
            self.logger.debug("Table rejected: less than 2 rows (%d)", len(rows))
            return False

        first_rows_text = " ".join(
            _extract_tc_text(c).lower()
            for r in rows[:2]
            for c in r.findall(TAG_TC)
        )
        has_bit = "bit" in first_rows_text
        has_octet = "octet" in first_rows_text

        if not (has_bit and has_octet):
            self.logger.debug(
                "Table rejected by header check (has_bit=%s, has_octet=%s). Sample: '%s'",
                has_bit, has_octet, first_rows_text[:80]
            )
            return False

        return True

    def _clean_frame_name(self, raw_name: str, clause_num: str) -> str:
        s = raw_name.strip()
        s = re.sub(r"(?i)\s*format$", "", s).strip()
        s = re.sub(r"\s+\d+$", "", s).strip()
        if not s or s.isdigit():
            if "5.5.2.1" in clause_num:
                return "DL PDU SESSION INFORMATION (PDU Type 0)"
            elif "5.5.2.2" in clause_num:
                return "UL PDU SESSION INFORMATION (PDU Type 1)"
            elif "6.5.2.1" in clause_num:
                return "DL PDU SET INFORMATION (PDU Type 0)"
            return f"User Plane Frame (Clause {clause_num})"
        return s

    def _is_header_or_bit_scale_row(self, cells: List[Any]) -> bool:
        """Identifies header rows and bit-scale indicator rows (e.g. 7 | 6 | 5 | 4 | 3 | 2 | 1 | 0)."""
        texts = [_extract_tc_text(c).strip() for c in cells if _extract_tc_text(c).strip()]
        if not texts:
            return True
        # Row of pure bit numbers 0-7
        if all(t.isdigit() and int(t) <= 7 for t in texts):
            return True
        # Header row containing "Bits" and "Number of Octets"
        joined = " ".join(t.lower() for t in texts)
        if any(k in joined for k in ("number of octet", "number of octets", "bits")):
            non_hdr_tokens = [t for t in texts if t.lower() not in ("bits", "octets", "number of octets", "number of octet")]
            if all(t.isdigit() and int(t) <= 7 for t in non_hdr_tokens):
                return True
        return False

    def _parse_frame_table(
        self, tbl_elem: Any, clause_num: str, ie_defs: Dict[str, Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        is_pdu_set = clause_num.startswith("6.")
        default_appl = "NG-U, Xn-U, F1-U, N9" if is_pdu_set else "NG-U, Xn-U, N9"

        ies: List[Dict[str, Any]] = []
        current_octet = 0

        for r_idx, row in enumerate(rows):
            cells = row.findall(TAG_TC)
            if not cells or self._is_header_or_bit_scale_row(cells):
                continue

            last_cell_text = _extract_tc_text(cells[-1]).strip()
            data_cells = cells[:-1] if len(cells) > 1 else cells

            # Track octet numbering
            if last_cell_text.isdigit():
                current_octet += int(last_cell_text)
            elif "or" in last_cell_text or "-" in last_cell_text:
                current_octet += 1

            for c in data_cells:
                raw_text = _extract_tc_text(c).strip()
                if not raw_text or raw_text.lower() in ("bits", "number of octets", "number of octet"):
                    continue

                if raw_text.isdigit() and int(raw_text) <= 7:
                    continue

                span = self._get_tc_grid_span(c)
                clean_name = self._clean_field_name(raw_text)
                norm_k = self._normalize_key(clean_name)

                # Determine bit or octet length
                if span < 8 and (len(data_cells) > 1 or not any(k in last_cell_text for k in ("octet", "or"))):
                    length_desc = f"{span} bit" if span == 1 else f"{span} bits"
                    fmt_desc = "Flag" if span == 1 else "Bit field"
                    bit_count = span
                else:
                    length_desc = f"{last_cell_text} octets" if last_cell_text.isdigit() else last_cell_text
                    bit_count = int(last_cell_text) * 8 if last_cell_text.isdigit() else 0
                    if "time" in norm_k:
                        fmt_desc = "Timestamp"
                    elif any(k in norm_k for k in ("sn", "sequencenumber", "result", "size", "importance")):
                        fmt_desc = "Integer"
                    else:
                        fmt_desc = "Octet field"

                # Multi-row split field stitching (e.g. PSSN across Octets 2 & 3)
                if ies and ies[-1]["name"] == clean_name and bit_count > 0:
                    prev_len_str = ies[-1]["length"]
                    prev_match = re.match(r"^(\d+)\s+bit", prev_len_str)
                    if prev_match:
                        combined_bits = int(prev_match.group(1)) + bit_count
                        ies[-1]["length"] = f"{combined_bits} bits"
                        ies[-1]["format"] = "Bit field"
                        self.logger.debug("Stitched multi-row field '%s': %s -> %s", clean_name, prev_len_str, ies[-1]["length"])
                        continue

                display_name = clean_name
                if norm_k == "spare":
                    display_name = f"Spare (Octet {current_octet})" if current_octet > 0 else f"Spare ({length_desc})"

                if last_cell_text.startswith("0") or "0 or" in last_cell_text or "0-" in last_cell_text:
                    presence = CONDITIONAL_FIELD_TRIGGERS.get(norm_k, "C")
                else:
                    presence = "M"

                ref_info = ie_defs.get(norm_k)
                if not ref_info:
                    base_token = re.sub(r"\(.*?\)", "", clean_name).strip()
                    ref_info = ie_defs.get(self._normalize_key(base_token))

                clause_ref = f"Clause {ref_info['clause']}" if ref_info else f"Clause {clause_num}"
                type_ref = f"{clean_name} ({clause_ref})"

                ies.append({
                    "iei": "-",
                    "name": display_name,
                    "field": display_name,
                    "information_element": display_name,
                    "field_path": display_name,
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
        s = re.sub(r"\(=[01]\)", "", s).strip()
        s = re.sub(r"\s+", " ", s).strip()
        return s