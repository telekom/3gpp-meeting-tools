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
RE_GTPU_CLAUSE_HEADER = re.compile(r"^((?:5\.2\.2|7\.[23]|8)(?:\.[0-9A-Za-z]+)*)\s*(.*)$")
RE_GTPU_TABLE_CAPTION = re.compile(r"^Table\s+([5678]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+)$", re.IGNORECASE)

# Static interface applicability mappings based on TS 29.281 Clause 5.2 and Clause 7
GTPU_EXT_HEADER_INTERFACES = {
    "udpport": "All Interfaces",
    "pdcpdu number": "S1-U, X2, N3, Xn, Iu",
    "pdcpdunumber": "S1-U, X2, N3, Xn, Iu",
    "longpdcppdunumber": "S1-U, X2, N3, Xn",
    "serviceclassindicator": "Gn, Gp, S5/S8, S4, S1-U, S12, Iu, S11-U",
    "rancontainer": "X2",
    "xwrancontainer": "Xw",
    "nrrancontainer": "X2-U, Xn-U, F1-U",
    "pdusessioncontainer": "N3, N9, N3mb, N19mb",
    "pdusetinformationcontainer": "N3, N9, Xn-U, F1-U",
}

GTPU_MESSAGE_INTERFACES = {
    "echorequest": "All Interfaces",
    "echoresponse": "All Interfaces",
    "supportedextensionheadersnotification": "All Interfaces",
    "errorindication": "All Interfaces",
    "endmarker": "S1-U, S11-U, S5/S8, N3, N9, X2, Xn",
    "tunnelstatus": "S5/S8, N3, N9",
    "gpdu": "N3, N9, N19, S1-U, S5/S8, F1-U, X2-U, Xn-U, W1-U, Gn, Gp, S4, S2a, S2b, S12, M1, Sn",
}


class GTPUDocxParser:
    """Dedicated parser for 3GPP TS 29.281 (GTPv1-U) specifications."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "29.281"):
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
        type_registry: Dict[str, Dict[str, str]] = {}
        ext_header_registry: Dict[str, Dict[str, str]] = {}

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(f"Scanning GTP-U {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

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

                    match_header = RE_GTPU_CLAUSE_HEADER.match(p_text)
                    if match_header:
                        cl_num = match_header.group(1).strip()
                        cl_title = match_header.group(2).strip()

                        # Capture Clause 8 Information Elements and Clause 5.2.2 Extension Headers
                        is_target_section = (
                            (cl_num.startswith("8.") and cl_num not in ("8", "8.1") and cl_title)
                            or (cl_num.startswith("5.2.2") and cl_title)
                        )

                        if is_target_section:
                            _finalize_current_ie()
                            current_ie_clause = cl_num
                            current_ie_name = cl_title
                            current_ie_html = [
                                f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                                f'{html.escape(cl_title)} (Clause {cl_num})</h2>'
                            ]
                        elif not (cl_num.startswith("8.") or cl_num.startswith("5.2.2")):
                            _finalize_current_ie()

                    match_cap = RE_GTPU_TABLE_CAPTION.match(p_text)
                    if match_cap:
                        clause_ref = match_cap.group(1).strip()
                        cap_name = match_cap.group(2).strip()
                        last_caption_info = (clause_ref, cap_name, p_text)
                    elif not p_text.startswith("Table "):
                        last_caption_info = None

                    if current_ie_clause:
                        if p_text.startswith(("Figure 5.", "Figure 8.")):
                            current_ie_html.append(
                                f'<p style="font-weight: bold; color: #475569; margin-top: 8px; margin-bottom: 3px; font-size: 11px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith(("Table 5.", "Table 8.")):
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

        if progress_callback:
            progress_callback("Analyzing Table 8.1-1 and Extension Header registries...", 60)

        # 1. Parse Master Registries (Table 8.1-1 and Table 5.2.1-3)
        self._seed_default_registries(type_registry, ext_header_registry)
        for clause_ref, cap_name, _, tbl_elem in raw_tables:
            cap_low = cap_name.lower()
            if "8.1-1" in clause_ref or "information elements" in cap_low:
                self._parse_clause_8_type_registry(tbl_elem, type_registry)
            elif "5.2.1-3" in clause_ref or "extension header type" in cap_low:
                self._parse_extension_header_registry(tbl_elem, ext_header_registry)

        # 2. Parse Clause 7 Signalling Message Tables
        messages: List[Dict[str, Any]] = []
        for clause_ref, cap_name, full_cap, tbl_elem in raw_tables:
            if not clause_ref.startswith("7."):
                continue

            parsed_ies = self._parse_clause_7_table(tbl_elem, type_registry, clause_ref, cap_name)
            if not parsed_ies:
                continue

            clean_name = self._sanitize_message_name(cap_name)
            base_clause = self._clean_clause_number(clause_ref)

            messages.append({
                "clause": base_clause,
                "clause_ref": clause_ref,
                "name": clean_name,
                "message_name": clean_name,
                "table_caption": full_cap,
                "ies": parsed_ies,
                "fields": parsed_ies,
            })

        # 3. Synthesize G-PDU Message (carrying T-PDU and Extension Headers)
        gpdu_msg = self._build_gpdu_message(ext_header_registry)
        messages.append(gpdu_msg)

        ie_definitions = list(ie_definitions_dict.values())
        self.logger.info("Completed parsing TS %s: Returning %d messages and %d IE definitions",
                         self.spec_number, len(messages), len(ie_definitions))
        return messages, ie_definitions

    def _seed_default_registries(self, type_reg: Dict[str, Dict[str, str]], ext_reg: Dict[str, Dict[str, str]]):
        """Pre-seeds standard GTP-U types to ensure parsing resilience across older releases."""
        default_types = {
            "recovery": {"type_id": "14", "format": "TV", "clause": "8.2", "length": "1"},
            "tunnelendpointidentifierdatai": {"type_id": "16", "format": "TV", "clause": "8.3", "length": "4"},
            "gtpupeeraddress": {"type_id": "133", "format": "TLV", "clause": "8.4", "length": "4 or 16"},
            "gsnaddress": {"type_id": "133", "format": "TLV", "clause": "8.4", "length": "4 or 16"},
            "extensionheadertypelist": {"type_id": "141", "format": "TLV", "clause": "8.5", "length": "Variable"},
            "gtputunnelstatusinformation": {"type_id": "230", "format": "TLV", "clause": "8.7", "length": "Variable"},
            "recoverytimestamp": {"type_id": "231", "format": "TLV", "clause": "8.7", "length": "Variable"},
            "privateextension": {"type_id": "255", "format": "TLV", "clause": "8.6", "length": "Variable"},
        }
        type_reg.update(default_types)

        default_exts = {
            "udpport": {"iei": "0x40", "clause": "5.2.2.1", "name": "UDP Port"},
            "pdcppdunumber": {"iei": "0xC0", "clause": "5.2.2.2", "name": "PDCP PDU Number"},
            "longpdcppdunumber": {"iei": "0x03", "clause": "5.2.2.2A", "name": "Long PDCP PDU Number"},
            "serviceclassindicator": {"iei": "0x20", "clause": "5.2.2.3", "name": "Service Class Indicator"},
            "rancontainer": {"iei": "0x81", "clause": "5.2.2.4", "name": "RAN Container"},
            "xwrancontainer": {"iei": "0x83", "clause": "5.2.2.5", "name": "Xw RAN Container"},
            "nrrancontainer": {"iei": "0x84", "clause": "5.2.2.6", "name": "NR RAN Container"},
            "pdusessioncontainer": {"iei": "0x85", "clause": "5.2.2.7", "name": "PDU Session Container"},
            "pdusetinformationcontainer": {"iei": "0x86", "clause": "5.2.2.8", "name": "PDU Set Information Container"},
        }
        ext_reg.update(default_exts)

    def _parse_clause_8_type_registry(self, tbl_elem: Any, registry: Dict[str, Dict[str, str]]):
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return

        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) >= 3:
                type_val = cells[0]
                fmt_val = cells[1] if len(cells) > 3 else "TLV"
                name_val = cells[2] if len(cells) > 3 else cells[1]
                ref_val = cells[3] if len(cells) > 3 else cells[2]

                clean_name = re.sub(r"(?i)\s*see\s+note.*$", "", name_val).strip()
                if clean_name and type_val.isdigit():
                    norm_k = self._normalize_key(clean_name)
                    registry[norm_k] = {
                        "type_id": type_val,
                        "format": fmt_val,
                        "clause": ref_val,
                        "length": "Variable" if "TLV" in fmt_val else ("1" if type_val == "14" else "4"),
                    }
                    if "gsn" in norm_k:
                        registry["gtpupeeraddress"] = registry[norm_k]

    def _parse_extension_header_registry(self, tbl_elem: Any, registry: Dict[str, Dict[str, str]]):
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return

        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) >= 2:
                code_bits = cells[0].replace(" ", "")
                name_val = re.sub(r"(?i)\s*see\s+note.*$", "", cells[1]).strip()

                if re.match(r"^[01]{8}$", code_bits) and name_val and "reserved" not in name_val.lower() and "no more" not in name_val.lower():
                    hex_val = f"0x{int(code_bits, 2):02X}"
                    norm_k = self._normalize_key(name_val)
                    if norm_k not in registry:
                        registry[norm_k] = {
                            "iei": hex_val,
                            "name": name_val,
                            "clause": "5.2.2",
                        }

    def _sanitize_message_name(self, caption: str) -> str:
        s = caption.strip().lstrip(": \t")
        s = re.sub(r"^Information\s+Elements\s+in\s+(?:an?\s+|the\s+)?", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+message\s+content$", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+message$", "", s, flags=re.IGNORECASE)
        return s.strip().lstrip(": \t")

    def _parse_clause_7_table(
        self,
        tbl_elem: Any,
        type_registry: Dict[str, Dict[str, str]],
        clause_ref: str,
        cap_name: str,
    ) -> List[Dict[str, Any]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        clean_msg = self._sanitize_message_name(cap_name)
        msg_norm = self._normalize_key(clean_msg)
        msg_appl = GTPU_MESSAGE_INTERFACES.get(msg_norm, "All Interfaces")

        ies: List[Dict[str, Any]] = []
        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) < 2:
                continue

            ie_name = cells[0]
            pres = cells[1].split()[0] if len(cells) > 1 and cells[1] else "O"
            ref = cells[2] if len(cells) > 2 else ""

            if not ie_name or any(kw in ie_name.lower() for kw in ("information element", "presence")):
                continue

            norm_name = self._normalize_key(ie_name)
            reg_info = type_registry.get(norm_name, {})
            iei_val = reg_info.get("type_id", "")
            fmt_val = reg_info.get("format", "TLV")
            len_val = reg_info.get("length", "Variable")
            type_ref = f"{ie_name} (Clause {ref})" if ref else ie_name

            ies.append({
                "iei": iei_val,
                "name": ie_name,
                "field": ie_name,
                "information_element": ie_name,
                "field_path": ie_name,
                "depth": 0,
                "type": type_ref,
                "type_reference": type_ref,
                "presence": pres,
                "format": fmt_val,
                "length": len_val,
                "applicability": msg_appl,
            })

        return ies

    def _build_gpdu_message(self, ext_reg: Dict[str, Dict[str, str]]) -> Dict[str, Any]:
        """Constructs a synthesized G-PDU message with T-PDU payload and all extension headers."""
        gpdu_appl = GTPU_MESSAGE_INTERFACES.get("gpdu", "All Interfaces")
        ies: List[Dict[str, Any]] = [
            {
                "iei": "-",
                "name": "T-PDU Payload",
                "field": "T-PDU Payload",
                "information_element": "T-PDU Payload",
                "field_path": "T-PDU Payload",
                "depth": 0,
                "type": "User Datagram / Frame (Clause 4.4.0)",
                "type_reference": "Clause 4.4.0",
                "presence": "M",
                "format": "Payload",
                "length": "Variable",
                "applicability": gpdu_appl,
            }
        ]

        # Order Extension Headers chronologically / logically
        ext_order = [
            ("udpport", "UDP Port", "Clause 5.2.2.1"),
            ("pdcppdunumber", "PDCP PDU Number", "Clause 5.2.2.2"),
            ("longpdcppdunumber", "Long PDCP PDU Number", "Clause 5.2.2.2A"),
            ("serviceclassindicator", "Service Class Indicator", "Clause 5.2.2.3"),
            ("rancontainer", "RAN Container", "Clause 5.2.2.4"),
            ("xwrancontainer", "Xw RAN Container", "Clause 5.2.2.5"),
            ("nrrancontainer", "NR RAN Container", "Clause 5.2.2.6"),
            ("pdusessioncontainer", "PDU Session Container", "Clause 5.2.2.7"),
            ("pdusetinformationcontainer", "PDU Set Information Container", "Clause 5.2.2.8"),
        ]

        for key, display_name, clause_ref in ext_order:
            reg_info = ext_reg.get(key, {})
            iei_val = reg_info.get("iei", "-")
            appl = GTPU_EXT_HEADER_INTERFACES.get(key, gpdu_appl)

            ies.append({
                "iei": iei_val,
                "name": display_name,
                "field": display_name,
                "information_element": display_name,
                "field_path": display_name,
                "depth": 0,
                "type": f"{display_name} ({clause_ref})",
                "type_reference": clause_ref,
                "presence": "O",
                "format": "Ext Header",
                "length": "4n octets",
                "applicability": appl,
            })

        return {
            "clause": "5.1",
            "clause_ref": "5.1",
            "name": "G-PDU",
            "message_name": "G-PDU",
            "table_caption": "G-PDU and Extension Headers (Clause 5)",
            "ies": ies,
            "fields": ies,
        }