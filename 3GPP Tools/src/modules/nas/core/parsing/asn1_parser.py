import html
import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any, Set

from modules.nas.core.parsing.protocol_parser_utils import (
    extract_document_root, _extract_p_text, _extract_tc_text, _convert_table_to_html
)
from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY, TAG_P, TAG_TBL, TAG_TR, TAG_TC
)

# --- Pre-compiled Regular Expressions ---
RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_CLAUSE_HEADER = re.compile(r"^((?:6|9|D\.6)(?:\.[0-9A-Za-z]+)+)\s*(.*)$")
RE_DESC_TABLE = re.compile(r"([A-Za-z0-9\-_]+)\s+field\s+descriptions", re.IGNORECASE)

# Standard ASN.1 Type Declaration
RE_TYPE_DECL = re.compile(r"([A-Za-z0-9\-]+)(?:\s*\{[^}]*\})?\s*::=\s*", re.MULTILINE)
RE_TYPE_KIND = re.compile(r"^(SEQUENCE|CHOICE|ENUMERATED|BIT STRING|OCTET STRING|INTEGER|BOOLEAN)", re.IGNORECASE)
RE_FIELD_LINE = re.compile(r"^([A-Za-z0-9\-]+)\s+(.+)$")
RE_STRIP_KEYWORDS = re.compile(r"\s+(?:OPTIONAL|MANDATORY|DEFAULT\s+[^,\s]+).*", re.IGNORECASE)

# RAN3 Information Object Set (IOS) & ProtocolIE Constants Patterns
RE_IE_ID_CONST = re.compile(r"^(id-[A-Za-z0-9\-]+)\s+(?:ProtocolIE-ID|INTEGER)\s*::=\s*(\d+)", re.MULTILINE)
RE_OBJECT_SET_ITEM = re.compile(
    r"\{\s*ID\s+([A-Za-z0-9\-]+)\s+CRITICALITY\s+([A-Za-z0-9\-]+)\s+(?:TYPE|EXTENSION)\s+([A-Za-z0-9\-]+(?:\s*\{[^}]*\})?)\s+PRESENCE\s+([A-Za-z0-9\-]+)\s*\}",
    re.DOTALL
)
RE_CONTAINER_REF = re.compile(
    r"Protocol(?:IE|Extension)-(?:Container|SingleContainer|ContainerList|ContainerPair)\s*\{[^}]*\{\s*([A-Za-z0-9\-]+)\s*\}\s*\}",
    re.IGNORECASE
)
RE_ELEM_PROC_MSG = re.compile(
    r"(?:INITIATING MESSAGE|SUCCESSFUL OUTCOME|UNSUCCESSFUL OUTCOME)\s+([A-Za-z0-9\-]+)",
    re.MULTILINE
)

# ASN.1 Type unwrapping patterns
RE_SETUP_RELEASE = re.compile(r"^SetupRelease\s*\{\s*([A-Za-z0-9\-]+)\s*\}")
RE_SEQ_OF = re.compile(r"^SEQUENCE\s*(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\s+([A-Za-z0-9\-]+)")
RE_OCTET_CONTAINING = re.compile(r"^OCTET STRING\s*\(CONTAINING\s+([A-Za-z0-9\-]+)\)")
RE_STRIP_EXTRANEOUS = re.compile(r"[\(\{\[].*$")


class ASN1DocxParser:
    """
    Extracts ASN.1 Message definitions, Information Elements, Sequence Fields,
    and Information Object Sets (IOS) from 3GPP TS 38.331, TS 36.331, and TS 38.413 (NGAP).
    """

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.331"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        raw_asn1_blocks: List[str] = []
        field_desc_tables: Dict[str, str] = {}
        field_individual_descs: Dict[str, Dict[str, str]] = {}
        clause_map: Dict[str, str] = {}
        total_files = len(self.docx_paths)
        current_clause_num = "6.2"

        # 1. First Pass: Scan XML for ASN.1 blocks, Field Description Tables, and Headings
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

        if progress_callback:
            progress_callback("Parsing ASN.1 structures and Information Object Sets...", 60)

        # 2. Extract Type Definitions, Information Object Sets, and ProtocolIE-IDs
        id_map = self._extract_constant_ids(full_asn1_module)
        object_sets = self._extract_object_sets(full_asn1_module, id_map)
        type_defs = self._extract_asn1_type_definitions(full_asn1_module)

        # 3. Identify Top-Level PDUs / Elementary Procedures
        known_pdu_messages = set(RE_ELEM_PROC_MSG.findall(full_asn1_module))
        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        is_rrc = "331" in self.spec_number
        is_ngap = any(k in self.spec_number for k in ["413", "423", "412", "473", "463"])

        for type_name, type_info in type_defs.items():
            clean_name = type_name.strip()
            assigned_clause = clause_map.get(clean_name.lower()) or type_info.get("clause") or (
                "6.2" if "Message" in clean_name else "6.3"
            )

            is_message = False
            if is_rrc:
                is_message = (
                    clean_name.startswith(("RRC", "SIB", "SystemInformation"))
                    or clean_name.endswith(("Request", "Response", "Command", "Complete", "Failure"))
                    or clean_name in {
                        "CellGroupConfig", "RadioBearerConfig", "MeasConfig",
                        "ServingCellConfig", "ServingCellConfigCommon", "UE-NR-Capability"
                    }
                ) and not clean_name.endswith("-IEs") and not re.search(r"-v\d+[a-z]?-IEs$", clean_name)
            elif is_ngap:
                is_message = (
                    clean_name in known_pdu_messages
                    or clean_name.endswith((
                        "Request", "Response", "Command", "Complete", "Failure",
                        "Acknowledge", "Indication", "Confirm", "Notify", "Report",
                        "Transfer", "Required", "Start", "Stop"
                    ))
                    or clean_name in {
                        "Paging", "InitialUEMessage", "ErrorIndication", "HandoverNotify",
                        "CellTrafficTrace", "TraceStart", "DeactivateTrace", "LocationReport",
                        "LocationReportingControl", "RRCInactiveTransitionReport"
                    }
                ) and not clean_name.endswith("-IEs")

            raw_asn1_html = (
                f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; '
                f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A; '
                f'white-space: pre-wrap;">{html.escape(type_info["raw_asn1"])}</pre>'
            )

            desc_table_html = (
                field_desc_tables.get(clean_name.lower())
                or field_desc_tables.get(f"{clean_name.lower()}-ies")
                or field_desc_tables.get(f"{clean_name.lower()}ies")
                or ""
            )

            if desc_table_html:
                desc_table_html = (
                    f'<h3 style="font-size: 12px; font-weight: bold; color: #1E293B; margin-top: 10px; '
                    f'margin-bottom: 4px;">{clean_name} field descriptions</h3>{desc_table_html}'
                )

            full_inspector_html = (
                f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 6px; '
                f'font-family: Segoe UI, sans-serif;">{clean_name} (Clause {assigned_clause})</h2>'
                f'{raw_asn1_html}{desc_table_html}'
            )

            ie_definitions.append({
                "clause": assigned_clause,
                "ie_name": clean_name,
                "raw_description": full_inspector_html,
                "structure_table": json.dumps([]),
            })

            # Unroll fields for top-level messages
            if is_message:
                unrolled_fields = self._unroll_message_fields(
                    root_name=clean_name,
                    type_defs=type_defs,
                    object_sets=object_sets,
                    field_individual_descs=field_individual_descs,
                )
                if unrolled_fields:
                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} ASN.1 PDU Definition",
                        "ies": unrolled_fields,
                    })

        return messages, ie_definitions

    def _fallback_extract_asn1(self) -> str:
        collected = []
        for docx_path in self.docx_paths:
            root = extract_document_root(docx_path)
            if root is not None:
                body = root.find(TAG_BODY)
                if body is not None:
                    for p in body.findall(TAG_P):
                        t = _extract_p_text(p)
                        if any(kw in t for kw in ("::=", "SEQUENCE", "CHOICE", "NGAP-PROTOCOL-IES")):
                            collected.append(t)
        return "\n".join(collected)

    def _extract_constant_ids(self, asn1_text: str) -> Dict[str, str]:
        """Extracts ProtocolIE-ID integer constants (e.g. id-AMF-UE-NGAP-ID ::= 10)."""
        id_map: Dict[str, str] = {}
        for match in RE_IE_ID_CONST.finditer(asn1_text):
            id_name = match.group(1).strip()
            num_val = match.group(2).strip()
            id_map[id_name] = num_val
        return id_map

    def _extract_object_sets(self, asn1_text: str, id_map: Dict[str, str]) -> Dict[str, List[Dict[str, Any]]]:
        """
        Extracts Information Object Sets using a deterministic bracket scanner
        to eliminate regex catastrophic backtracking on large specifications.
        """
        object_sets: Dict[str, List[Dict[str, Any]]] = {}

        # Remove single-line comments while preserving structure
        clean_lines = [line.split("--")[0] for line in asn1_text.splitlines()]
        cleaned_text = "\n".join(clean_lines)

        pos = 0
        while True:
            assign_idx = cleaned_text.find("::=", pos)
            if assign_idx == -1:
                break

            # Find declaration header before ::=
            header_part = cleaned_text[max(0, assign_idx - 120):assign_idx].strip()
            tokens = header_part.split()

            set_name = ""
            is_object_set = False

            if len(tokens) >= 2:
                class_candidate = tokens[-1].upper()
                name_candidate = tokens[-2]

                if any(kw in class_candidate for kw in ("PROTOCOL-IES", "PROTOCOL-EXTENSION", "PROTOCOL-IES-PAIR", "PRIVATE-IES")):
                    set_name = name_candidate
                    is_object_set = True
                elif name_candidate.endswith(("-IEs", "IEs", "-Extensions", "Extensions")):
                    set_name = name_candidate
                    is_object_set = True

            # Locate opening brace after ::=
            brace_start = cleaned_text.find("{", assign_idx)
            next_assign = cleaned_text.find("::=", assign_idx + 3)

            if brace_start != -1 and (next_assign == -1 or brace_start < next_assign):
                depth = 0
                brace_end = -1
                for idx in range(brace_start, len(cleaned_text)):
                    ch = cleaned_text[idx]
                    if ch == "{":
                        depth += 1
                    elif ch == "}":
                        depth -= 1
                        if depth == 0:
                            brace_end = idx
                            break

                if brace_end > brace_start:
                    body_content = cleaned_text[brace_start + 1:brace_end]
                    if is_object_set or "ID " in body_content:
                        if not set_name and len(tokens) >= 1:
                            set_name = tokens[-1]

                        fields: List[Dict[str, Any]] = []
                        for item_match in RE_OBJECT_SET_ITEM.finditer(body_content):
                            ie_id_name = item_match.group(1).strip()
                            crit = item_match.group(2).strip()
                            ftype = item_match.group(3).strip()
                            pres_str = item_match.group(4).strip().lower()

                            presence = "M" if "mandatory" in pres_str else ("C" if "conditional" in pres_str else "O")
                            iei_val = id_map.get(ie_id_name, "")
                            clean_field_name = ie_id_name[3:] if ie_id_name.startswith("id-") else ie_id_name

                            fields.append({
                                "name": clean_field_name,
                                "type": ftype,
                                "presence": presence,
                                "criticality": crit,
                                "format": "IE",
                                "iei": iei_val,
                                "id_name": ie_id_name,
                            })

                        if fields and set_name:
                            object_sets[set_name] = fields

                    pos = brace_end + 1
                    continue

            pos = assign_idx + 3

        return object_sets

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

            # Clean module boundary leak if 'END' is encountered
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

    def _unroll_message_fields(
            self,
            root_name: str,
            type_defs: Dict[str, Dict[str, Any]],
            object_sets: Dict[str, List[Dict[str, Any]]],
            field_individual_descs: Dict[str, Dict[str, str]],
            max_depth: int = 4,
    ) -> List[Dict[str, Any]]:
        """
        Recursively unrolls sequence fields and Information Object Sets (IOS),
        resolving ProtocolIE-Containers and binding field descriptions.
        """
        result: List[Dict[str, Any]] = []

        def recurse(current_type: str, path_prefix: str, depth: int, visited: Set[str]):
            if depth > max_depth or current_type in visited:
                return

            new_visited = visited | {current_type}
            descs_for_type = field_individual_descs.get(current_type.lower(), {})

            # Case A: Direct Object Set (e.g. InitialUEMessage-IEs)
            if current_type in object_sets:
                for f in object_sets[current_type]:
                    _process_field(f, path_prefix, depth, new_visited, descs_for_type)
                return

            # Case B: Type Definition
            t_info = type_defs.get(current_type)
            if not t_info:
                return

            # Handle Container Alias (e.g. SEQUENCE OF ProtocolIE-SingleContainer { {Set} })
            raw_asn = t_info.get("raw_asn1", "")
            container_match = RE_CONTAINER_REF.search(raw_asn)
            if container_match and not t_info.get("fields"):
                set_name = container_match.group(1).strip()
                if set_name in object_sets:
                    for f in object_sets[set_name]:
                        _process_field(f, path_prefix, depth, new_visited, descs_for_type)
                    return

            for f in t_info.get("fields", []):
                _process_field(f, path_prefix, depth, new_visited, descs_for_type)

        def _process_field(f: Dict[str, Any], path_prefix: str, depth: int, visited_set: Set[str], descs_map: Dict[str, str]):
            fname = f["name"]
            ftype = f["type"]
            presence = f.get("presence", "O")
            fmt = f.get("format", "IE")
            iei = f.get("iei", "")

            # Check if ftype wraps an Object Set via ProtocolIE-Container
            container_match = RE_CONTAINER_REF.search(ftype)
            if container_match:
                set_name = container_match.group(1).strip()
                if set_name in object_sets:
                    for set_field in object_sets[set_name]:
                        _process_field(set_field, path_prefix, depth, visited_set, descs_map)
                    return

            # Extract clean type identifier
            clean_type = RE_SETUP_RELEASE.sub(r"\1", ftype)
            clean_type = RE_SEQ_OF.sub(r"\1", clean_type)
            clean_type = RE_OCTET_CONTAINING.sub(r"\1", clean_type)
            clean_type = RE_STRIP_EXTRANEOUS.sub("", clean_type).strip()

            full_path = f"{path_prefix}.{fname}" if path_prefix else fname
            field_desc = descs_map.get(fname.lower()) or descs_map.get(clean_type.lower(), "")

            if fname == "nonCriticalExtension" and clean_type in type_defs:
                recurse(clean_type, path_prefix, depth, visited_set)
                return

            result.append({
                "iei": iei,
                "information_element": fname,
                "field_path": full_path,
                "depth": depth,
                "type_reference": ftype,
                "clean_type": clean_type,
                "presence": presence,
                "format": fmt,
                "length": "-",
                "field_description": field_desc,
            })

            # Unroll nested criticalExtensions or sub-structures
            if fname == "criticalExtensions" and clean_type in type_defs:
                recurse(clean_type, full_path, depth, visited_set)
            elif clean_type in object_sets and depth < max_depth:
                recurse(clean_type, full_path, depth + 1, visited_set)
            elif clean_type in type_defs and depth < max_depth:
                target_t_info = type_defs[clean_type]
                if target_t_info.get("fields") or RE_CONTAINER_REF.search(target_t_info.get("raw_asn1", "")):
                    recurse(clean_type, full_path, depth + 1, visited_set)

        recurse(root_name, "", 0, set())

        # Fallback to <root_name>-IEs or <root_name>IEs if direct root returned empty
        if not result:
            if f"{root_name}-IEs" in object_sets:
                recurse(f"{root_name}-IEs", "", 0, set())
            elif f"{root_name}IEs" in object_sets:
                recurse(f"{root_name}IEs", "", 0, set())
            elif f"{root_name}-IEs" in type_defs:
                recurse(f"{root_name}-IEs", "", 0, set())

        return result


# Backward-compatibility alias
RRCAsn1DocxParser = ASN1DocxParser