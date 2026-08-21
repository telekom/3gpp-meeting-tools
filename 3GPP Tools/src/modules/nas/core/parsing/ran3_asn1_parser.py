import json
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any, Set

from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.protocol_parser_constants import (
    RE_IE_ID_CONST, RE_OBJECT_SET_ITEM, RE_CONTAINER_REF, RE_ELEM_PROC_MSG,
    RE_SETUP_RELEASE, RE_SEQ_OF, RE_OCTET_CONTAINING, RE_STRIP_EXTRANEOUS
)


class RAN3Asn1Parser(BaseAsn1DocxParser):
    """Dedicated parser for 3GPP RAN3 Specifications using Information Object Sets (NGAP, XnAP, S1AP, F1AP)."""

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        full_asn1_module, field_desc_tables, field_individual_descs, clause_map = self._scan_xml_documents(
            progress_callback=progress_callback
        )

        if progress_callback:
            progress_callback("Parsing RAN3 Information Object Sets and ProtocolIE Constants...", 65)

        id_map = self._extract_constant_ids(full_asn1_module)
        object_sets = self._extract_object_sets(full_asn1_module, id_map)
        type_defs = self._extract_asn1_type_definitions(full_asn1_module)
        known_pdu_messages = set(RE_ELEM_PROC_MSG.findall(full_asn1_module))

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        for type_name, type_info in type_defs.items():
            clean_name = type_name.strip()
            assigned_clause = clause_map.get(clean_name.lower()) or type_info.get("clause") or (
                "9.2" if "Message" in clean_name else "9.3"
            )

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

            desc_table_html = (
                field_desc_tables.get(clean_name.lower())
                or field_desc_tables.get(f"{clean_name.lower()}-ies")
                or field_desc_tables.get(f"{clean_name.lower()}ies")
                or ""
            )

            inspector_html = self._build_inspector_html(
                type_name=clean_name,
                assigned_clause=assigned_clause,
                raw_asn1=type_info["raw_asn1"],
                desc_table_html=desc_table_html,
            )

            ie_definitions.append({
                "clause": assigned_clause,
                "ie_name": clean_name,
                "raw_description": inspector_html,
                "structure_table": json.dumps([]),
            })

            if is_message:
                unrolled = self._unroll_ran3_fields(
                    root_name=clean_name,
                    type_defs=type_defs,
                    object_sets=object_sets,
                    field_individual_descs=field_individual_descs,
                )
                if unrolled:
                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} ASN.1 PDU Definition",
                        "ies": unrolled,
                    })

        return messages, ie_definitions

    def _extract_constant_ids(self, asn1_text: str) -> Dict[str, str]:
        """Extracts ProtocolIE-ID integer constants (e.g. id-AMF-UE-NGAP-ID ::= 10)."""
        id_map: Dict[str, str] = {}
        for match in RE_IE_ID_CONST.finditer(asn1_text):
            id_map[match.group(1).strip()] = match.group(2).strip()
        return id_map

    def _extract_object_sets(self, asn1_text: str, id_map: Dict[str, str]) -> Dict[str, List[Dict[str, Any]]]:
        """
        Extracts Information Object Sets using a deterministic bracket scanner
        to eliminate regex catastrophic backtracking on large specifications.
        """
        object_sets: Dict[str, List[Dict[str, Any]]] = {}
        clean_lines = [line.split("--")[0] for line in asn1_text.splitlines()]
        cleaned_text = "\n".join(clean_lines)

        pos = 0
        while True:
            assign_idx = cleaned_text.find("::=", pos)
            if assign_idx == -1:
                break

            header_part = cleaned_text[max(0, assign_idx - 140):assign_idx].strip()
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

    def _unroll_ran3_fields(
            self,
            root_name: str,
            type_defs: Dict[str, Dict[str, Any]],
            object_sets: Dict[str, List[Dict[str, Any]]],
            field_individual_descs: Dict[str, Dict[str, str]],
            max_depth: int = 4,
    ) -> List[Dict[str, Any]]:
        """Recursively unrolls RAN3 Information Object Sets and resolves ProtocolIE-Containers."""
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

            # Resolve ProtocolIE-Container wrapping
            container_match = RE_CONTAINER_REF.search(ftype)
            if container_match:
                set_name = container_match.group(1).strip()
                if set_name in object_sets:
                    for set_field in object_sets[set_name]:
                        _process_field(set_field, path_prefix, depth, visited_set, descs_map)
                    return

            clean_type = RE_SETUP_RELEASE.sub(r"\1", ftype)
            clean_type = RE_SEQ_OF.sub(r"\1", clean_type)
            clean_type = RE_OCTET_CONTAINING.sub(r"\1", clean_type)
            clean_type = RE_STRIP_EXTRANEOUS.sub("", clean_type).strip()

            full_path = f"{path_prefix}.{fname}" if path_prefix else fname
            field_desc = descs_map.get(fname.lower()) or descs_map.get(clean_type.lower(), "")

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

            if clean_type in object_sets and depth < max_depth:
                recurse(clean_type, full_path, depth + 1, visited_set)
            elif clean_type in type_defs and depth < max_depth:
                target_t_info = type_defs[clean_type]
                if target_t_info.get("fields") or RE_CONTAINER_REF.search(target_t_info.get("raw_asn1", "")):
                    recurse(clean_type, full_path, depth + 1, visited_set)

        recurse(root_name, "", 0, set())

        if not result:
            if f"{root_name}-IEs" in object_sets:
                recurse(f"{root_name}-IEs", "", 0, set())
            elif f"{root_name}IEs" in object_sets:
                recurse(f"{root_name}IEs", "", 0, set())

        return result