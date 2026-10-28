import json
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any, Set

from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.protocol_parser_constants import (
    RE_IE_ID_CONST, RE_OBJECT_SET_ITEM, RE_CONTAINER_REF, RE_ELEM_PROC_MSG,
    RE_SETUP_RELEASE, RE_SEQ_OF, RE_OCTET_CONTAINING, RE_STRIP_EXTRANEOUS
)

RE_ELEM_PROC_BLOCK = re.compile(
    r"[A-Za-z0-9\-]+\s+[A-Za-z0-9\-]*ELEMENTARY-PROCEDURE\s*::=\s*\{([^}]+)\}",
    re.DOTALL | re.IGNORECASE,
)


class RAN3Asn1Parser(BaseAsn1DocxParser):
    """Dedicated parser for 3GPP RAN3 Specifications (NGAP, XnAP, S1AP, F1AP)."""

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
        elem_procs = self._extract_elementary_procedures(full_asn1_module, id_map)

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []
        ie_def_by_name: Dict[str, Dict[str, Any]] = {}

        # Synthesize Clause 9.3.1.1 MessageType definition if present in specification tables
        msg_type_clause = clause_map.get("messagetype") or clause_map.get("message type") or "9.3.1.1"
        msg_type_desc_html = field_desc_tables.get("messagetype") or field_desc_tables.get("message type") or ""
        if msg_type_desc_html or "ProcedureCode" in type_defs:
            msg_type_asn1_parts = [
                type_defs[k]["raw_asn1"]
                for k in ("ProcedureCode", "TriggeringMessage", "Criticality")
                if k in type_defs and type_defs[k].get("raw_asn1")
            ]
            msg_type_inspector_html = self._build_inspector_html(
                type_name="MessageType",
                assigned_clause=msg_type_clause,
                raw_asn1="\n\n".join(msg_type_asn1_parts),
                desc_table_html=msg_type_desc_html,
            )
            msg_type_record = {
                "clause": msg_type_clause,
                "ie_name": "MessageType",
                "raw_description": msg_type_inspector_html,
                "structure_table": json.dumps([]),
            }
            ie_definitions.append(msg_type_record)
            ie_def_by_name["MessageType"] = msg_type_record

        for type_name, type_info in type_defs.items():
            clean_name = type_name.strip()
            norm_key = self._normalize_key(clean_name)

            assigned_clause = (
                clause_map.get(norm_key)
                or clause_map.get(clean_name.lower())
                or type_info.get("clause")
                or ("9.2" if "Message" in clean_name else "9.3")
            )

            if known_pdu_messages:
                is_message = (
                    clean_name in known_pdu_messages
                    or clean_name == "PrivateMessage"
                    or str(assigned_clause).startswith(("9.2.", "9.3.4.", "9.3.5.", "9.3.6."))
                ) and not clean_name.endswith("-IEs")
            else:
                is_message = (
                    clean_name.endswith((
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
                field_desc_tables.get(norm_key)
                or field_desc_tables.get(clean_name.lower())
                or field_desc_tables.get(f"{norm_key}ies")
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

            ie_def_record = {
                "clause": assigned_clause,
                "ie_name": clean_name,
                "raw_description": inspector_html,
                "structure_table": json.dumps([]),
            }
            ie_definitions.append(ie_def_record)
            ie_def_by_name[clean_name] = ie_def_record

            if is_message:
                unrolled = self._unroll_ran3_fields(
                    root_name=clean_name,
                    type_defs=type_defs,
                    object_sets=object_sets,
                    field_individual_descs=field_individual_descs,
                )
                if unrolled:
                    # Prepend MessageType (Clause 9.3.1.1) for Clause 9.2 Elementary Procedure messages
                    is_elem_proc_msg = (
                        clean_name in known_pdu_messages
                        or clean_name == "PrivateMessage"
                        or str(assigned_clause).startswith("9.2.")
                    )
                    if is_elem_proc_msg:
                        proc_info = elem_procs.get(clean_name, {})
                        proc_code = proc_info.get("proc_code", "")
                        proc_crit = proc_info.get("criticality", "reject")
                        unrolled.insert(0, {
                            "iei": proc_code,
                            "information_element": "MessageType",
                            "field_path": "MessageType",
                            "depth": 0,
                            "type_reference": f"MessageType ({msg_type_clause})",
                            "clean_type": "MessageType",
                            "presence": "M",
                            "format": "IE",
                            "length": proc_crit,
                            "field_description": "The Message Type IE uniquely identifies the message being sent. It is mandatory for all messages.",
                        })

                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} (Clause {assigned_clause})",
                        "ies": unrolled,
                        "_desc_table_html": desc_table_html,
                    })

        # Backfill parent message clause & tabular specification for inline message-specific List/Item types
        for msg in messages:
            m_clause = msg.get("clause", "")
            m_name = msg.get("message_name", "")
            m_table_html = msg.pop("_desc_table_html", "")
            if m_clause and m_clause != m_name and m_table_html:
                for item in msg.get("ies", []):
                    ctype = item.get("clean_type", "")
                    target_def = ie_def_by_name.get(ctype)
                    if target_def and target_def["clause"] == ctype and ctype in type_defs:
                        target_def["clause"] = m_clause
                        target_def["raw_description"] = self._build_inspector_html(
                            type_name=ctype,
                            assigned_clause=m_clause,
                            raw_asn1=type_defs[ctype]["raw_asn1"],
                            desc_table_html=m_table_html,
                        )

        return messages, ie_definitions

    def _extract_constant_ids(self, asn1_text: str) -> Dict[str, str]:
        """Extracts ProtocolIE-ID and ProcedureCode integer constants (e.g. id-AMF-UE-NGAP-ID ::= 10)."""
        id_map: Dict[str, str] = {}
        for match in RE_IE_ID_CONST.finditer(asn1_text):
            id_map[match.group(1).strip()] = match.group(2).strip()
        return id_map

    def _extract_elementary_procedures(self, asn1_text: str, id_map: Dict[str, str]) -> Dict[str, Dict[str, str]]:
        """Maps Elementary Procedure messages to their ProcedureCode and Criticality."""
        proc_map: Dict[str, Dict[str, str]] = {}
        for match in RE_ELEM_PROC_BLOCK.finditer(asn1_text):
            body = match.group(1)
            code_match = re.search(r"PROCEDURE\s+CODE\s+([A-Za-z0-9\-]+)", body)
            crit_match = re.search(r"CRITICALITY\s+([A-Za-z0-9\-]+)", body)
            proc_id_name = code_match.group(1).strip() if code_match else ""
            proc_code = id_map.get(proc_id_name, "")
            crit = crit_match.group(1).strip() if crit_match else "reject"

            for msg_match in RE_ELEM_PROC_MSG.finditer(body):
                msg_name = msg_match.group(1).strip()
                proc_map[msg_name] = {
                    "proc_code": proc_code,
                    "criticality": crit,
                    "proc_id_name": proc_id_name,
                }
        return proc_map

    def _extract_object_sets(self, asn1_text: str, id_map: Dict[str, str]) -> Dict[str, List[Dict[str, Any]]]:
        """Extracts Information Object Sets using bracket scanning."""
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
        """Recursively unrolls RAN3 Information Object Sets, SEQUENCE OF lists, and ProtocolIE-Containers."""
        result: List[Dict[str, Any]] = []
        root_norm = self._normalize_key(root_name)
        root_descs = field_individual_descs.get(root_norm) or field_individual_descs.get(root_name.lower(), {})

        def recurse(current_type: str, path_prefix: str, depth: int, visited: Set[str]):
            if depth > max_depth or current_type in visited:
                return

            new_visited = visited | {current_type}
            norm_curr = self._normalize_key(current_type)
            descs_for_type = (
                field_individual_descs.get(norm_curr)
                or field_individual_descs.get(current_type.lower())
                or root_descs
            )

            if current_type in object_sets:
                for f in object_sets[current_type]:
                    _process_field(f, path_prefix, depth, new_visited, descs_for_type)
                return

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

            seq_of_match = RE_SEQ_OF.search(raw_asn)
            if seq_of_match and not t_info.get("fields"):
                item_type = seq_of_match.group(1).strip()
                seq_type_expr = raw_asn.split("::=", 1)[-1].strip() if "::=" in raw_asn else f"SEQUENCE OF {item_type}"
                item_field = {
                    "name": item_type,
                    "type": seq_type_expr,
                    "presence": "M",
                    "criticality": "-",
                    "format": "SEQUENCE OF",
                    "iei": "",
                }
                _process_field(item_field, path_prefix, depth, new_visited, descs_for_type)
                return

            for f in t_info.get("fields", []):
                _process_field(f, path_prefix, depth, new_visited, descs_for_type)

        def _process_field(f: Dict[str, Any], path_prefix: str, depth: int, visited_set: Set[str], descs_map: Dict[str, str]):
            fname = f["name"]
            ftype = f["type"].strip().rstrip(",").strip()
            presence = f.get("presence", "O")
            fmt = f.get("format", "IE")
            iei = f.get("iei", "")
            crit = f.get("criticality", "-") or "-"

            container_match = RE_CONTAINER_REF.search(ftype)
            if container_match:
                set_name = container_match.group(1).strip()
                if set_name in object_sets:
                    for set_field in object_sets[set_name]:
                        _process_field(set_field, path_prefix, depth, visited_set, descs_map)
                return

            is_octet_containing = bool(RE_OCTET_CONTAINING.search(ftype))

            clean_type = RE_SETUP_RELEASE.sub(r"\1", ftype)
            clean_type = RE_SEQ_OF.sub(r"\1", clean_type)
            clean_type = RE_OCTET_CONTAINING.sub(r"\1", clean_type)
            clean_type = RE_STRIP_EXTRANEOUS.sub("", clean_type).strip().rstrip(",").strip()

            full_path = f"{path_prefix}.{fname}" if path_prefix else fname
            norm_fname = self._normalize_key(fname)
            norm_ctype = self._normalize_key(clean_type)
            field_desc = (
                descs_map.get(norm_fname)
                or descs_map.get(fname.lower())
                or descs_map.get(norm_ctype)
                or root_descs.get(norm_fname)
                or root_descs.get(norm_ctype)
                or ""
            )

            result.append({
                "iei": iei,
                "information_element": fname,
                "field_path": full_path,
                "depth": depth,
                "type_reference": ftype,
                "clean_type": clean_type,
                "presence": presence,
                "format": fmt,
                "length": crit,
                "field_description": field_desc,
            })

            # Do not inline-expand OCTET STRING (CONTAINING ...) transparent containers inside the parent message tree
            if is_octet_containing:
                return

            if clean_type in object_sets and depth < max_depth:
                recurse(clean_type, full_path, depth + 1, visited_set)
            elif clean_type in type_defs and depth < max_depth:
                target_t_info = type_defs[clean_type]
                target_raw_asn = target_t_info.get("raw_asn1", "")
                if (
                    target_t_info.get("fields")
                    or RE_CONTAINER_REF.search(target_raw_asn)
                    or RE_SEQ_OF.search(target_raw_asn)
                ):
                    recurse(clean_type, full_path, depth + 1, visited_set)

        recurse(root_name, "", 0, set())

        if not result:
            if f"{root_name}-IEs" in object_sets:
                recurse(f"{root_name}-IEs", "", 0, set())
            elif f"{root_name}IEs" in object_sets:
                recurse(f"{root_name}IEs", "", 0, set())

        return result
