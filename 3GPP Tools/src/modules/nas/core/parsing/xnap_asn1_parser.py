import json
import logging
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any

from modules.nas.core.parsing.ran3_asn1_parser import RAN3Asn1Parser
from modules.nas.core.parsing.protocol_parser_constants import RE_ELEM_PROC_MSG


class XnAPAsn1Parser(RAN3Asn1Parser):
    """
    Dedicated parser for 3GPP TS 38.423 (XnAP - Xn Application Protocol).

    Structure of TS 38.423:
      - Clause 9.1: Message Functional Definition and Content (Messages for Basic Mobility,
        Dual Connectivity, Global Procedures, IAB Procedures, and L1/L2 Triggered Mobility).
      - Clause 9.2: Information Element Definitions (Radio Network Layer, RRM, Other IEs).
      - Clause 9.3: Message and Information Element Abstract Syntax (with ASN.1).
        - 9.3.2: Elementary Procedure Definitions
        - 9.3.3: PDU Definitions
        - 9.3.4: Information Element Definitions
        - 9.3.5: Common Extension Definitions
        - 9.3.6: Constant Definitions
        - 9.3.7: Container Definitions

    Unlike TS 38.413 (NGAP) where messages are under Clause 9.2, in TS 38.423:
      - Elementary procedure messages belong to Clause 9.1.
      - Types under Clause 9.2 and Clause 9.3 are Information Elements and must not
        be misclassified as messages.
    """

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.423"):
        super().__init__(docx_paths, spec_number=spec_number)
        self.logger = logging.getLogger(__name__)

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        full_asn1_module, field_desc_tables, field_individual_descs, clause_map = self._scan_xml_documents(
            progress_callback=progress_callback
        )

        if progress_callback:
            progress_callback("Parsing XnAP Information Object Sets and ProtocolIE Constants...", 65)

        # Alias acronym-named elementary procedures in clause_map and field_desc_tables
        # e.g., CPCCancel corresponds to Clause 9.1.2.31 CONDITIONAL PSCELL CHANGE CANCEL
        cpc_clause = (
            clause_map.get("conditionalpscellchangecancel")
            or clause_map.get("conditional pscell change cancel")
            or "9.1.2.31"
        )
        clause_map["cpccancel"] = cpc_clause
        cpc_html = (
            field_desc_tables.get("conditionalpscellchangecancel")
            or field_desc_tables.get("conditional pscell change cancel")
            or ""
        )
        if cpc_html:
            field_desc_tables["cpccancel"] = cpc_html

        id_map = self._extract_constant_ids(full_asn1_module)
        object_sets = self._extract_object_sets(full_asn1_module, id_map)
        type_defs = self._extract_asn1_type_definitions(full_asn1_module)
        known_pdu_messages = set(RE_ELEM_PROC_MSG.findall(full_asn1_module))
        elem_procs = self._extract_elementary_procedures(full_asn1_module, id_map)

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []
        ie_def_by_name: Dict[str, Dict[str, Any]] = {}

        # In TS 38.423, MessageType is defined in Clause 9.2.3.1 (Other IEs)
        msg_type_clause = (
            clause_map.get("messagetype")
            or clause_map.get("message type")
            or "9.2.3.1"
        )
        msg_type_desc_html = (
            field_desc_tables.get("messagetype")
            or field_desc_tables.get("message type")
            or ""
        )
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
                or ("9.1" if "Message" in clean_name else "9.2")
            )

            # In XnAP, messages are strictly elementary procedures or Clause 9.1 definitions.
            # Clause 9.2 and Clause 9.3 types are Information Elements.
            if known_pdu_messages:
                is_message = (
                    clean_name in known_pdu_messages
                    or clean_name == "PrivateMessage"
                    or str(assigned_clause).startswith("9.1.")
                ) and not clean_name.endswith("-IEs")
            else:
                is_message = (
                    clean_name.endswith((
                        "Request", "Response", "Command", "Complete", "Failure",
                        "Acknowledge", "Indication", "Confirm", "Notify", "Report",
                        "Transfer", "Required", "Start", "Stop"
                    ))
                    or clean_name in {
                        "RANPaging", "EarlyStatusTransfer", "SNStatusTransfer",
                        "CPCCancel", "RRCTransfer", "CellTrafficTrace", "TraceStart",
                        "DeactivateTrace", "RANMulticastGroupPaging"
                    }
                    or str(assigned_clause).startswith("9.1.")
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
                    # Prepend MessageType (Clause 9.2.3.1) for Clause 9.1 Elementary Procedure messages
                    is_elem_proc_msg = (
                        clean_name in known_pdu_messages
                        or clean_name == "PrivateMessage"
                        or str(assigned_clause).startswith("9.1.")
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

                    msg_desc, msg_dir = self._extract_prose_summary_and_direction(desc_table_html)
                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} (Clause {assigned_clause})",
                        "description": msg_desc,
                        "direction": msg_dir,
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
