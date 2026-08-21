import json
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any, Set

from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.protocol_parser_constants import (
    RE_SETUP_RELEASE, RE_SEQ_OF, RE_OCTET_CONTAINING, RE_STRIP_EXTRANEOUS
)


class RRCAsn1Parser(BaseAsn1DocxParser):
    """Dedicated parser for 3GPP RAN2 RRC Specifications (TS 38.331 and TS 36.331)."""

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        full_asn1_module, field_desc_tables, field_individual_descs, clause_map = self._scan_xml_documents(
            progress_callback=progress_callback
        )

        if progress_callback:
            progress_callback("Parsing RRC ASN.1 structures and building evolution records...", 65)

        type_defs = self._extract_asn1_type_definitions(full_asn1_module)
        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        for type_name, type_info in type_defs.items():
            clean_name = type_name.strip()
            assigned_clause = clause_map.get(clean_name.lower()) or type_info.get("clause") or (
                "6.2" if "Message" in clean_name else "6.3"
            )

            is_message = (
                clean_name.startswith(("RRC", "SIB", "SystemInformation"))
                or clean_name.endswith(("Request", "Response", "Command", "Complete", "Failure"))
                or clean_name in {
                    "CellGroupConfig", "RadioBearerConfig", "MeasConfig",
                    "ServingCellConfig", "ServingCellConfigCommon", "UE-NR-Capability"
                }
            ) and not clean_name.endswith("-IEs") and not re.search(r"-v\d+[a-z]?-IEs$", clean_name)

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
                unrolled = self._unroll_rrc_fields(clean_name, type_defs, field_individual_descs)
                if unrolled:
                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} ASN.1 PDU Definition",
                        "ies": unrolled,
                    })

        return messages, ie_definitions

    def _unroll_rrc_fields(
            self,
            root_name: str,
            type_defs: Dict[str, Dict[str, Any]],
            field_individual_descs: Dict[str, Dict[str, str]],
            max_depth: int = 4,
    ) -> List[Dict[str, Any]]:
        """Recursively unrolls RRC sequence fields and chains non-critical extensions."""
        result: List[Dict[str, Any]] = []

        def recurse(current_type: str, path_prefix: str, depth: int, visited: Set[str]):
            if depth > max_depth or current_type in visited:
                return

            t_info = type_defs.get(current_type)
            if not t_info or not t_info["fields"]:
                return

            new_visited = visited | {current_type}
            descs_for_type = field_individual_descs.get(current_type.lower(), {})

            for f in t_info["fields"]:
                fname = f["name"]
                ftype = f["type"]
                presence = f["presence"]
                fmt = f["format"]

                clean_type = RE_SETUP_RELEASE.sub(r"\1", ftype)
                clean_type = RE_SEQ_OF.sub(r"\1", clean_type)
                clean_type = RE_OCTET_CONTAINING.sub(r"\1", clean_type)
                clean_type = RE_STRIP_EXTRANEOUS.sub("", clean_type).strip()

                full_path = f"{path_prefix}.{fname}" if path_prefix else fname
                field_desc = descs_for_type.get(fname.lower()) or descs_for_type.get(clean_type.lower(), "")

                if fname == "nonCriticalExtension" and clean_type in type_defs:
                    recurse(clean_type, path_prefix, depth, new_visited)
                    continue

                result.append({
                    "iei": "",
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

                if fname == "criticalExtensions" and clean_type in type_defs:
                    recurse(clean_type, full_path, depth, new_visited)
                elif clean_type in type_defs and depth < max_depth and not clean_type.endswith("List"):
                    recurse(clean_type, full_path, depth + 1, new_visited)

        recurse(root_name, "", 0, set())
        if not result and f"{root_name}-IEs" in type_defs:
            recurse(f"{root_name}-IEs", "", 0, set())

        return result