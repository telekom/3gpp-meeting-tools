import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any, Set

from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.protocol_parser_constants import (
    RE_SETUP_RELEASE, RE_SEQ_OF, RE_OCTET_CONTAINING,
    RE_CRITICAL_EXT_IES, RE_STRIP_EXTRANEOUS, RE_COND_NEED
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
            norm_key = self._normalize_key(clean_name)
            assigned_clause = (
                clause_map.get(norm_key)
                or clause_map.get(clean_name.lower())
                or type_info.get("clause")
                or ("6.2" if "Message" in clean_name or clean_name.startswith("RRC") else "6.3")
            )

            is_message = self._is_rrc_message(clean_name)

            desc_table_html = (
                field_desc_tables.get(norm_key)
                or field_desc_tables.get(clean_name.lower())
                or field_desc_tables.get(f"{clean_name.lower()}-ies")
                or field_desc_tables.get(f"{clean_name.lower()}ies")
                or field_desc_tables.get(self._normalize_key(f"{clean_name}-IEs"))
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
                unrolled = self._unroll_rrc_fields(
                    root_name=clean_name,
                    type_defs=type_defs,
                    field_individual_descs=field_individual_descs,
                )
                if unrolled:
                    messages.append({
                        "clause": assigned_clause,
                        "message_name": clean_name,
                        "table_caption": f"{clean_name} ASN.1 PDU Definition (Clause {assigned_clause})",
                        "ies": unrolled,
                    })

        return messages, ie_definitions

    @staticmethod
    def _is_rrc_message(name: str) -> bool:
        """Determines if an ASN.1 type represents a standalone RRC message or PDU."""
        if name.endswith(("-IEs", "IEs")) or re.search(r"-v\d+[a-z]?-IEs$", name):
            return False

        if name.startswith(("RRC", "SIB", "SystemInformation", "MasterInformationBlock")):
            return True

        if name.endswith((
            "Request", "Response", "Command", "Complete", "Failure",
            "Reject", "Reestablishment", "Reconfiguration", "Release", "Resume"
        )):
            return True

        known_rrc_pdus = {
            "CellGroupConfig", "RadioBearerConfig", "MeasConfig",
            "ServingCellConfig", "ServingCellConfigCommon", "UE-NR-Capability",
            "UE-MRDC-Capability", "UE-CapabilityRAT-ContainerList", "VarMeasConfig",
            "VarConditionalReconfig", "HandoverCommand", "HandoverPreparationInformation"
        }
        return name in known_rrc_pdus

    def _unwrap_clean_type(self, ftype: str) -> str:
        """Unwraps parameterized RRC types such as SetupRelease, OCTET STRING CONTAINING, and SEQUENCE OF."""
        s = ftype.strip()

        sr_match = RE_SETUP_RELEASE.search(s)
        if sr_match:
            return sr_match.group(1).strip()

        oc_match = RE_OCTET_CONTAINING.search(s)
        if oc_match:
            return oc_match.group(1).strip()

        so_match = RE_SEQ_OF.search(s)
        if so_match:
            return so_match.group(1).strip()

        clean = RE_STRIP_EXTRANEOUS.sub("", s).strip()
        return clean

    def _extract_critical_extension_targets(self, ftype: str, type_defs: Dict[str, Dict[str, Any]]) -> List[str]:
        """Extracts valid target message IE structures from criticalExtensions CHOICE definitions."""
        found_ies = RE_CRITICAL_EXT_IES.findall(ftype)
        if found_ies:
            return [t for t in found_ies if t in type_defs]

        targets: List[str] = []
        tokens = re.findall(r"\b([A-Z][A-Za-z0-9\-]+)\b", ftype)
        for tok in tokens:
            if tok in type_defs and tok not in ("NULL", "SEQUENCE", "CHOICE", "BOOLEAN", "INTEGER"):
                if not tok.startswith("criticalExtensionsFuture"):
                    targets.append(tok)
        return targets

    def _unroll_rrc_fields(
            self,
            root_name: str,
            type_defs: Dict[str, Dict[str, Any]],
            field_individual_descs: Dict[str, Dict[str, str]],
            max_depth: int = 3,
    ) -> List[Dict[str, Any]]:
        """
        Recursively unrolls RRC message sequences, criticalExtensions choices,
        and versioned nonCriticalExtension chains across all 3GPP releases.
        """
        result: List[Dict[str, Any]] = []

        def _get_description(type_ctx: str, field_name: str, type_name: str) -> str:
            norm_ctx = self._normalize_key(type_ctx)
            norm_field = self._normalize_key(field_name)
            norm_type = self._normalize_key(type_name)

            descs = (
                field_individual_descs.get(norm_ctx)
                or field_individual_descs.get(type_ctx.lower())
                or field_individual_descs.get(self._normalize_key(root_name))
                or {}
            )

            return (
                descs.get(norm_field)
                or descs.get(field_name.lower())
                or descs.get(norm_type)
                or ""
            )

        def _unroll_type(current_type: str, path_prefix: str, depth: int, visited: Set[str]):
            if depth > max_depth or current_type in visited:
                return

            t_info = type_defs.get(current_type)
            if not t_info or not t_info.get("fields"):
                return

            new_visited = visited | {current_type}

            for f in t_info["fields"]:
                fname = f["name"]
                ftype = f["type"]
                presence = f.get("presence", "O")
                fmt = f.get("format", "SEQUENCE")

                # Parse RRC Condition / Need Codes from raw comments
                need_match = RE_COND_NEED.search(ftype)
                if need_match:
                    code = need_match.group(1).strip()
                    presence = f"C ({code})" if "Cond" in code else f"O ({code})"

                clean_type = self._unwrap_clean_type(ftype)
                full_path = f"{path_prefix}.{fname}" if path_prefix else fname
                field_desc = _get_description(current_type, fname, clean_type)

                # 1. Unroll criticalExtensions CHOICE envelope
                if fname == "criticalExtensions" or "criticalExtensions" in fname:
                    crit_targets = self._extract_critical_extension_targets(ftype, type_defs)
                    if crit_targets:
                        for target_ie in crit_targets:
                            _unroll_type(target_ie, path_prefix, depth, new_visited)
                        continue

                # 2. Chain horizontal nonCriticalExtensions across releases seamlessly
                if fname == "nonCriticalExtension":
                    if clean_type in type_defs:
                        _unroll_type(clean_type, path_prefix, depth, new_visited)
                    continue

                # 3. Unroll late nonCriticalExtension
                if fname == "lateNonCriticalExtension":
                    if clean_type in type_defs:
                        _unroll_type(clean_type, full_path, depth + 1, new_visited)
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

                # 4. Recursively unroll complex child types
                if (
                    clean_type in type_defs
                    and depth < max_depth
                    and not clean_type.endswith("List")
                    and clean_type != current_type
                ):
                    child_info = type_defs[clean_type]
                    if child_info.get("fields"):
                        _unroll_type(clean_type, full_path, depth + 1, new_visited)

        _unroll_type(root_name, "", 0, set())

        # Fallback for message definitions referencing external -IEs
        if not result:
            candidate_ies = f"{root_name}-IEs"
            if candidate_ies in type_defs:
                _unroll_type(candidate_ies, "", 0, set())

        return result