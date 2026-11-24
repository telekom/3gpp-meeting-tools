"""IE finalization and encapsulated NAS container sub-IE unrolling helpers."""

from __future__ import annotations

import json
from typing import Any, Dict, List

from modules.nas.core.parsing.nas.catalogs import (
    EPCO_SUB_IES,
    PAYLOAD_CONTAINER_SUB_IES,
    SERVICE_LEVEL_AA_SUB_IES,
    STRUCTURAL_SECTION_CLAUSES,
    UE_POLICY_PART_SUB_IES,
)


def finalize_ie_definition(
    current_ie_def: Dict[str, Any],
    ie_definitions: List[Dict[str, Any]],
    is_policy_spec: bool,
) -> None:
    """Validate and append a parsed Clause 9 / Annex D IE definition."""
    clause = current_ie_def.get("clause", "").strip()
    ie_name = current_ie_def.get("ie_name", "").strip()
    html_blocks = current_ie_def.get("html_content", [])

    # Skip umbrella section headers, Void clauses, or empty heading stubs
    if not clause or not ie_name or len(html_blocks) <= 1:
        return
    if is_policy_spec:
        if clause in ("5.1", "12.1"):
            return
    else:
        if clause in STRUCTURAL_SECTION_CLAUSES:
            return
    lower_name = ie_name.lower()
    if lower_name.startswith("void") or lower_name in ("general", "overview"):
        return
    if not is_policy_spec and lower_name == "information elements coding":
        return

    ie_definitions.append(
        {
            "clause": clause,
            "ie_name": ie_name,
            "raw_description": "".join(html_blocks),
            "structure_table": json.dumps([]),
        }
    )


def unroll_container_sub_ies(
    messages: List[Dict[str, Any]],
    ie_definitions: List[Dict[str, Any]],
) -> None:
    """
    Unrolls encapsulated NAS container payloads (Payload container, Service-level-AA container,
    Type 6 IE container, ePCO) into depth > 0 child rows when the target Clause 9 / Annex D
    IE definitions exist in the parsed specification version.
    """
    defs_by_clause: Dict[str, Dict[str, Any]] = {
        d["clause"]: d for d in ie_definitions if d.get("clause")
    }
    if not defs_by_clause:
        return

    # Find helper table 8.2.7.54.1.1 (Registration accept type 6 IE container contents) if present
    type6_sub_ies: List[Dict[str, Any]] = []
    for m in messages:
        if m.get("clause", "").startswith("8.2.7.54"):
            type6_sub_ies = m.get("ies", [])
            break

    def _build_slaa_children(parent_path: str, base_depth: int) -> List[Dict[str, Any]]:
        children = []
        for sub_iei, fallback_name, sub_clause, pres, fmt, length in SERVICE_LEVEL_AA_SUB_IES:
            if sub_clause in defs_by_clause:
                resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                children.append(
                    {
                        "iei": sub_iei,
                        "information_element": resolved_name,
                        "field_path": f"{parent_path} > {resolved_name}",
                        "depth": base_depth,
                        "type_reference": f"{resolved_name} {sub_clause}",
                        "presence": pres,
                        "format": fmt,
                        "length": length,
                    }
                )
        return children

    def _build_ue_policy_children(parent_path: str, base_depth: int) -> List[Dict[str, Any]]:
        children = []
        for sub_iei, part_name, sub_ref, pres, fmt, length in UE_POLICY_PART_SUB_IES:
            children.append(
                {
                    "iei": sub_iei,
                    "information_element": part_name,
                    "field_path": f"{parent_path} > {part_name}",
                    "depth": base_depth,
                    "type_reference": sub_ref,
                    "presence": pres,
                    "format": fmt,
                    "length": length,
                }
            )
        return children

    for msg in messages:
        if msg.get("clause", "").startswith("8.2.7.54"):
            continue

        expanded_ies: List[Dict[str, Any]] = []
        for ie in msg.get("ies", []):
            expanded_ies.append(ie)
            type_ref = ie.get("type_reference", "")
            parent_name = ie.get("information_element", "")

            # 1. Payload container (Clause 9.11.3.39)
            if "9.11.3.39" in type_ref:
                for sub_iei, fallback_name, sub_clause, pres, fmt, length in PAYLOAD_CONTAINER_SUB_IES:
                    if sub_clause in defs_by_clause:
                        resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                        sub_path = f"{parent_name} > {resolved_name}"
                        expanded_ies.append(
                            {
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": sub_path,
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            }
                        )
                        if sub_clause == "9.11.2.10":
                            expanded_ies.extend(_build_slaa_children(sub_path, base_depth=2))
                        elif sub_clause == "D.6.2":
                            expanded_ies.extend(_build_ue_policy_children(sub_path, base_depth=2))

            # 2. Top-level Service-level-AA container (Clause 9.11.2.10)
            elif "9.11.2.10" in type_ref:
                expanded_ies.extend(_build_slaa_children(parent_name, base_depth=1))

            # 3. Top-level UE policy section management list (Clause D.6.2)
            elif "D.6.2" in type_ref and "D.6.2" in defs_by_clause:
                expanded_ies.extend(_build_ue_policy_children(parent_name, base_depth=1))

            # 4. Registration accept type 6 IE container (Clause 9.11.3.98)
            elif "9.11.3.98" in type_ref and type6_sub_ies:
                for sub in type6_sub_ies:
                    sub_name = sub.get("information_element", "")
                    expanded_ies.append(
                        {
                            "iei": sub.get("iei", ""),
                            "information_element": sub_name,
                            "field_path": f"{parent_name} > {sub_name}",
                            "depth": 1,
                            "type_reference": sub.get("type_reference", ""),
                            "presence": sub.get("presence", ""),
                            "format": sub.get("format", ""),
                            "length": sub.get("length", ""),
                        }
                    )

            # 5. Extended protocol configuration options (Clause 9.11.4.6)
            elif "9.11.4.6" in type_ref:
                for sub_iei, fallback_name, sub_clause, pres, fmt, length in EPCO_SUB_IES:
                    if sub_clause in defs_by_clause:
                        resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                        expanded_ies.append(
                            {
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": f"{parent_name} > {resolved_name}",
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            }
                        )

        msg["ies"] = expanded_ies
