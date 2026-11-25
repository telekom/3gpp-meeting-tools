"""Message, common header, and hierarchical G-PDU builders for TS 29.281."""

from __future__ import annotations

import html
import json
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.nas.core.parsing.gtpu.catalogs import (
    GTPU_COMMON_HEADER_FIELDS,
    GTPU_CROSS_SPEC_CONTAINERS,
    GTPU_EXT_HEADER_DESCRIPTIONS,
    GTPU_EXT_HEADER_INTERFACES,
    GTPU_EXT_HEADER_ORDER,
    GTPU_MESSAGE_DEFAULT_CONTEXT,
    GTPU_MESSAGE_INTERFACES,
    GTPU_SIGNALLING_EXT_HEADERS,
)
from modules.nas.core.parsing.protocol_parser_constants import TAG_TC, TAG_TR
from modules.nas.core.parsing.protocol_parser_utils import _convert_table_to_html, _extract_tc_text


def section_header_html(title: str, clause: str) -> str:
    return (
        f'<h2 style="color: #0284C7; margin-top: 4px; margin-bottom: 6px; '
        f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
        f'{html.escape(title)} (Clause {html.escape(clause)})</h2>'
    )


def sanitize_message_name(caption: str) -> str:
    s = caption.strip().lstrip(": \t")
    s = re.sub(r"^Information\s+Elements\s+in\s+(?:an?\s+|the\s+)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+message\s+content$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+message$", "", s, flags=re.IGNORECASE)
    return s.strip().lstrip(": \t")


def make_ie_row(
    *,
    iei: str,
    name: str,
    field_path: str,
    depth: int,
    type_reference: str,
    presence: str,
    fmt: str,
    length: str,
    applicability: str,
    field_description: str = "",
) -> Dict[str, Any]:
    return {
        "iei": iei,
        "name": name,
        "field": name,
        "information_element": name,
        "field_path": field_path,
        "depth": depth,
        "type": type_reference,
        "type_reference": type_reference,
        "presence": presence,
        "format": fmt,
        "length": length,
        "applicability": applicability,
        "field_description": field_description,
    }


def build_ext_header_rows(
    key: str,
    display_name: str,
    clause_ref: str,
    ext_reg: Dict[str, Dict[str, Any]],
    parent_path: str,
    base_depth: int,
    presence: str = "O",
    applicability_override: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Build an extension header row plus any TS 38.415 cross-spec child frame rows."""
    reg_info = ext_reg.get(key, {})
    iei_val = reg_info.get("iei", "-")
    appl = applicability_override or GTPU_EXT_HEADER_INTERFACES.get(key, "All Interfaces")
    desc = GTPU_EXT_HEADER_DESCRIPTIONS.get(key, f"{display_name} ({clause_ref}).")
    row_path = f"{parent_path} > {display_name}" if parent_path else display_name

    rows = [
        make_ie_row(
            iei=iei_val,
            name=display_name,
            field_path=row_path,
            depth=base_depth,
            type_reference=f"{display_name} ({clause_ref})",
            presence=presence,
            fmt="Ext Header",
            length="4n octets",
            applicability=appl,
            field_description=desc,
        )
    ]
    for child in GTPU_CROSS_SPEC_CONTAINERS.get(key, []):
        rows.append(
            make_ie_row(
                iei=child["iei"],
                name=child["name"],
                field_path=f"{row_path} > {child['name']}",
                depth=base_depth + 1,
                type_reference=child["type_reference"],
                presence=child["presence"],
                fmt=child["format"],
                length=child["length"],
                applicability=child["applicability"],
                field_description=child["field_description"],
            )
        )
    return rows


def _parse_clause_7_table(
    tbl_elem: Any,
    type_registry: Dict[str, Dict[str, str]],
    clean_msg: str,
    msg_norm: str,
    normalize_key: Callable[[str], str],
) -> List[Dict[str, Any]]:
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return []
    msg_appl = GTPU_MESSAGE_INTERFACES.get(msg_norm, "All Interfaces")
    ies: List[Dict[str, Any]] = []
    for row in rows[1:]:
        cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
        if len(cells) < 2:
            continue
        ie_name = cells[0]
        pres = cells[1].split()[0] if cells[1] else "O"
        ref = cells[2] if len(cells) > 2 else ""
        if not ie_name or any(kw in ie_name.lower() for kw in ("information element", "presence")):
            continue

        reg_info = type_registry.get(normalize_key(ie_name), {})
        type_ref = f"{ie_name} (Clause {ref})" if ref else ie_name
        ies.append(
            make_ie_row(
                iei=reg_info.get("type_id", ""),
                name=ie_name,
                field_path=ie_name,
                depth=0,
                type_reference=type_ref,
                presence=pres,
                fmt=reg_info.get("format", "TLV"),
                length=reg_info.get("length", "Variable"),
                applicability=msg_appl,
                field_description=f"{ie_name} carried in {clean_msg} ({type_ref}).",
            )
        )
    return ies


def build_clause_7_messages(
    raw_tables: List[Tuple[str, str, str, Any]],
    type_registry: Dict[str, Dict[str, str]],
    ext_header_registry: Dict[str, Dict[str, Any]],
    clause_7_prose_map: Dict[str, List[str]],
    ie_definitions_dict: Dict[str, Dict[str, Any]],
    normalize_key: Callable[[str], str],
) -> List[Dict[str, Any]]:
    """Build Clause 7 signalling messages and attach applicable extension headers."""
    messages: List[Dict[str, Any]] = []
    for clause_ref, cap_name, full_cap, tbl_elem in raw_tables:
        if not clause_ref.startswith("7."):
            continue
        clean_name = sanitize_message_name(cap_name)
        norm_msg_k = normalize_key(clean_name)
        parsed_ies = _parse_clause_7_table(tbl_elem, type_registry, clean_name, norm_msg_k, normalize_key)
        if not parsed_ies:
            continue

        for ext_key, ext_pres, ext_appl in GTPU_SIGNALLING_EXT_HEADERS.get(norm_msg_k, []):
            reg_entry = ext_header_registry.get(ext_key, {})
            ext_name = reg_entry.get("name", ext_key)
            ext_clause = reg_entry.get("clause", "5.2.2")
            parsed_ies.extend(
                build_ext_header_rows(
                    key=ext_key,
                    display_name=ext_name,
                    clause_ref=f"Clause {ext_clause}",
                    ext_reg=ext_header_registry,
                    parent_path="",
                    base_depth=0,
                    presence=ext_pres,
                    applicability_override=ext_appl,
                )
            )

        base_clause = clause_ref.split("-")[0].strip()
        default_desc, default_dir = GTPU_MESSAGE_DEFAULT_CONTEXT.get(
            norm_msg_k,
            (f"GTPv1-U {clean_name} message (Clause {base_clause}).", "GTP-U Peer → GTP-U Peer"),
        )
        prose_lines = clause_7_prose_map.get(base_clause, [])
        msg_desc = " ".join(prose_lines).strip() if prose_lines else default_desc
        msg_dir = default_dir

        if norm_msg_k not in ie_definitions_dict:
            tbl_html = _convert_table_to_html(tbl_elem, is_figure_diagram=False)
            ie_definitions_dict[norm_msg_k] = {
                "clause": base_clause,
                "name": clean_name,
                "ie_name": clean_name,
                "raw_description": (
                    f"{section_header_html(clean_name, base_clause)}"
                    f'<p style="margin: 4px 0; line-height: 1.4; color: #0369A1; font-weight: bold; font-size: 11px;">'
                    f'Direction: {html.escape(msg_dir)}</p>'
                    f'<p style="margin: 4px 0; line-height: 1.4; color: #334155; font-size: 11px;">'
                    f'{html.escape(msg_desc)}</p>{tbl_html}'
                ),
                "structure_table": json.dumps([]),
            }

        messages.append(
            {
                "clause": base_clause,
                "clause_ref": clause_ref,
                "name": clean_name,
                "message_name": clean_name,
                "table_caption": full_cap,
                "description": msg_desc,
                "direction": msg_dir,
                "ies": parsed_ies,
                "fields": parsed_ies,
            }
        )
    return messages


def build_common_header_message(
    ie_definitions_dict: Dict[str, Dict[str, Any]], fig_511_tbl: Any
) -> Dict[str, Any]:
    """Synthesize the standalone GTP-U Header (Common) message from Clause 5.1."""
    desc, direction = GTPU_MESSAGE_DEFAULT_CONTEXT["gtpuheadercommon"]
    appl = GTPU_MESSAGE_INTERFACES["gtpuheadercommon"]
    ies = [
        make_ie_row(
            iei=f["iei"],
            name=f["name"],
            field_path=f["name"],
            depth=0,
            type_reference="GTP-U Header (Clause 5.1)",
            presence=f["presence"],
            fmt=f["format"],
            length=f["length"],
            applicability=appl,
            field_description=f["field_description"],
        )
        for f in GTPU_COMMON_HEADER_FIELDS
    ]
    fig_html = _convert_table_to_html(fig_511_tbl, is_figure_diagram=True) if fig_511_tbl is not None else ""
    ie_definitions_dict["gtpuheadercommon"] = {
        "clause": "5.1 (Header)",
        "name": "GTP-U Header (Common)",
        "ie_name": "GTP-U Header (Common)",
        "raw_description": (
            f"{section_header_html('GTP-U Header (Common)', '5.1')}"
            f'<p style="margin: 4px 0; line-height: 1.4; color: #0369A1; font-weight: bold; font-size: 11px;">'
            f'Direction: {html.escape(direction)}</p>'
            f'<p style="margin: 4px 0; line-height: 1.4; color: #334155; font-size: 11px;">'
            f'{html.escape(desc)}</p>{fig_html}'
        ),
        "structure_table": json.dumps([]),
    }
    return {
        "clause": "5.1",
        "clause_ref": "5.1",
        "name": "GTP-U Header (Common)",
        "message_name": "GTP-U Header (Common)",
        "table_caption": "Figure 5.1-1: Outline of the GTP-U Header (Clause 5.1)",
        "description": desc,
        "direction": direction,
        "ies": ies,
        "fields": ies,
    }


def build_gpdu_message(
    ext_reg: Dict[str, Dict[str, Any]],
    ie_definitions_dict: Dict[str, Dict[str, Any]],
    fig_511_tbl: Any,
    fig_5213_tbl: Any,
) -> Dict[str, Any]:
    """Construct a hierarchical G-PDU message with GTP-U Header, Extension Headers, TS 38.415 links, and T-PDU."""
    gpdu_appl = GTPU_MESSAGE_INTERFACES["gpdu"]
    ies: List[Dict[str, Any]] = [
        make_ie_row(
            iei="-",
            name="GTP-U Header",
            field_path="GTP-U Header",
            depth=0,
            type_reference="GTP-U Header (Clause 5.1)",
            presence="M",
            fmt="Header",
            length="8 or 12 octets",
            applicability=gpdu_appl,
            field_description=(
                "Mandatory 8-octet GTPv1-U header plus optional 4-octet extension block "
                "(Sequence Number, N-PDU Number, Next Extension Header Type) when E, S, or PN=1."
            ),
        )
    ]
    for f in GTPU_COMMON_HEADER_FIELDS:
        ies.append(
            make_ie_row(
                iei=f["iei"],
                name=f["name"],
                field_path=f"GTP-U Header > {f['name']}",
                depth=1,
                type_reference="GTP-U Header (Clause 5.1)",
                presence=f["presence"],
                fmt=f["format"],
                length=f["length"],
                applicability=gpdu_appl,
                field_description=f["field_description"],
            )
        )

    ext_parent = "GTP-U Extension Headers"
    ies.append(
        make_ie_row(
            iei="E=1",
            name=ext_parent,
            field_path=ext_parent,
            depth=0,
            type_reference="GTP-U Extension Header Format (Clause 5.2.1)",
            presence="O",
            fmt="Ext Header Chain",
            length="4n octets",
            applicability=gpdu_appl,
            field_description=(
                "Chained 4n-octet Extension Headers linked via Next Extension Header Type "
                "(0x00 terminates chain). Bits 7–8 encode intermediate/endpoint comprehension rules."
            ),
        )
    )

    seen_ext_keys = set()
    for key, display_name, clause_ref in GTPU_EXT_HEADER_ORDER:
        seen_ext_keys.add(key)
        ies.extend(
            build_ext_header_rows(
                key=key,
                display_name=display_name,
                clause_ref=clause_ref,
                ext_reg=ext_reg,
                parent_path=ext_parent,
                base_depth=1,
            )
        )
    for key, reg_info in ext_reg.items():
        if key not in seen_ext_keys:
            ies.extend(
                build_ext_header_rows(
                    key=key,
                    display_name=reg_info.get("name", key),
                    clause_ref=f"Clause {reg_info.get('clause', '5.2.2')}",
                    ext_reg=ext_reg,
                    parent_path=ext_parent,
                    base_depth=1,
                )
            )

    ies.append(
        make_ie_row(
            iei="-",
            name="T-PDU Payload",
            field_path="T-PDU Payload",
            depth=0,
            type_reference="Clause 4.4.0",
            presence="C",
            fmt="Payload",
            length="Variable",
            applicability=gpdu_appl,
            field_description=(
                "Encapsulated user data packet (IP/Ethernet/Unstructured). May be omitted when "
                "sending dummy/control G-PDUs carrying RAN Container or PDU Session Container "
                "(e.g. QoS Monitoring)."
            ),
        )
    )

    gpdu_desc, gpdu_dir = GTPU_MESSAGE_DEFAULT_CONTEXT["gpdu"]
    fig_511_html = _convert_table_to_html(fig_511_tbl, is_figure_diagram=True) if fig_511_tbl is not None else ""
    fig_5213_html = _convert_table_to_html(fig_5213_tbl, is_figure_diagram=False) if fig_5213_tbl is not None else ""
    ie_definitions_dict["gpdu"] = {
        "clause": "5.1 (G-PDU)",
        "name": "G-PDU",
        "ie_name": "G-PDU",
        "raw_description": (
            f"{section_header_html('G-PDU', '5.1')}"
            f'<p style="margin: 4px 0; line-height: 1.4; color: #0369A1; font-weight: bold; font-size: 11px;">'
            f'Direction: {html.escape(gpdu_dir)}</p>'
            f'<p style="margin: 4px 0; line-height: 1.4; color: #334155; font-size: 11px;">'
            f'{html.escape(gpdu_desc)}</p>{fig_511_html}{fig_5213_html}'
        ),
        "structure_table": json.dumps([]),
    }
    return {
        "clause": "5.1",
        "clause_ref": "5.1",
        "name": "G-PDU",
        "message_name": "G-PDU",
        "table_caption": "G-PDU Packet Structure and Extension Headers (Clause 5)",
        "description": gpdu_desc,
        "direction": gpdu_dir,
        "ies": ies,
        "fields": ies,
    }
