"""Deterministic packet stack composer & GTP-U 4-octet extension header padding solver."""

from __future__ import annotations

from typing import Any, Dict, List, Tuple

from .models import (
    ComposedPacket,
    PacketConfig,
    PacketField,
    PacketLayer,
)
from .up_catalog import (
    UP_PRESET_MAP,
    UP_SCENARIO_PRESETS,
    build_dl_pdu_session_fields,
    build_dl_pdu_set_fields,
    build_nr_ran_container_fields,
    build_pdcp_sn_fields,
    build_ul_pdu_session_fields,
    get_flags_for_target,
    load_up_release_capabilities,
)


def build_transport_layers(transport_mode: str) -> List[PacketLayer]:
    """Construct outer IP + UDP (Port 2152) transport layers per TS 29.281 Cl. 4 / TS 38.414."""
    mode = (transport_mode or "collapsed").lower()
    if mode == "none":
        return []

    if mode == "collapsed":
        return [
            PacketLayer(
                layer_id="transport_collapsed",
                title="Outer Transport (IP + UDP Dst Port 2152) — Collapsed",
                spec_number="38.414",
                clause="5",
                role="transport",
                collapsible=True,
                collapsed=True,
                fields=[
                    PacketField(
                        name="Outer IPv4/IPv6 + UDP (Port 2152)",
                        short_label="IPv4 (20B) or IPv6 (40B) + UDP (8B, Dst Port 2152)",
                        bits=0,
                        presence="M",
                        value_hint="28 B (IPv4) / 48 B (IPv6)",
                        spec_number="38.414",
                        clause="5",
                        is_variable=True,
                        variable_display_octets=8,
                        description="Switch Transport dropdown to IPv4 or IPv6 to expand every transport octet.",
                    )
                ],
            )
        ]

    layers: List[PacketLayer] = []
    if mode == "ipv6_udp":
        layers.append(
            PacketLayer(
                layer_id="outer_ipv6",
                title="Outer IPv6 Header",
                spec_number="RFC 8200",
                clause="3",
                role="transport",
                fields=[
                    PacketField("Version", "Ver=6", 4, "M", "6", "RFC 8200", "3"),
                    PacketField("Traffic Class (DSCP/ECN)", "Traffic Class", 8, "M", "DSCP", "RFC 8200", "3"),
                    PacketField("Flow Label", "Flow Label (20-bit)", 20, "M", "0", "RFC 8200", "3"),
                    PacketField("Payload Length", "Payload Length (16-bit)", 16, "M", "UDP+GTP", "RFC 8200", "3"),
                    PacketField("Next Header", "Next=17 (UDP)", 8, "M", "17", "RFC 8200", "3"),
                    PacketField("Hop Limit", "Hop Limit", 8, "M", "64", "RFC 8200", "3"),
                    PacketField("Source IPv6 Address", "Source IPv6 Address (128-bit)", 128, "M", "UPF/gNB", "RFC 8200", "3"),
                    PacketField("Destination IPv6 Address", "Destination IPv6 Address (128-bit)", 128, "M", "gNB/UPF", "RFC 8200", "3"),
                ],
            )
        )
    else:  # ipv4_udp
        layers.append(
            PacketLayer(
                layer_id="outer_ipv4",
                title="Outer IPv4 Header",
                spec_number="RFC 791",
                clause="3.1",
                role="transport",
                fields=[
                    PacketField("Version", "Ver=4", 4, "M", "4", "RFC 791", "3.1"),
                    PacketField("IHL", "IHL=5", 4, "M", "5", "RFC 791", "3.1"),
                    PacketField("DSCP / ECN", "DSCP/ECN", 8, "M", "QoS", "RFC 791", "3.1"),
                    PacketField("Total Length", "Total Length (16-bit)", 16, "M", "Octets", "RFC 791", "3.1"),
                    PacketField("Identification", "Identification (16-bit)", 16, "M", "-", "RFC 791", "3.1"),
                    PacketField("Flags", "Flags", 3, "M", "-", "RFC 791", "3.1"),
                    PacketField("Fragment Offset", "Fragment Offset (13-bit)", 13, "M", "0", "RFC 791", "3.1"),
                    PacketField("Time to Live", "TTL", 8, "M", "64", "RFC 791", "3.1"),
                    PacketField("Protocol", "Proto=17 (UDP)", 8, "M", "17", "RFC 791", "3.1"),
                    PacketField("Header Checksum", "Header Checksum (16-bit)", 16, "M", "CRC", "RFC 791", "3.1"),
                    PacketField("Source IPv4 Address", "Source IPv4 Address (32-bit)", 32, "M", "UPF/gNB", "RFC 791", "3.1"),
                    PacketField("Destination IPv4 Address", "Destination IPv4 Address (32-bit)", 32, "M", "gNB/UPF", "RFC 791", "3.1"),
                ],
            )
        )

    layers.append(
        PacketLayer(
            layer_id="outer_udp",
            title="Outer UDP Header",
            spec_number="29.281",
            clause="4.4.2",
            role="transport",
            fields=[
                PacketField("Source Port", "Source Port (16-bit)", 16, "M", "Hash/2152", "29.281", "4.4.2"),
                PacketField("Destination Port", "Dst Port = 2152 (GTP-U)", 16, "M", "2152", "29.281", "4.4.2"),
                PacketField("UDP Length", "UDP Length (16-bit)", 16, "M", "8+GTP", "RFC 768", "-"),
                PacketField("UDP Checksum", "UDP Checksum (16-bit)", 16, "M", "0/CRC", "29.281", "4.4.2"),
            ],
        )
    )
    return layers


def build_gtpu_header_layers(
    *,
    msg_type_val: int,
    msg_type_label: str,
    e_flag: bool,
    s_flag: bool,
    pn_flag: bool = False,
    first_next_ext_hex: str = "0x00",
    first_next_ext_name: str = "No more extension headers",
) -> List[PacketLayer]:
    """Build mandatory 8-octet GTPv1-U Header plus optional 4-octet block (Octets 9–12)."""
    base_fields = [
        PacketField("Version", "Ver=1", 3, "M", "001", "29.281", "5.1", description="GTPv1-U version ('001')."),
        PacketField("Protocol Type (PT)", "PT=1", 1, "M", "1", "29.281", "5.1", description="'1' for GTP (TS 29.281)."),
        PacketField("Spare (*)", "*", 1, "M", "0", "29.281", "5.1", is_spare=True, description="Spare bit, always 0."),
        PacketField("Extension Header flag (E)", "E", 1, "M", "1" if e_flag else "0", "29.281", "5.1", description="1 = Next Extension Header Type present in Octet 12."),
        PacketField("Sequence number flag (S)", "S", 1, "M", "1" if s_flag else "0", "29.281", "5.1", flag_key="gtpu_seq", description="1 = Sequence Number in Octets 9–10 is meaningful."),
        PacketField("N-PDU Number flag (PN)", "PN", 1, "M", "1" if pn_flag else "0", "29.281", "5.1", description="1 = N-PDU Number in Octet 11 is meaningful."),
        PacketField("Message Type", msg_type_label, 8, "M", f"0x{msg_type_val:02X} ({msg_type_val})", "29.281", "6.1", description=f"GTP-U Message Type ({msg_type_val})."),
        PacketField("Length", "Length (octets after Octet 8)", 16, "M", "Payload len", "29.281", "5.1", description="Payload length in octets following the mandatory 8-octet header."),
        PacketField("Tunnel Endpoint Identifier (TEID)", "TEID (32-bit)", 32, "M", "32-bit", "29.281", "5.1", description="Unambiguously identifies the GTP-U tunnel endpoint."),
    ]
    layers = [
        PacketLayer(
            layer_id="gtpu_base",
            title="GTPv1-U Mandatory Header (Octets 1–8)",
            spec_number="29.281",
            clause="5.1",
            role="gtpu_header",
            fields=base_fields,
        )
    ]

    if e_flag or s_flag or pn_flag:
        opt_fields = [
            PacketField("Sequence Number", "Sequence Number (16-bit)" if s_flag else "Sequence Number (ignored, S=0)", 16, "C", "16-bit" if s_flag else "0x0000", "29.281", "5.1", condition="E=1 or S=1 or PN=1", flag_key="gtpu_seq"),
            PacketField("N-PDU Number", "N-PDU (ignored)", 8, "C", "0x00", "29.281", "5.1", condition="E=1 or S=1 or PN=1"),
            PacketField("Next Extension Header Type", f"Next Ext = {first_next_ext_hex}", 8, "C", f"{first_next_ext_hex} ({first_next_ext_name})", "29.281", "5.2.1", condition="E=1 or S=1 or PN=1"),
        ]
        layers.append(
            PacketLayer(
                layer_id="gtpu_opt",
                title="GTPv1-U Optional Header (Octets 9–12)",
                spec_number="29.281",
                clause="5.1",
                role="gtpu_header",
                fields=opt_fields,
            )
        )
    return layers


def compose_gtpu_extension_header(
    *,
    layer_id: str,
    ext_code_hex: str,
    ext_name: str,
    ext_clause: str,
    inner_fields: List[PacketField],
    next_ext_hex: str,
    next_ext_name: str,
    inner_role: str = "up_frame",
    inner_spec: str = "38.415",
    inner_clause: str = "5.5.2.1",
    padding_clause: str = "6.5.3.13",
) -> Tuple[List[PacketLayer], str, int]:
    """
    Wrap an extension-header content payload in TS 29.281 Clause 5.2.1 framing:
      [Length (1B)] + [Content] + [Padding (0–3B)] + [Next Extension Header Type (1B)]
    Returns (layers, summary_text, fixed_octets_or_0_if_variable).
    """
    has_var = any(f.is_variable for f in inner_fields)
    fixed_inner_bits = sum(f.bits for f in inner_fields if not f.is_variable)
    if fixed_inner_bits % 8 != 0:
        raise ValueError(
            f"Extension header '{ext_name}' inner fixed fields must be octet-aligned (got {fixed_inner_bits} bits)."
        )
    fixed_inner_octets = fixed_inner_bits // 8

    if has_var:
        len_hint = "n (4n octets)"
        padding_field = PacketField(
            name="Padding (4-octet align)",
            short_label="Padding (0–3 octets for 4n alignment)",
            bits=0,
            presence="O",
            value_hint="0–3 B",
            spec_number="29.281",
            clause="5.2.1",
            is_padding=True,
            is_variable=True,
            variable_display_octets=2,
            description="0–3 octets so total extension header length is an integer multiple of 4 octets.",
        )
        total_fixed = 0
        summary = f"{ext_name} ({ext_code_hex}): 4n octets (variable) → Next {next_ext_hex}"
    else:
        raw_octets = 1 + fixed_inner_octets + 1
        pad_octets = (4 - (raw_octets % 4)) % 4
        total_fixed = raw_octets + pad_octets
        n_units = total_fixed // 4
        len_hint = f"n={n_units} ({total_fixed} B)"
        padding_field = (
            PacketField(
                name="Padding",
                short_label=f"Padding ({pad_octets} octet{'s' if pad_octets != 1 else ''})",
                bits=pad_octets * 8,
                presence="O",
                value_hint=f"{pad_octets} B (0x00..)",
                spec_number=inner_spec,
                clause=padding_clause,
                is_padding=True,
                description=f"Auto-computed {pad_octets}-octet padding so GTP-U Extension Header length ({total_fixed} octets = {n_units}×4) satisfies TS 29.281 Cl. 5.2.1.",
            )
            if pad_octets > 0
            else None
        )
        summary = f"{ext_name} ({ext_code_hex}): {total_fixed} B (n={n_units}, pad={pad_octets} B) → Next {next_ext_hex}"

    hdr_open = PacketLayer(
        layer_id=f"{layer_id}_open",
        title=f"GTP-U Extension Header: {ext_name} ({ext_code_hex}) — Framing",
        spec_number="29.281",
        clause=ext_clause,
        role="gtpu_ext_wrapper",
        fields=[
            PacketField(
                name="Extension Header Length",
                short_label=f"Ext Hdr Length = {len_hint}",
                bits=8,
                presence="M",
                value_hint=len_hint,
                spec_number="29.281",
                clause="5.2.1",
                description="Length of this extension header in 4-octet units (1 octet).",
            )
        ],
    )

    content_fields = list(inner_fields)
    if padding_field is not None:
        content_fields.append(padding_field)

    content_layer = PacketLayer(
        layer_id=f"{layer_id}_body",
        title=f"{ext_name} Content",
        spec_number=inner_spec,
        clause=inner_clause,
        role=inner_role,
        fields=content_fields,
    )

    hdr_close = PacketLayer(
        layer_id=f"{layer_id}_close",
        title=f"{ext_name} — Chain Terminator / Pointer",
        spec_number="29.281",
        clause="5.2.1",
        role="gtpu_ext_wrapper",
        fields=[
            PacketField(
                name="Next Extension Header Type",
                short_label=f"Next Ext = {next_ext_hex} ({next_ext_name})",
                bits=8,
                presence="M",
                value_hint=next_ext_hex,
                spec_number="29.281",
                clause="5.2.1",
                description="0x00 if no further extension header follows, or codepoint of next extension header.",
            )
        ],
    )
    return [hdr_open, content_layer, hdr_close], summary, total_fixed


def compose_up_packet(config: PacketConfig, db: Any = None) -> ComposedPacket:
    """Compose a complete 5G User Plane packet stack from `PacketConfig` and optional protocol DB."""
    preset = UP_PRESET_MAP.get(config.preset_id, UP_SCENARIO_PRESETS[0])
    caps = load_up_release_capabilities(db, config.spec_version_38415, config.spec_version_29281)
    available_flags = get_flags_for_target(config.target_frame, preset.preset_id, caps)

    # Normalize & enforce flag dependencies
    effective_flags: Dict[str, bool] = {}
    for opt in available_flags:
        if opt.key in config.flags:
            effective_flags[opt.key] = bool(config.flags[opt.key])
        else:
            effective_flags[opt.key] = opt.default_enabled
    for opt in available_flags:
        if opt.requires_flag and effective_flags.get(opt.key):
            effective_flags[opt.requires_flag] = True

    is_end_marker = config.target_frame == "end_marker" or preset.preset_id == "end_marker"
    msg_type_val = 254 if is_end_marker else 255
    msg_type_label = "MsgType=254 (End Marker)" if is_end_marker else "MsgType=255 (G-PDU)"

    # Build ordered extension header chain per TS 29.281
    ext_specs: List[Dict[str, Any]] = []
    if not is_end_marker and effective_flags.get("ext_pdcp_sn"):
        ext_specs.append({
            "layer_id": "ext_pdcp_sn",
            "code": "0x03",
            "name": "Long PDCP PDU Number",
            "ext_clause": "5.2.2.2A",
            "inner_fields": build_pdcp_sn_fields(),
            "inner_role": "gtpu_ext_wrapper",
            "inner_spec": "29.281",
            "inner_clause": "5.2.2.2A",
        })
    if not is_end_marker and effective_flags.get("ext_nr_ran"):
        ext_specs.append({
            "layer_id": "ext_nr_ran",
            "code": "0x84",
            "name": "NR RAN Container",
            "ext_clause": "5.2.2.6",
            "inner_fields": build_nr_ran_container_fields(),
            "inner_role": "ran_container",
            "inner_spec": "38.425",
            "inner_clause": "5.5",
        })

    if config.target_frame == "ul_pdu_session" and effective_flags.get("ext_pdu_session", True):
        ext_specs.append({
            "layer_id": "ext_pdu_session_ul",
            "code": "0x85",
            "name": "PDU Session Container (UL PDU SESSION INFO)",
            "ext_clause": "5.2.2.7",
            "inner_fields": build_ul_pdu_session_fields(effective_flags, caps),
            "inner_role": "up_frame",
            "inner_spec": "38.415",
            "inner_clause": "5.5.2.2",
        })
    elif effective_flags.get("ext_pdu_session", config.target_frame in ("dl_pdu_session", "end_marker")):
        ext_specs.append({
            "layer_id": "ext_pdu_session_dl",
            "code": "0x85",
            "name": "PDU Session Container (DL PDU SESSION INFO)",
            "ext_clause": "5.2.2.7",
            "inner_fields": build_dl_pdu_session_fields(effective_flags, caps),
            "inner_role": "up_frame",
            "inner_spec": "38.415",
            "inner_clause": "5.5.2.1",
        })

    if config.target_frame == "dl_pdu_set" or effective_flags.get("ext_pdu_set"):
        ext_specs.append({
            "layer_id": "ext_pdu_set",
            "code": "0x04",
            "name": "PDU Set Information Container (DL PDU SET INFO)",
            "ext_clause": "5.2.2.8",
            "inner_fields": build_dl_pdu_set_fields(effective_flags, caps),
            "inner_role": "up_frame",
            "inner_spec": "38.415",
            "inner_clause": "6.5.2.1",
        })

    layers: List[PacketLayer] = []
    transport_layers = build_transport_layers(config.transport_mode)
    layers.extend(transport_layers)

    e_flag = bool(ext_specs)
    s_flag = bool(effective_flags.get("gtpu_seq"))
    first_ext_hex = ext_specs[0]["code"] if ext_specs else "0x00"
    first_ext_name = ext_specs[0]["name"] if ext_specs else "None"

    gtpu_layers = build_gtpu_header_layers(
        msg_type_val=msg_type_val,
        msg_type_label=msg_type_label,
        e_flag=e_flag,
        s_flag=s_flag,
        pn_flag=False,
        first_next_ext_hex=first_ext_hex,
        first_next_ext_name=first_ext_name,
    )
    layers.extend(gtpu_layers)

    ext_summaries: List[str] = []
    ext_fixed_total = 0
    for idx, spec in enumerate(ext_specs):
        nxt = ext_specs[idx + 1] if idx + 1 < len(ext_specs) else None
        next_hex = nxt["code"] if nxt else "0x00"
        next_name = nxt["name"] if nxt else "No more extension headers"
        ext_layers, summary, fixed_oct = compose_gtpu_extension_header(
            layer_id=spec["layer_id"],
            ext_code_hex=spec["code"],
            ext_name=spec["name"],
            ext_clause=spec["ext_clause"],
            inner_fields=spec["inner_fields"],
            next_ext_hex=next_hex,
            next_ext_name=next_name,
            inner_role=spec["inner_role"],
            inner_spec=spec["inner_spec"],
            inner_clause=spec["inner_clause"],
        )
        layers.extend(ext_layers)
        ext_summaries.append(summary)
        ext_fixed_total += fixed_oct

    if config.include_tpdu and not is_end_marker:
        layers.append(
            PacketLayer(
                layer_id="tpdu_payload",
                title="Encapsulated User Data (T-PDU)",
                spec_number="29.281",
                clause="4.4.0",
                role="payload",
                fields=[
                    PacketField(
                        name="T-PDU Payload",
                        short_label=config.tpdu_label,
                        bits=0,
                        presence="C",
                        value_hint="Inner UE IP / Ethernet packet",
                        spec_number="29.281",
                        clause="4.4.0",
                        is_variable=True,
                        variable_display_octets=16,
                        description="Inner user-plane PDU (IPv4, IPv6, Ethernet, or Unstructured).",
                    )
                ],
            )
        )

    gtpu_octets = sum(l.fixed_octets for l in gtpu_layers)
    overhead_3gpp = gtpu_octets + ext_fixed_total
    if config.transport_mode == "ipv4_udp":
        overhead_transport = 28
    elif config.transport_mode == "ipv6_udp":
        overhead_transport = 48
    else:
        overhead_transport = 0

    frame_title_map = {
        "dl_pdu_session": "DL PDU SESSION INFORMATION (PDU Type 0)",
        "ul_pdu_session": "UL PDU SESSION INFORMATION (PDU Type 1)",
        "dl_pdu_set": "DL PDU SET INFORMATION (PDU Type 0)",
        "end_marker": "GTP-U End Marker (MsgType 254)",
    }
    direction_map = {
        "dl_pdu_session": "UPF → NG-RAN node",
        "ul_pdu_session": "NG-RAN node → UPF",
        "dl_pdu_set": "UPF / gNB-CU-UP → NG-RAN / gNB-DU",
        "end_marker": "Source GTP-U Peer → Target GTP-U Peer",
    }

    spec_refs = ["TS 29.281 Cl. 5.1 / 5.2.1"]
    if caps.get("version_29281"):
        spec_refs[0] = f"TS 29.281 v{caps['version_29281']}"
    if any(l.spec_number == "38.415" for l in layers):
        v38 = f" v{caps['version_38415']}" if caps.get("version_38415") else ""
        spec_refs.append(f"TS 38.415{v38}")
    if any(l.spec_number == "38.425" for l in layers):
        spec_refs.append("TS 38.425 (NR UP)")

    return ComposedPacket(
        title=f"Full 5G UP Packet Stack — {frame_title_map.get(config.target_frame, preset.label)}",
        subtitle=preset.description,
        direction=direction_map.get(config.target_frame, "GTP-U Peer ↔ GTP-U Peer"),
        interface=preset.interface,
        spec_refs=spec_refs,
        layers=layers,
        available_flags=available_flags,
        available_presets=UP_SCENARIO_PRESETS,
        config=PacketConfig(
            target_frame=config.target_frame,
            preset_id=preset.preset_id,
            transport_mode=config.transport_mode,
            flags=effective_flags,
            include_tpdu=config.include_tpdu and not is_end_marker,
            tpdu_label=config.tpdu_label,
            spec_version_38415=caps.get("version_38415", ""),
            spec_version_29281=caps.get("version_29281", ""),
            highlight_field=config.highlight_field,
        ),
        overhead_3gpp_octets=overhead_3gpp,
        overhead_transport_octets=overhead_transport,
        overhead_total_octets=overhead_3gpp + overhead_transport,
        ext_header_summaries=ext_summaries,
        warnings=[],
    )
