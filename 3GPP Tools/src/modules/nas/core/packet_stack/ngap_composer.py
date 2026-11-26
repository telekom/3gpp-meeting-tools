"""N2 / NGAP packet stack composer (IP + SCTP Common/DATA/SACK + NGAP APER + optional NAS-PDU)."""

from __future__ import annotations

from typing import Any, Dict, List

from .models import ComposedPacket, FlagOption, PacketConfig, PacketField, PacketLayer
from .ngap_catalog import (
    NGAP_PRESET_MAP,
    NGAP_SCENARIO_PRESETS,
    default_nas_message_for_ngap,
    is_non_ue_ngap_message,
    load_nas_message_fields_from_db,
    load_ngap_message_from_db,
)


def build_sctp_transport_layers(
    transport_mode: str,
    *,
    stream_id: int = 1,
    include_sack: bool = False,
    dst_port: int = 38412,
    ppid: int = 60,
    proto_name: str = "NGAP",
) -> List[PacketLayer]:
    """Construct outer IP (Proto=132) + SCTP Common (12B) + DATA Chunk (16B) [+ SACK (16B)]."""
    mode = (transport_mode or "collapsed").lower()
    if mode == "none":
        return []

    if mode == "collapsed":
        return [
            PacketLayer(
                layer_id="sctp_collapsed",
                title=f"Outer Transport (IP + SCTP Port {dst_port}, PPID={ppid}) — Collapsed",
                spec_number="38.412",
                clause="7",
                role="transport",
                collapsible=True,
                collapsed=True,
                fields=[
                    PacketField(
                        name=f"Outer IP + SCTP ({proto_name})",
                        short_label=(
                            f"IPv4 (20B) / IPv6 (40B) + SCTP Common (12B) + DATA (16B, Stream={stream_id}, PPID={ppid})"
                            + (" + SACK (16B)" if include_sack else "")
                        ),
                        bits=0,
                        presence="M",
                        value_hint=f"48 B (IPv4) / 68 B (IPv6){' +16 B SACK' if include_sack else ''}",
                        spec_number="38.412",
                        clause="7",
                        is_variable=True,
                        variable_display_octets=8,
                        description="Switch Outer dropdown to expand IPv4/IPv6 + SCTP octets.",
                    )
                ],
            )
        ]

    layers: List[PacketLayer] = []
    if mode in ("ipv6_sctp", "ipv6_udp"):
        layers.append(
            PacketLayer(
                layer_id="outer_ipv6_sctp",
                title="Outer IPv6 Header",
                spec_number="RFC 8200",
                clause="3",
                role="transport",
                fields=[
                    PacketField("Version", "Ver=6", 4, "M", "6", "RFC 8200", "3"),
                    PacketField("Traffic Class (MSB)", "TC (hi)", 4, "M", "DSCP", "RFC 8200", "3"),
                    PacketField("Traffic Class (LSB)", "TC (lo)", 4, "M", "ECN", "RFC 8200", "3"),
                    PacketField("Flow Label (MSB)", "Flow (hi)", 4, "M", "0", "RFC 8200", "3"),
                    PacketField("Flow Label (LSB)", "Flow Label (bits 15..0)", 16, "M", "0", "RFC 8200", "3"),
                    PacketField("Payload Length", "Payload Length (16-bit)", 16, "M", "SCTP", "RFC 8200", "3"),
                    PacketField("Next Header", "Next=132 (SCTP)", 8, "M", "132", "RFC 8200", "3"),
                    PacketField("Hop Limit", "Hop Limit", 8, "M", "64", "RFC 8200", "3"),
                    PacketField("Source IPv6 Address", "Source IPv6 Address (128-bit)", 128, "M", "gNB/AMF", "RFC 8200", "3"),
                    PacketField("Destination IPv6 Address", "Destination IPv6 Address (128-bit)", 128, "M", "AMF/gNB", "RFC 8200", "3"),
                ],
            )
        )
    elif mode in ("ipv4_sctp", "ipv4_udp"):
        layers.append(
            PacketLayer(
                layer_id="outer_ipv4_sctp",
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
                    PacketField("Fragment Offset (MSB)", "Frag (hi)", 5, "M", "0", "RFC 791", "3.1"),
                    PacketField("Fragment Offset (LSB)", "Frag Offset (bits 7..0)", 8, "M", "0", "RFC 791", "3.1"),
                    PacketField("Time to Live", "TTL", 8, "M", "64", "RFC 791", "3.1"),
                    PacketField("Protocol", "Proto=132 (SCTP)", 8, "M", "132", "RFC 791", "3.1"),
                    PacketField("Header Checksum", "Header Checksum (16-bit)", 16, "M", "CRC", "RFC 791", "3.1"),
                    PacketField("Source IPv4 Address", "Source IPv4 Address (32-bit)", 32, "M", "gNB/AMF", "RFC 791", "3.1"),
                    PacketField("Destination IPv4 Address", "Destination IPv4 Address (32-bit)", 32, "M", "AMF/gNB", "RFC 791", "3.1"),
                ],
            )
        )

    layers.append(
        PacketLayer(
            layer_id="sctp_common",
            title="SCTP Common Header (12 octets)",
            spec_number="RFC 9260",
            clause="3.1",
            role="sctp_chunk",
            fields=[
                PacketField("Source Port", "Source Port (16-bit)", 16, "M", str(dst_port), "RFC 9260", "3.1"),
                PacketField("Destination Port", f"Dst Port = {dst_port} ({proto_name})", 16, "M", str(dst_port), "38.412", "7"),
                PacketField("Verification Tag", "Verification Tag (32-bit)", 32, "M", "Assoc Tag", "RFC 9260", "3.1"),
                PacketField("SCTP Checksum", "Checksum (CRC32c, 32-bit)", 32, "M", "CRC32c", "RFC 9260", "3.1"),
            ],
        )
    )

    if include_sack:
        layers.append(
            PacketLayer(
                layer_id="sctp_sack",
                title="SCTP SACK Chunk (Bundled, 16 octets)",
                spec_number="RFC 9260",
                clause="3.3.4",
                role="sctp_chunk",
                fields=[
                    PacketField("Chunk Type (SACK)", "Type=3 (SACK)", 8, "O", "3", "RFC 9260", "3.3.4"),
                    PacketField("Chunk Flags", "Flags=0", 8, "O", "0x00", "RFC 9260", "3.3.4"),
                    PacketField("Chunk Length", "Length=16", 16, "O", "16", "RFC 9260", "3.3.4"),
                    PacketField("Cumulative TSN Ack", "Cumulative TSN Ack (32-bit)", 32, "O", "TSN", "RFC 9260", "3.3.4"),
                    PacketField("Advertised Receiver Window Credit", "a_rwnd (32-bit)", 32, "O", "Window", "RFC 9260", "3.3.4"),
                    PacketField("Number of Gap Ack Blocks", "Gap Blks=0", 16, "O", "0", "RFC 9260", "3.3.4"),
                    PacketField("Number of Duplicate TSNs", "Dup TSNs=0", 16, "O", "0", "RFC 9260", "3.3.4"),
                ],
            )
        )

    stream_hint = "0 (Non-UE Common)" if stream_id == 0 else f"≠0 (UE Stream {stream_id})"
    layers.append(
        PacketLayer(
            layer_id="sctp_data",
            title="SCTP DATA Chunk Header (16 octets)",
            spec_number="RFC 9260",
            clause="3.3.1",
            role="sctp_chunk",
            fields=[
                PacketField("Chunk Type (DATA)", "Type=0 (DATA)", 8, "M", "0", "RFC 9260", "3.3.1"),
                PacketField("Reserved", "Res", 4, "M", "0", "RFC 9260", "3.3.1", is_spare=True),
                PacketField("Immediate (I)", "I", 1, "M", "0", "RFC 9260", "3.3.1"),
                PacketField("Unordered (U)", "U", 1, "M", "0", "RFC 9260", "3.3.1"),
                PacketField("Beginning (B)", "B=1", 1, "M", "1", "RFC 9260", "3.3.1"),
                PacketField("Ending (E)", "E=1", 1, "M", "1", "RFC 9260", "3.3.1"),
                PacketField("Chunk Length", "Chunk Length (16 + NGAP)", 16, "M", "16+N", "RFC 9260", "3.3.1"),
                PacketField("Transmission Sequence Number (TSN)", "TSN (32-bit)", 32, "M", "32-bit", "RFC 9260", "3.3.1"),
                PacketField("Stream Identifier", f"Stream = {stream_hint}", 16, "M", stream_hint, "38.412", "7"),
                PacketField("Stream Sequence Number (SSN)", "SSN (16-bit)", 16, "M", "16-bit", "RFC 9260", "3.3.1"),
                PacketField("Payload Protocol Identifier (PPID)", f"PPID = {ppid} ({proto_name})", 32, "M", str(ppid), "38.412", "7"),
            ],
        )
    )
    return layers


def _ie_header_fields(ie: Dict[str, str], msg_clause: str) -> List[PacketField]:
    ie_id = ie.get("id") or "?"
    name = ie["name"]
    crit = ie.get("criticality", "ignore")
    pres = ie.get("presence", "M")
    return [
        PacketField(
            name=name,
            short_label=f"IE ID = {ie_id} ({name})",
            bits=16,
            presence=pres,
            value_hint=ie_id,
            spec_number="38.413",
            clause=msg_clause,
            description=ie.get("field_description") or f"ProtocolIE-ID {ie_id} ({name}, {ie.get('type_reference', name)}).",
        ),
        PacketField(
            name=f"{name} Criticality",
            short_label=f"Crit={crit}",
            bits=2,
            presence=pres,
            value_hint=crit,
            spec_number="38.413",
            clause="9.4",
        ),
        PacketField("APER Pad", "Pad", 6, pres, "0", "X.691", "-", is_spare=True),
        PacketField(
            name=f"{name} Value Length",
            short_label=f"{name} Len (APER 1–2B)",
            bits=8,
            presence=pres,
            value_hint="APER Len",
            spec_number="X.691",
            clause="-",
        ),
    ]


def build_ngap_pdu_layers(
    msg_meta: Dict[str, Any],
    included_ies: List[Dict[str, str]],
    *,
    expand_nas_pdu: bool = False,
    include_nas_security: bool = True,
    nas_fields: List[PacketField] | None = None,
    nas_msg_name: str = "",
    nas_clause: str = "",
) -> List[PacketLayer]:
    """Build NGAP APER PDU envelope and ProtocolIE-Field layers (with optional inline NAS-PDU)."""
    msg_name = msg_meta["message_name"]
    msg_clause = msg_meta["clause"]
    proc_code = msg_meta.get("procedure_code", "0")
    msg_crit = msg_meta.get("msg_criticality", "reject")
    choice = int(msg_meta.get("pdu_choice", 0))
    choice_lbl = {0: "00=InitMsg", 1: "01=SuccOut", 2: "10=UnsuccOut"}.get(choice, "00=InitMsg")

    layers: List[PacketLayer] = [
        PacketLayer(
            layer_id="ngap_envelope",
            title=f"NGAP APER PDU Envelope — {msg_name}",
            spec_number="38.413",
            clause=msg_clause,
            role="ngap_envelope",
            fields=[
                PacketField("NGAP-PDU Ext", "E=0", 1, "M", "0", "38.413", "9.4"),
                PacketField("NGAP-PDU CHOICE", choice_lbl, 2, "M", choice_lbl, "38.413", "9.4"),
                PacketField("APER Alignment Pad", "APER Pad", 5, "M", "0", "X.691", "-", is_spare=True),
                PacketField("Procedure Code", f"procedureCode = {proc_code} ({msg_name})", 8, "M", proc_code, "38.413", "9.3.1.1"),
                PacketField("PDU Criticality", f"Crit={msg_crit}", 2, "M", msg_crit, "38.413", "9.3.1.1"),
                PacketField("APER Alignment Pad", "APER Pad", 6, "M", "0", "X.691", "-", is_spare=True),
                PacketField("Open-Type Length", "Open-Type Length (APER 1–2B)", 8, "M", "Len", "X.691", "-"),
                PacketField("ProtocolIEs Count", f"protocolIEs Sequence Count = {len(included_ies)}", 16, "M", str(len(included_ies)), "38.413", msg_clause),
            ],
        )
    ]

    ie_fields: List[PacketField] = []
    part_idx = 1

    def _flush_ie_layer() -> None:
        nonlocal ie_fields, part_idx
        if not ie_fields:
            return
        suffix = "" if part_idx == 1 else f" (cont. #{part_idx})"
        layers.append(
            PacketLayer(
                layer_id=f"ngap_ies_{part_idx}",
                title=f"NGAP ProtocolIEs — {msg_name}{suffix}",
                spec_number="38.413",
                clause=msg_clause,
                role="ngap_ie",
                fields=list(ie_fields),
            )
        )
        ie_fields = []
        part_idx += 1

    for ie in included_ies:
        ie_fields.extend(_ie_header_fields(ie, msg_clause))
        is_nas = ie["name"].strip().upper() == "NAS-PDU"
        if is_nas and expand_nas_pdu and nas_fields:
            _flush_ie_layer()
            if include_nas_security:
                layers.append(
                    PacketLayer(
                        layer_id="nas_security_hdr",
                        title="5GMM Security Protected NAS Header (7 octets)",
                        spec_number="24.501",
                        clause="9.1",
                        role="nas_security",
                        fields=[
                            PacketField("Extended protocol discriminator", "EPD = 0x7E (5GMM)", 8, "M", "0x7E", "24.501", "9.2"),
                            PacketField("Spare half octet", "Spare", 4, "M", "0", "24.501", "9.5", is_spare=True),
                            PacketField("Security header type", "SecHdr=2 (Int+Ciph)", 4, "M", "2/4", "24.501", "9.3"),
                            PacketField("Message authentication code", "NAS MAC (32-bit)", 32, "M", "MAC", "24.501", "9.8"),
                            PacketField("Sequence number", "NAS SQN (8-bit)", 8, "M", "SQN", "24.501", "9.10"),
                        ],
                    )
                )
            layers.append(
                PacketLayer(
                    layer_id="nas_plain_pdu",
                    title=f"5GS NAS Message — {nas_msg_name}",
                    spec_number="24.501",
                    clause=nas_clause or "8.2",
                    role="nas_message",
                    fields=list(nas_fields),
                )
            )
        else:
            ie_fields.append(
                PacketField(
                    name=f"{ie['name']} Value",
                    short_label=f"{ie['name']} Value ({ie.get('type_reference', ie['name'])}, APER)",
                    bits=0,
                    presence=ie.get("presence", "M"),
                    value_hint="APER",
                    spec_number="38.413",
                    clause=msg_clause,
                    is_variable=True,
                    variable_display_octets=4,
                    description=ie.get("field_description") or f"ASN.1 APER-encoded {ie.get('type_reference', ie['name'])}.",
                )
            )

    _flush_ie_layer()
    return layers


def compose_ngap_packet(config: PacketConfig, db: Any = None) -> ComposedPacket:
    """Compose an end-to-end N2 / NGAP packet stack from `PacketConfig` and `3gpp_protocol_data.db`."""
    msg_name = (config.message_name or "DownlinkNASTransport").strip()
    msg_meta = load_ngap_message_from_db(db, msg_name, config.spec_version_38413)
    msg_name = msg_meta["message_name"]
    all_ies = msg_meta.get("ies") or []

    default_preset = "n2_non_ue" if is_non_ue_ngap_message(msg_name) else "n2_ue_signalling"
    preset = NGAP_PRESET_MAP.get(config.preset_id, NGAP_PRESET_MAP[default_preset])

    has_nas_ie = any(ie["name"].strip().upper() == "NAS-PDU" for ie in all_ies)
    non_ue_default = preset.preset_id == "n2_non_ue" or is_non_ue_ngap_message(msg_name)

    available_flags: List[FlagOption] = [
        FlagOption("sctp_non_ue_stream", "Stream 0 (Non-UE)", "Use SCTP Stream 0 (non-UE signalling) vs Stream ≠ 0 (per-UE association).", "sctp", non_ue_default),
        FlagOption("sctp_sack", "+ SCTP SACK Chunk", "Bundle a 16-octet SCTP SACK acknowledgement chunk before DATA.", "sctp", preset.preset_id == "n2_sctp_bundled_sack"),
    ]
    if has_nas_ie:
        available_flags.extend([
            FlagOption("expand_nas_pdu", "Expand NAS-PDU (24.501)", "Decode encapsulated NAS-PDU (IE ID 38) into TS 24.501 fields.", "sctp", preset.preset_id == "n2_ue_with_nas"),
            FlagOption("nas_security_hdr", "5GMM Security Hdr (7B)", "Include 7-octet ciphered/integrity-protected 5GMM NAS header.", "sctp", True, requires_flag="expand_nas_pdu"),
        ])

    # Build optional IE toggles (cap to 14 pills, prioritizing highlighted IE)
    opt_ies = [ie for ie in all_ies if ie.get("presence", "M").upper() != "M"]
    hl_low = (config.highlight_field or "").strip().lower()
    opt_ies.sort(key=lambda x: (0 if hl_low and hl_low == x["name"].lower() else 1))
    shown_opt = opt_ies[:14]

    opt_key_by_name: Dict[str, str] = {}
    for ie in shown_opt:
        key = f"ie_{ie['id'] or ie['name']}"
        opt_key_by_name[ie["name"]] = key
        is_hl = bool(hl_low and hl_low == ie["name"].lower())
        available_flags.append(
            FlagOption(
                key=key,
                label=f"{ie['name']} ({ie['id'] or ie['presence']})",
                tooltip=f"Include {ie['presence']} IE '{ie['name']}' (ID {ie['id']}, Crit: {ie['criticality']}).",
                group="ngap_ie",
                default_enabled=is_hl,
            )
        )

    effective_flags: Dict[str, bool] = {}
    for opt in available_flags:
        if opt.key in config.flags:
            effective_flags[opt.key] = bool(config.flags[opt.key])
        else:
            effective_flags[opt.key] = opt.default_enabled
    for opt in available_flags:
        if opt.requires_flag and effective_flags.get(opt.key):
            effective_flags[opt.requires_flag] = True

    included_ies = [
        ie for ie in all_ies
        if ie.get("presence", "M").upper() == "M"
        or effective_flags.get(opt_key_by_name.get(ie["name"], ""), False)
    ]

    stream_id = 0 if effective_flags.get("sctp_non_ue_stream") else 1
    include_sack = bool(effective_flags.get("sctp_sack"))
    expand_nas = bool(has_nas_ie and effective_flags.get("expand_nas_pdu"))
    include_sec = bool(effective_flags.get("nas_security_hdr", True))

    nas_name = (config.nas_message_name or default_nas_message_for_ngap(msg_name)).strip()
    nas_fields: List[PacketField] = []
    nas_clause = ""
    nas_ver = ""
    if expand_nas:
        nas_fields, nas_clause, nas_ver = load_nas_message_fields_from_db(db, nas_name)

    sctp_layers = build_sctp_transport_layers(
        config.transport_mode,
        stream_id=stream_id,
        include_sack=include_sack,
    )
    ngap_layers = build_ngap_pdu_layers(
        msg_meta,
        included_ies,
        expand_nas_pdu=expand_nas,
        include_nas_security=include_sec,
        nas_fields=nas_fields,
        nas_msg_name=nas_name,
        nas_clause=nas_clause,
    )
    layers = sctp_layers + ngap_layers

    overhead_3gpp = sum(l.fixed_octets for l in ngap_layers)
    mode = (config.transport_mode or "collapsed").lower()
    if mode in ("ipv4_sctp", "ipv4_udp"):
        overhead_transport = 48 + (16 if include_sack else 0)
    elif mode in ("ipv6_sctp", "ipv6_udp"):
        overhead_transport = 68 + (16 if include_sack else 0)
    elif mode == "sctp_only":
        overhead_transport = 28 + (16 if include_sack else 0)
    else:
        overhead_transport = 0

    spec_refs = ["TS 38.412 (SCTP Port 38412, PPID 60)"]
    v38413 = f" v{msg_meta['version']}" if msg_meta.get("version") else ""
    spec_refs.append(f"TS 38.413{v38413} Cl. {msg_meta['clause']}")
    if expand_nas:
        v24 = f" v{nas_ver}" if nas_ver else ""
        spec_refs.append(f"TS 24.501{v24} ({nas_name})")

    nas_tag = f" + {nas_name}" if expand_nas else ""
    return ComposedPacket(
        title=f"Full N2 Control Plane Stack — {msg_name}{nas_tag}",
        subtitle=preset.description,
        direction=msg_meta["direction"],
        interface=f"N2 (SCTP Stream {'0' if stream_id == 0 else '≠ 0'})",
        spec_refs=spec_refs,
        layers=layers,
        available_flags=available_flags,
        available_presets=NGAP_SCENARIO_PRESETS,
        config=PacketConfig(
            stack_family="ngap",
            target_frame="ngap_pdu",
            preset_id=preset.preset_id,
            transport_mode=config.transport_mode,
            flags=effective_flags,
            include_tpdu=False,
            message_name=msg_name,
            nas_message_name=nas_name,
            spec_version_38413=msg_meta.get("version", ""),
            highlight_field=config.highlight_field,
        ),
        overhead_3gpp_octets=overhead_3gpp,
        overhead_transport_octets=overhead_transport,
        overhead_total_octets=overhead_3gpp + overhead_transport,
        ext_header_summaries=[
            f"NGAP ProcCode={msg_meta['procedure_code']} ({msg_meta['msg_criticality']})",
            f"{len(included_ies)}/{len(all_ies)} Top-Level IEs",
        ],
        warnings=[],
    )
