"""5G NR RAN Uu (TS 38.300/38.331/37.324/38.323/38.322/38.321) CP & UP packet stack composer."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from .models import ComposedPacket, FlagOption, PacketConfig, PacketField, PacketLayer
from .nas_composer import build_nas_security_layer
from .ngap_catalog import load_nas_message_fields_from_db
from .ran_uu_catalog import (
    RAN_UU_PRESET_MAP,
    RAN_UU_SCENARIO_PRESETS,
    infer_rrc_channel_profile,
    load_rrc_message_from_db,
)

_DEPTH_SETS: Dict[str, set[str]] = {
    "uu_depth_l3": {"l3", "sdap"},
    "uu_depth_pdcp": {"l3", "sdap", "pdcp"},
    "uu_depth_rlc": {"l3", "sdap", "pdcp", "rlc"},
    "uu_depth_mac": {"l3", "sdap", "pdcp", "rlc", "mac"},
}


def _depth_includes(depth_mode: str, layer: str) -> bool:
    return layer in _DEPTH_SETS.get((depth_mode or "uu_depth_mac").lower(), _DEPTH_SETS["uu_depth_mac"])


def build_mac_subheader_layer(lcid: int, channel_label: str, seg_tag: str = "") -> PacketLayer:
    return PacketLayer(
        layer_id=f"uu_mac_sub{'_' + seg_tag if seg_tag else ''}",
        title=f"MAC Subheader — {channel_label}" + (f" ({seg_tag})" if seg_tag else ""),
        spec_number="38.321",
        clause="6.2.1",
        role="uu_mac",
        fields=[
            PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
            PacketField("F", "F", 1, "M", "0/1", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
            PacketField("LCID", f"LCID = {lcid} ({channel_label})", 6, "M", str(lcid), "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
            PacketField("L", "L (MAC SDU Len, 8/16-bit)", 8, "M", "Len", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
        ],
    )


def build_mac_ce_layer(is_uplink: bool) -> PacketLayer:
    if is_uplink:
        fields = [
            PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
            PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
            PacketField("LCID (MAC CE)", "LCID=61 (Short BSR)", 6, "O", "61", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
            PacketField("LCG ID", "LCG ID", 3, "O", "0..7", "38.321", "6.1.3.1", def_name="Buffer Status Report MAC CEs"),
            PacketField("Buffer Size", "Buffer Size (BSR Index)", 5, "O", "0..31", "38.321", "6.1.3.1", def_name="Buffer Status Report MAC CEs"),
        ]
        title = "MAC Control Element — Short BSR (UL: After SDU, Fig. 6.1.2-5)"
    else:
        fields = [
            PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
            PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
            PacketField("LCID (MAC CE)", "LCID=61 (Timing Adv)", 6, "O", "61", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
            PacketField("TAG ID", "TAG", 2, "O", "0..3", "38.321", "6.1.3.4", def_name="Timing Advance Command MAC CE"),
            PacketField("Timing Advance Command", "Timing Advance Cmd", 6, "O", "0..63", "38.321", "6.1.3.4", def_name="Timing Advance Command MAC CE"),
        ]
        title = "MAC Control Element — Timing Advance Cmd (DL: Front of TB, Fig. 6.1.2-4)"
    return PacketLayer(layer_id="uu_mac_ce", title=title, spec_number="38.321", clause="6.1.3", role="uu_mac", fields=fields)


def build_rlc_header_layer(rlc_mode: str, *, si_bits: str = "00", include_so: bool = False, seg_tag: str = "") -> PacketLayer:
    mode = (rlc_mode or "AM").upper()
    suffix = f" ({seg_tag})" if seg_tag else ""
    if mode == "TM":
        return PacketLayer(
            layer_id="uu_rlc_tm",
            title=f"RLC TM — Transparent Mode{suffix}",
            spec_number="38.322",
            clause="6.2.2.2",
            role="uu_rlc",
            fields=[
                PacketField("TMD PDU", "RLC TM — Transparent (0-byte header, no SN/SI)", 0, "M", "0 B", "38.322", "6.2.2.2", is_variable=False, def_name="TMD PDU")
            ],
        )

    if mode == "UM":
        if si_bits == "00":
            fields = [
                PacketField("SI", "SI=00 (Full)", 2, "M", "00", "38.322", "6.2.3.4", def_name="Segmentation Info (SI) field"),
                PacketField("R", "R=000000", 6, "M", "0", "38.322", "6.2.3.8", is_spare=True, def_name="Reserved (R) field"),
            ]
        else:
            fields = [
                PacketField("SI", f"SI={si_bits}", 2, "M", si_bits, "38.322", "6.2.3.4", def_name="Segmentation Info (SI) field"),
                PacketField("R", "R", 2, "M", "0", "38.322", "6.2.3.8", is_spare=True, def_name="Reserved (R) field"),
                PacketField("SN", "SN (hi 4b)", 4, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
                PacketField("SN (lo)", "SN (bits 7..0)", 8, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
            ]
            if include_so:
                fields.append(PacketField("SO", "SO = K (Segment Offset, 16-bit)", 16, "C", "K", "38.322", "6.2.3.5", condition="SI=10 or 11", def_name="Segment Offset (SO) field"))
        return PacketLayer(layer_id=f"uu_rlc_um{'_' + seg_tag if seg_tag else ''}", title=f"RLC UMD Header{suffix}", spec_number="38.322", clause="6.2.2.3", role="uu_rlc", fields=fields)

    fields = [
        PacketField("D/C", "D/C=1", 1, "M", "1", "38.322", "6.2.3.6", def_name="Data/Control (D/C) field"),
        PacketField("P", "P", 1, "M", "0/1", "38.322", "6.2.3.7", def_name="Polling bit (P) field"),
        PacketField("SI", f"SI={si_bits}", 2, "M", si_bits, "38.322", "6.2.3.4", def_name="Segmentation Info (SI) field"),
        PacketField("SN", "SN (hi 4b)", 4, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
        PacketField("SN (lo)", "SN (bits 7..0)", 8, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
    ]
    if include_so:
        fields.append(PacketField("SO", "SO = K (Segment Offset, 16-bit)", 16, "C", "K", "38.322", "6.2.3.5", condition="SI=10 or 11", def_name="Segment Offset (SO) field"))
    return PacketLayer(layer_id=f"uu_rlc_am{'_' + seg_tag if seg_tag else ''}", title=f"RLC AMD Header (12-bit SN){suffix}", spec_number="38.322", clause="6.2.2.4", role="uu_rlc", fields=fields)


def build_pdcp_header_layer(*, is_srb: bool, sn_18: bool = False) -> PacketLayer:
    if is_srb:
        return PacketLayer(
            layer_id="uu_pdcp_srb",
            title="PDCP Data PDU Header — SRB (12-bit SN)",
            spec_number="38.323",
            clause="6.2.2.1",
            role="uu_pdcp",
            fields=[
                PacketField("R", "R=0000", 4, "M", "0", "38.323", "6.3.6", is_spare=True),
                PacketField("PDCP SN", "PDCP SN (hi 4b)", 4, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
                PacketField("PDCP SN (lo)", "PDCP SN (bits 7..0)", 8, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
            ],
        )
    if sn_18:
        return PacketLayer(
            layer_id="uu_pdcp_drb18",
            title="PDCP Data PDU Header — DRB (18-bit SN)",
            spec_number="38.323",
            clause="6.2.2.3",
            role="uu_pdcp",
            fields=[
                PacketField("D/C", "D/C=1", 1, "M", "1", "38.323", "6.3.7", def_name="D/C"),
                PacketField("R", "R=00000", 5, "M", "0", "38.323", "6.3.6", is_spare=True),
                PacketField("PDCP SN", "PDCP SN (17..16)", 2, "M", "18-bit", "38.323", "6.3.2", def_name="PDCP SN"),
                PacketField("PDCP SN (lo)", "PDCP SN (bits 15..0)", 16, "M", "18-bit", "38.323", "6.3.2", def_name="PDCP SN"),
            ],
        )
    return PacketLayer(
        layer_id="uu_pdcp_drb12",
        title="PDCP Data PDU Header — DRB (12-bit SN)",
        spec_number="38.323",
        clause="6.2.2.2",
        role="uu_pdcp",
        fields=[
            PacketField("D/C", "D/C=1", 1, "M", "1", "38.323", "6.3.7", def_name="D/C"),
            PacketField("R", "R=000", 3, "M", "0", "38.323", "6.3.6", is_spare=True),
            PacketField("PDCP SN", "PDCP SN (hi 4b)", 4, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
            PacketField("PDCP SN (lo)", "PDCP SN (bits 7..0)", 8, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
        ],
    )


def build_pdcp_maci_layer(is_srb: bool) -> PacketLayer:
    clause = "6.2.2.1" if is_srb else "6.2.2.2"
    return PacketLayer(
        layer_id="uu_pdcp_maci",
        title="PDCP Integrity Protection Trailer",
        spec_number="38.323",
        clause=clause,
        role="uu_pdcp",
        fields=[PacketField("MAC-I", "MAC-I (32-bit Integrity Check)", 32, "M" if is_srb else "C", "32-bit", "38.323", "6.3.4", def_name="MAC-I")],
    )


def build_sdap_header_layer(is_uplink: bool) -> PacketLayer:
    if is_uplink:
        return PacketLayer(
            layer_id="uu_sdap_ul",
            title="SDAP UL Header (TS 37.324 Cl. 6.2.2.3)",
            spec_number="37.324",
            clause="6.2.2.3",
            role="uu_sdap",
            fields=[
                PacketField("D/C", "D/C=1", 1, "M", "1", "37.324", "6.3.3", def_name="D/C"),
                PacketField("R", "R=0", 1, "M", "0", "37.324", "6.3.5", is_spare=True),
                PacketField("QFI", "QFI (6-bit QoS Flow ID)", 6, "M", "0..63", "37.324", "6.3.4", def_name="QFI"),
            ],
        )
    return PacketLayer(
        layer_id="uu_sdap_dl",
        title="SDAP DL Header (TS 37.324 Cl. 6.2.2.2)",
        spec_number="37.324",
        clause="6.2.2.2",
        role="uu_sdap",
        fields=[
            PacketField("RDI", "RDI", 1, "M", "0/1", "37.324", "6.3.7", def_name="RDI"),
            PacketField("RQI", "RQI", 1, "M", "0/1", "37.324", "6.3.6", def_name="RQI"),
            PacketField("QFI", "QFI (6-bit QoS Flow ID)", 6, "M", "0..63", "37.324", "6.3.4", def_name="QFI"),
        ],
    )


def _assemble_with_segmentation(
    *,
    is_uplink: bool,
    depth_mode: str,
    lcid: Optional[int],
    channel_label: str,
    rlc_mode: str,
    segmented: bool,
    mac_ce: bool,
    mac_padding: bool,
    pdcp_layer: Optional[PacketLayer],
    upper_layers: List[PacketLayer],
    payload_layer: Optional[PacketLayer],
    maci_layer: Optional[PacketLayer],
    sdu_desc: str,
) -> List[PacketLayer]:
    """Arrange MAC CE / MAC / RLC / PDCP / Upper / MAC-I across 1 TB or 2 segmented RLC TBs."""
    has_mac = _depth_includes(depth_mode, "mac") and lcid is not None
    has_rlc = _depth_includes(depth_mode, "rlc")
    can_seg = segmented and has_rlc and rlc_mode in ("AM", "UM")

    layers: List[PacketLayer] = []
    if has_mac and mac_ce and not is_uplink:
        layers.append(build_mac_ce_layer(is_uplink=False))

    if not can_seg:
        if has_mac and lcid is not None:
            layers.append(build_mac_subheader_layer(lcid, channel_label))
        if has_rlc:
            layers.append(build_rlc_header_layer(rlc_mode, si_bits="00", include_so=False))
        if pdcp_layer is not None:
            layers.append(pdcp_layer)
        layers.extend(upper_layers)
        if payload_layer is not None:
            layers.append(payload_layer)
        if maci_layer is not None:
            layers.append(maci_layer)
    else:
        # TB #1: SI=01 (First Segment, no SO), carries PDCP/SDAP/RRC headers + first slice
        if has_mac and lcid is not None:
            layers.append(build_mac_subheader_layer(lcid, channel_label, seg_tag="TB #1"))
        layers.append(build_rlc_header_layer(rlc_mode, si_bits="01", include_so=False, seg_tag="TB #1: First"))
        if pdcp_layer is not None:
            layers.append(pdcp_layer)
        layers.extend(upper_layers)
        layers.append(
            PacketLayer(
                layer_id="uu_seg_slice1",
                title=f"{sdu_desc} — Segment #1 [Bytes 0..K-1]",
                spec_number="38.322",
                clause="6.2.2",
                role="payload",
                fields=[PacketField(f"{sdu_desc} Slice #1", "SDU Payload Slice #1 (Bytes 0..K-1)", 0, "M", "Bytes 0..K-1", "38.322", "6.2.2", is_variable=True, variable_display_octets=6)],
            )
        )
        # Segmentation Divider (no data cells -> colspan=11 banner row)
        layers.append(
            PacketLayer(
                layer_id="uu_seg_divider",
                title="✂️ RLC Segmentation Boundary — SDU split at byte K (SO = K) across 2 MAC TBs (PDCP/SDAP/RRC headers not repeated)",
                spec_number="38.322",
                clause="6.2.2",
                role="segmentation_divider",
                fields=[],
            )
        )
        # TB #2: SI=10 (Last Segment + 16-bit SO=K), continuation + trailing MAC-I
        if has_mac and lcid is not None:
            layers.append(build_mac_subheader_layer(lcid, channel_label, seg_tag="TB #2"))
        layers.append(build_rlc_header_layer(rlc_mode, si_bits="10", include_so=True, seg_tag="TB #2: Last"))
        layers.append(
            PacketLayer(
                layer_id="uu_seg_slice2",
                title=f"{sdu_desc} — Segment #2 [Bytes K..End]",
                spec_number="38.322",
                clause="6.2.2",
                role="payload",
                fields=[PacketField(f"{sdu_desc} Slice #2", "SDU Continuation Slice #2 (Bytes K..End)", 0, "M", "Bytes K..End", "38.322", "6.2.2", is_variable=True, variable_display_octets=6)],
            )
        )
        if maci_layer is not None:
            layers.append(maci_layer)

    if has_mac and mac_ce and is_uplink:
        layers.append(build_mac_ce_layer(is_uplink=True))
    if has_mac and mac_padding:
        layers.append(
            PacketLayer(
                layer_id="uu_mac_pad",
                title="MAC Padding Subheader + Padding (to TB size)",
                spec_number="38.321",
                clause="6.2.1",
                role="uu_mac",
                fields=[
                    PacketField("R/F/LCID=63", "LCID=63 (Padding Subheader, 1B)", 8, "O", "63", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
                    PacketField("MAC Padding", "Zero Padding to TB Size", 0, "O", "Pad", "38.321", "6.2.1", is_padding=True, is_variable=True, variable_display_octets=4),
                ],
            )
        )
    return layers


def _compose_uu_up(config: PacketConfig, preset: Any, depth_mode: str) -> ComposedPacket:
    is_ul = preset.preset_id == "uu_up_drb_ul"
    has_mac = _depth_includes(depth_mode, "mac")
    has_rlc = _depth_includes(depth_mode, "rlc")
    has_pdcp = _depth_includes(depth_mode, "pdcp")

    available_flags: List[FlagOption] = [
        FlagOption("sdap_header", "SDAP Header (1B)", "Include 1-octet TS 37.324 SDAP header (RDI/RQI/QFI or D/C/QFI).", "uu_l2", True),
        FlagOption("pdcp_sn_18", "PDCP 18-bit SN (3B)", "Use 18-bit PDCP DRB SN (3 octets) vs 12-bit SN (2 octets).", "uu_l2", False),
        FlagOption("pdcp_up_integrity", "PDCP UP MAC-I (4B)", "Enable UP integrity protection (32-bit MAC-I trailer).", "uu_l2", False),
        FlagOption("rlc_um", "RLC UM (vs AM)", "Use RLC Unacknowledged Mode instead of Acknowledged Mode.", "uu_l2", False),
        FlagOption("rlc_segmented", "RLC Segmented (2 TBs)", "Split RLC SDU across 2 MAC TBs (TB #1 SI=01, TB #2 SI=10 + 16-bit SO).", "uu_l2", False),
        FlagOption("mac_ce", "+ MAC CE (BSR / TA)", "Include MAC Control Element (DL: front of TB; UL: tail before padding).", "uu_l2", is_ul),
        FlagOption("mac_padding", "+ MAC Padding", "Append LCID=63 Padding Subheader + padding bytes.", "uu_l2", False),
    ]
    eff = {o.key: (bool(config.flags[o.key]) if o.key in config.flags else o.default_enabled) for o in available_flags}

    rlc_mode = "UM" if eff["rlc_um"] else "AM"
    sdap_on = bool(eff["sdap_header"])
    sn_18 = bool(eff["pdcp_sn_18"])
    mac_ce = bool(eff["mac_ce"])
    up_maci = bool(eff["pdcp_up_integrity"])
    segmented = bool(eff["rlc_segmented"])

    pdcp_layer = build_pdcp_header_layer(is_srb=False, sn_18=sn_18) if has_pdcp else None
    upper_layers = [build_sdap_header_layer(is_uplink=is_ul)] if sdap_on else []
    payload_layer = (
        PacketLayer(
            layer_id="uu_up_tpdu",
            title="Encapsulated User Plane IP / Ethernet Packet",
            spec_number="37.324",
            clause="6.3.2",
            role="payload",
            fields=[PacketField("Data", "User Plane IP / Ethernet PDU", 0, "M", "Variable", "37.324", "6.3.2", is_variable=True, variable_display_octets=8, def_name="Data")],
        )
        if config.include_tpdu
        else None
    )
    maci_layer = build_pdcp_maci_layer(is_srb=False) if (has_pdcp and up_maci) else None

    layers = _assemble_with_segmentation(
        is_uplink=is_ul,
        depth_mode=depth_mode,
        lcid=4,
        channel_label="DRB1 / DTCH",
        rlc_mode=rlc_mode,
        segmented=segmented,
        mac_ce=mac_ce,
        mac_padding=eff["mac_padding"],
        pdcp_layer=pdcp_layer,
        upper_layers=upper_layers,
        payload_layer=payload_layer,
        maci_layer=maci_layer,
        sdu_desc="IP/PDCP SDU",
    )

    transport_ch = "UL-SCH" if is_ul else "DL-SCH"
    ch_chain = (["IP Packet"] if config.include_tpdu else []) + (["SDAP (QFI)"] if sdap_on else []) + ["DRB1 (LCID=4)", f"PDCP ({'18b' if sn_18 else '12b'})", f"RLC {rlc_mode}", "DTCH", transport_ch]
    ov = sum(l.fixed_octets for l in layers)
    spec_refs = ["TS 38.300", "TS 37.324 (SDAP)"] + (["TS 38.323 (PDCP)"] if has_pdcp else []) + (["TS 38.322 (RLC)"] if has_rlc else []) + (["TS 38.321 (MAC)"] if has_mac else [])
    summary = (["MAC"] if has_mac else []) + ([f"RLC {rlc_mode}"] if has_rlc else []) + (["PDCP DRB"] if has_pdcp else []) + (["SDAP"] if sdap_on else []) + (["2-TB Segmented"] if segmented else [])

    return ComposedPacket(
        title=f"5G NR Uu User Plane Stack — {'Uplink' if is_ul else 'Downlink'} DRB",
        subtitle=preset.description,
        direction="UE → NG-RAN node" if is_ul else "NG-RAN node → UE",
        interface=f"Uu UP (DRB1 / DTCH → {transport_ch})",
        spec_refs=spec_refs,
        layers=layers,
        available_flags=available_flags,
        available_presets=RAN_UU_SCENARIO_PRESETS,
        config=PacketConfig(
            stack_family="ran_uu",
            target_frame="uu_up",
            preset_id=preset.preset_id,
            transport_mode=depth_mode,
            flags=eff,
            include_tpdu=config.include_tpdu,
            highlight_field=config.highlight_field,
        ),
        overhead_3gpp_octets=ov,
        overhead_transport_octets=0,
        overhead_total_octets=ov,
        ext_header_summaries=summary,
        warnings=[],
        channel_mapping_chain=ch_chain,
    )


def _unwrap_rrc_type_ref(raw_tref: str, ie_name: str) -> tuple[str, str]:
    """Return `(compact_label_type, clean_def_name)` for an ASN.1 RRC `type_reference`."""
    flat = " ".join((raw_tref or "").split()).strip()
    if not flat:
        return ("APER", ie_name)
    m_cont = re.search(r"CONTAINING\s+([A-Za-z0-9\-]+)", flat)
    if m_cont:
        inner = m_cont.group(1)
        return (f"CONTAINING {inner}", inner)
    m_sr = re.match(r"^SetupRelease\s*\{\s*([A-Za-z0-9\-]+)\s*\}", flat)
    if m_sr:
        inner = m_sr.group(1)
        return (f"SetupRelease<{inner}>", inner)
    m_seq_of = re.match(r"^SEQUENCE\s*(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\s+([A-Za-z0-9\-]+)", flat)
    if m_seq_of:
        inner = m_seq_of.group(1)
        return (f"SEQ OF {inner}", inner)
    base = re.sub(r"[\(\{\[].*$", "", flat).strip().rstrip(",").strip()
    if base.upper() in ("SEQUENCE", "CHOICE", "ENUMERATED", "INTEGER", "BIT STRING", "OCTET STRING", "BOOLEAN", "NULL"):
        return (base.upper(), ie_name)
    return (flat[:28], base or ie_name)


def compose_ran_uu_packet(config: PacketConfig, db: Any = None) -> ComposedPacket:
    """Compose an NR Uu Control Plane or User Plane packet stack with channel mapping, depth, and RLC segmentation."""
    depth_mode = (config.transport_mode or "uu_depth_mac").lower()
    if not depth_mode.startswith("uu_depth_"):
        depth_mode = "uu_depth_mac"

    preset = RAN_UU_PRESET_MAP.get(config.preset_id, RAN_UU_PRESET_MAP["uu_cp_srb1_rrc"])
    if config.target_frame == "uu_up" or preset.preset_id.startswith("uu_up_"):
        return _compose_uu_up(config, preset, depth_mode)

    default_msg = {
        "uu_cp_srb2_nas": "DLInformationTransfer",
        "uu_cp_srb0_ccch": "RRCSetup",
        "uu_cp_bcch_pcch": "Paging",
    }.get(preset.preset_id, "RRCReconfiguration")
    msg_name = (config.message_name or default_msg).strip()
    rrc_meta = load_rrc_message_from_db(db, msg_name, config.spec_version_38331)
    prof = infer_rrc_channel_profile(rrc_meta["message_name"], rrc_meta.get("description", ""), rrc_meta.get("direction", ""))
    effective_preset = RAN_UU_PRESET_MAP.get(config.preset_id, RAN_UU_PRESET_MAP[prof["preset_id"]])

    has_mac = _depth_includes(depth_mode, "mac") and prof["lcid"] is not None
    has_rlc = _depth_includes(depth_mode, "rlc")
    has_pdcp = _depth_includes(depth_mode, "pdcp") and prof["pdcp"]

    available_flags: List[FlagOption] = []
    if prof["pdcp"]:
        available_flags.extend([
            FlagOption("rlc_segmented", "RLC Segmented (2 TBs)", "Split RRC/PDCP SDU across 2 MAC TBs (SI=01 then SI=10 + 16-bit SO).", "uu_l2", False),
            FlagOption("mac_ce", "+ MAC CE (BSR / TA)", "Include MAC CE (DL: front of TB; UL: tail).", "uu_l2", False),
        ])
    if prof["supports_nas"]:
        available_flags.extend([
            FlagOption("expand_nas_pdu", "Expand dedicatedNAS-Message", "Decode encapsulated 5GS NAS message (TS 24.501) inline.", "uu_l2", effective_preset.default_flags.get("expand_nas_pdu", False)),
            FlagOption("nas_security_hdr", "5GMM Security Hdr (7B)", "Include 7-octet 5GMM Security Header inside dedicatedNAS-Message.", "uu_l2", True, requires_flag="expand_nas_pdu"),
        ])

    rrc_ies = rrc_meta.get("ies") or []
    hl_low = (config.highlight_field or "").strip().lower()
    opt_ies = [ie for ie in rrc_ies if not ie.get("presence", "M").upper().startswith("M")]
    opt_ies.sort(key=lambda x: (0 if hl_low and hl_low in x["name"].lower() else 1))
    opt_key_by_name: Dict[str, str] = {}
    for ie in opt_ies[:12]:
        key = "rrc_" + re.sub(r"[^a-z0-9]+", "_", ie["name"].lower()).strip("_")
        opt_key_by_name[ie["name"]] = key
        is_hl = bool(hl_low and (hl_low == ie["name"].lower() or hl_low in (ie.get("type_reference") or "").lower()))
        available_flags.append(
            FlagOption(key, ie["name"][:26], f"Include optional RRC IE '{ie['name']}' ({' '.join((ie.get('type_reference') or '').split())[:48]}).", "uu_rrc", is_hl)
        )

    eff: Dict[str, bool] = {o.key: (bool(config.flags[o.key]) if o.key in config.flags else o.default_enabled) for o in available_flags}
    for o in available_flags:
        if o.requires_flag and eff.get(o.key):
            eff[o.requires_flag] = True

    expand_nas = bool(prof["supports_nas"] and eff.get("expand_nas_pdu"))
    nas_msg = (config.nas_message_name or ("UL NAS TRANSPORT" if prof["is_uplink"] else "DL NAS TRANSPORT")).strip()

    has_trans_id = any(ie["name"] == "rrc-TransactionIdentifier" for ie in rrc_ies)
    if has_trans_id:
        rrc_fields: List[PacketField] = [
            PacketField("rrc-TransactionIdentifier", "TransID", 2, "M", "0..3", "38.331", rrc_meta["clause"], def_name="RRC-TransactionIdentifier"),
            PacketField("criticalExtensions", f"{rrc_meta['message_name']}-IEs", 6, "M", "CHOICE", "38.331", rrc_meta["clause"], def_name=rrc_meta["message_name"]),
        ]
    else:
        rrc_fields = [
            PacketField("message CHOICE", f"c1: {rrc_meta['message_name']}", 8, "M", "CHOICE", "38.331", rrc_meta["clause"], def_name=rrc_meta["message_name"]),
        ]
    for ie in rrc_ies:
        if ie["name"] == "rrc-TransactionIdentifier":
            continue
        is_mand = ie.get("presence", "M").upper().startswith("M")
        if not is_mand and not eff.get(opt_key_by_name.get(ie["name"], ""), False):
            if not (expand_nas and "nas" in ie["name"].lower()):
                continue
        if expand_nas and "nas" in ie["name"].lower():
            continue
        short_type, clean_def = _unwrap_rrc_type_ref(ie.get("type_reference", ""), ie["name"])
        rrc_fields.append(
            PacketField(ie["name"], f"{ie['name']} ({short_type})", 0, ie.get("presence", "M"), "APER", "38.331", rrc_meta["clause"], is_variable=True, variable_display_octets=3, def_name=clean_def)
        )

    upper_layers: List[PacketLayer] = [
        PacketLayer(
            layer_id="uu_rrc_pdu",
            title=f"RRC Message — {rrc_meta['message_name']} ({prof['bearer']} / {prof['logical']})",
            spec_number="38.331",
            clause=rrc_meta["clause"],
            role="ran_container",
            fields=rrc_fields,
        )
    ]
    if expand_nas:
        if eff.get("nas_security_hdr", True):
            upper_layers.append(build_nas_security_layer())
        nas_fields, nas_cl, _nv = load_nas_message_fields_from_db(db, nas_msg)
        upper_layers.append(PacketLayer("uu_nas_inner", f"5GS NAS Message (dedicatedNAS-Message) — {nas_msg}", "24.501", nas_cl or "8.2", "nas_message", fields=nas_fields))

    pdcp_layer = build_pdcp_header_layer(is_srb=True) if has_pdcp else None
    maci_layer = build_pdcp_maci_layer(is_srb=True) if has_pdcp else None
    segmented = bool(prof["pdcp"] and eff.get("rlc_segmented"))

    layers = _assemble_with_segmentation(
        is_uplink=prof["is_uplink"],
        depth_mode=depth_mode,
        lcid=prof["lcid"],
        channel_label=f"{prof['bearer']} / {prof['logical']}",
        rlc_mode=prof["rlc_mode"],
        segmented=segmented,
        mac_ce=bool(prof["pdcp"] and eff.get("mac_ce")),
        mac_padding=False,
        pdcp_layer=pdcp_layer,
        upper_layers=upper_layers,
        payload_layer=None,
        maci_layer=maci_layer,
        sdu_desc=rrc_meta["message_name"],
    )

    ch_chain = (["5GS NAS"] if expand_nas else []) + [f"RRC: {rrc_meta['message_name']}"] + ([prof["bearer"]] if prof["bearer"] != "N/A" else []) + (["PDCP (12b+MAC-I)"] if prof["pdcp"] else ["No PDCP"]) + [f"RLC {prof['rlc_mode']}", f"{prof['logical']}" + (f" (LCID={prof['lcid']})" if prof["lcid"] is not None else ""), prof["transport"]]
    ov = sum(l.fixed_octets for l in layers)
    spec_refs = ["TS 38.300", f"TS 38.331{' v' + rrc_meta['version'] if rrc_meta.get('version') else ''} Cl. {rrc_meta['clause']}"] + (["TS 38.323 (PDCP)"] if has_pdcp else []) + (["TS 38.322 (RLC)"] if has_rlc else []) + (["TS 38.321 (MAC)"] if has_mac else []) + (["TS 24.501 (NAS)"] if expand_nas else [])
    summary = ([f"MAC LCID={prof['lcid']}"] if has_mac else []) + ([f"RLC {prof['rlc_mode']}"] if has_rlc else []) + (["PDCP SRB"] if has_pdcp else []) + [rrc_meta["message_name"]] + ([nas_msg] if expand_nas else [])

    return ComposedPacket(
        title=f"5G NR Uu Control Plane Stack — {rrc_meta['message_name']}" + (f" + {nas_msg}" if expand_nas else ""),
        subtitle=effective_preset.description,
        direction=prof["direction"],
        interface=f"Uu CP ({prof['bearer']} / {prof['logical']} → {prof['transport']})",
        spec_refs=spec_refs,
        layers=layers,
        available_flags=available_flags,
        available_presets=RAN_UU_SCENARIO_PRESETS,
        config=PacketConfig(
            stack_family="ran_uu",
            target_frame="uu_cp",
            preset_id=effective_preset.preset_id,
            transport_mode=depth_mode,
            flags=eff,
            include_tpdu=False,
            message_name=rrc_meta["message_name"],
            nas_message_name=nas_msg,
            spec_version_38331=rrc_meta.get("version", ""),
            highlight_field=config.highlight_field,
        ),
        overhead_3gpp_octets=ov,
        overhead_transport_octets=0,
        overhead_total_octets=ov,
        ext_header_summaries=summary,
        warnings=[],
        channel_mapping_chain=ch_chain,
    )
