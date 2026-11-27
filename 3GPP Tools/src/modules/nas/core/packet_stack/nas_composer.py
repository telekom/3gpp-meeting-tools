"""5GS NAS (TS 24.501) + UE Policy (TS 24.526/588/555/514/578) + 5GSM/SOR/SMS packet composer."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from .models import ComposedPacket, FlagOption, PacketConfig, PacketField, PacketLayer
from .nas_catalog import (
    NAS_PRESET_MAP,
    NAS_SCENARIO_PRESETS,
    UE_POLICY_SPECS,
    load_nas_message_from_db,
)
from .ngap_catalog import load_ngap_message_from_db
from .ngap_composer import build_sctp_transport_layers


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", (text or "").lower()).strip("_")


def _nas_len_spec(length_str: str) -> Tuple[int, bool, int]:
    """Return `(bits, is_variable, var_display_octets)` for a TS 24.501 length string."""
    s = (length_str or "").strip().lower()
    if s in ("1/2", "0.5"):
        return 4, False, 0
    if s.isdigit():
        return int(s) * 8, False, 0
    return 0, True, 4


def _shorten_nas_label(ie: Dict[str, str], payload_container_label: str = "") -> str:
    name = ie["name"]
    low = name.lower()
    fmt = ie.get("format", "")
    iei = ie.get("iei", "")
    prefix = f"[IEI {iei}] " if iei and iei != "-" else ""
    if "extended protocol discriminator" in low:
        return "EPD = 0x7E (5GMM)" if "5gsm" not in (ie.get("field_description") or "").lower() else "EPD = 0x2E (5GSM)"
    if "security header type" in low:
        return "SecHdr=0"
    if "spare half octet" in low:
        return "Spare"
    if "message identity" in low:
        short_msg = re.sub(r"\s+message\s+identity$", "", name, flags=re.IGNORECASE)
        return f"MsgType ({short_msg})"
    if "payload container type" in low:
        return payload_container_label or "Payload Type"
    if low == "payload container":
        return "Payload Container (LV-E Len + Value)"
    return f"{prefix}{name} ({fmt}, {ie.get('length') or 'Var'})" if fmt and fmt != "V" else f"{prefix}{name}"


def build_nas_security_layer() -> PacketLayer:
    """TS 24.501 Clause 9.1 — 7-octet 5GMM Security Protected Header."""
    return PacketLayer(
        layer_id="nas_security_hdr",
        title="5GMM Security Protected NAS Header (7 octets)",
        spec_number="24.501",
        clause="9.1",
        role="nas_security",
        fields=[
            PacketField("Extended protocol discriminator", "EPD = 0x7E (5GMM)", 8, "M", "0x7E", "24.501", "9.2", flag_key="nas_security_hdr", def_name="Extended protocol discriminator"),
            PacketField("Spare half octet", "Spare", 4, "M", "0000", "24.501", "9.5", is_spare=True),
            PacketField("Security header type", "SecHdr=2 (Int+Ciph)", 4, "M", "0010", "24.501", "9.3", flag_key="nas_security_hdr", def_name="Security header type"),
            PacketField("Message authentication code", "NAS-MAC (32-bit)", 32, "M", "MAC", "24.501", "9.8", def_name="Message authentication code"),
            PacketField("Sequence number", "NAS-SQN (8-bit)", 8, "M", "SQN", "24.501", "9.10", def_name="Sequence number"),
        ],
    )


def build_nas_carrier_layers(transport_mode: str, *, is_uplink: bool = False) -> Tuple[List[PacketLayer], List[PacketLayer]]:
    """Construct optional N2 (NGAP+SCTP) or Uu (MAC→RLC AM→PDCP SRB→RRC) carrier wrappers."""
    mode = (transport_mode or "n1_standalone").lower()
    if mode == "over_n2_ngap":
        ngap_msg = "UplinkNASTransport" if is_uplink else "DownlinkNASTransport"
        meta = load_ngap_message_from_db(None, ngap_msg)
        prefix = build_sctp_transport_layers("collapsed", stream_id=1)
        prefix.append(
            PacketLayer(
                layer_id="carrier_ngap",
                title=f"N2 Carrier — NGAP {ngap_msg} (IE ID 38: NAS-PDU)",
                spec_number="38.413",
                clause=meta["clause"],
                role="ngap_envelope",
                fields=[
                    PacketField("NGAP-PDU Ext", "E=0", 1, "M", "0", "38.413", "9.4"),
                    PacketField("NGAP-PDU CHOICE", "00=InitMsg", 2, "M", "00", "38.413", "9.4"),
                    PacketField("APER Pad", "Pad", 5, "M", "0", "X.691", "-", is_spare=True),
                    PacketField("Procedure Code", f"ProcCode={meta['procedure_code']} ({ngap_msg})", 8, "M", meta["procedure_code"], "38.413", "9.3.1.1", def_name="MessageType"),
                    PacketField("AMF-UE-NGAP-ID", "IE 10: AMF-UE-NGAP-ID (APER)", 0, "M", "APER", "38.413", "9.3.3.1", is_variable=True, variable_display_octets=3, def_name="AMF-UE-NGAP-ID"),
                    PacketField("RAN-UE-NGAP-ID", "IE 85: RAN-UE-NGAP-ID (APER)", 0, "M", "APER", "38.413", "9.3.3.2", is_variable=True, variable_display_octets=3, def_name="RAN-UE-NGAP-ID"),
                    PacketField("NAS-PDU", "IE ID = 38 (NAS-PDU)", 16, "M", "38", "38.413", "9.3.3.4", def_name="NAS-PDU"),
                    PacketField("Criticality", "Crit=reject", 2, "M", "reject", "38.413", "9.4", def_name="Criticality"),
                    PacketField("APER Pad", "Pad", 6, "M", "0", "X.691", "-", is_spare=True),
                    PacketField("NAS-PDU Length", "NAS-PDU OCTET STRING Len (APER)", 8, "M", "Len", "X.691", "-"),
                ],
            )
        )
        return prefix, []

    if mode == "over_uu_rrc":
        rrc_msg = "ULInformationTransfer" if is_uplink else "DLInformationTransfer"
        rrc_clause = "6.2.2"
        prefix = [
            PacketLayer(
                layer_id="carrier_uu_mac",
                title="Uu MAC Subheader (SRB2 / DCCH)",
                spec_number="38.321",
                clause="6.2.1",
                role="uu_l2",
                fields=[
                    PacketField("R", "R=0", 1, "M", "0", "38.321", "6.2.1", is_spare=True),
                    PacketField("F", "F", 1, "M", "0/1", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
                    PacketField("LCID", "LCID = 2 (SRB2 / DCCH)", 6, "M", "2", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
                    PacketField("L", "MAC SDU Length (8/16-bit)", 8, "M", "Len", "38.321", "6.2.1", def_name="MAC subheader for DL-SCH and UL-SCH"),
                ],
            ),
            PacketLayer(
                layer_id="carrier_uu_rlc",
                title="Uu RLC AM Data Header",
                spec_number="38.322",
                clause="6.2.2.4",
                role="uu_l2",
                fields=[
                    PacketField("D/C", "D/C=1", 1, "M", "1", "38.322", "6.2.3.6", def_name="Data/Control (D/C) field"),
                    PacketField("P", "P", 1, "M", "0/1", "38.322", "6.2.3.7", def_name="Polling bit (P) field"),
                    PacketField("SI", "SI=00", 2, "M", "00", "38.322", "6.2.3.4", def_name="Segmentation Info (SI) field"),
                    PacketField("SN", "SN (hi 4b)", 4, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
                    PacketField("SN (lo)", "SN (bits 7..0)", 8, "M", "12-bit", "38.322", "6.2.3.3", def_name="Sequence Number (SN) field"),
                ],
            ),
            PacketLayer(
                layer_id="carrier_uu_pdcp",
                title="Uu PDCP SRB Header (12-bit SN)",
                spec_number="38.323",
                clause="6.2.2.1",
                role="uu_l2",
                fields=[
                    PacketField("R", "R=0000", 4, "M", "0", "38.323", "6.3.6", is_spare=True),
                    PacketField("PDCP SN", "PDCP SN (hi)", 4, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
                    PacketField("PDCP SN (lo)", "PDCP SN (bits 7..0)", 8, "M", "12-bit", "38.323", "6.3.2", def_name="PDCP SN"),
                ],
            ),
            PacketLayer(
                layer_id="carrier_uu_rrc",
                title=f"Uu RRC Envelope — {rrc_msg}",
                spec_number="38.331",
                clause=rrc_clause,
                role="ran_container",
                fields=[
                    PacketField("rrc-TransactionIdentifier", "TransID", 2, "M", "0..3", "38.331", rrc_clause),
                    PacketField("criticalExtensions", f"{rrc_msg}-IEs CHOICE", 6, "M", "0", "38.331", rrc_clause, def_name=rrc_msg),
                    PacketField("dedicatedNAS-Message", "dedicatedNAS-Message Len (APER)", 8, "M", "Len", "38.331", rrc_clause, def_name=rrc_msg),
                ],
            ),
        ]
        suffix = [
            PacketLayer(
                layer_id="carrier_uu_pdcp_maci",
                title="Uu PDCP Integrity Trailer",
                spec_number="38.323",
                clause="6.2.2.1",
                role="uu_l2",
                fields=[PacketField("MAC-I", "MAC-I (32-bit PDCP Integrity)", 32, "M", "MAC-I", "38.323", "6.3.4", def_name="MAC-I")],
            )
        ]
        return prefix, suffix

    return [], []


_URSP_TD_OPTIONS: List[Tuple[str, str, str, int, bool]] = [
    ("ursp_td_match_all", "TD: Match-all (0x01)", "Match-all type (0x01)", 8, False),
    ("ursp_td_dnn", "TD: DNN (0x88)", "DNN type (0x88 + Len + DNN)", 0, True),
    ("ursp_td_os_app", "TD: OS Id + App Id", "OS Id + OS App Id type (0x08)", 0, True),
    ("ursp_td_ipv4", "TD: IPv4 Remote (9B)", "IPv4 remote address type (0x10, 9 B)", 72, False),
    ("ursp_td_ipv6", "TD: IPv6 Prefix (18B)", "IPv6 remote address/prefix (0x21, 18 B)", 144, False),
    ("ursp_td_fqdn", "TD: Dest FQDN (0xA1)", "Destination FQDN type (0xA1)", 0, True),
    ("ursp_td_conn_cap", "TD: Conn Cap (0x90)", "Connection capabilities type (0x90)", 0, True),
    ("ursp_td_port", "TD: Remote Port (3B)", "Single remote port type (0x50, 3 B)", 24, False),
]

_URSP_RSD_OPTIONS: List[Tuple[str, str, str, int, bool]] = [
    ("ursp_rsd_ssc", "RSD: SSC Mode (0x01)", "SSC mode type (0x01, SSC 1/2/3)", 16, False),
    ("ursp_rsd_snssai", "RSD: S-NSSAI (0x02)", "S-NSSAI type (0x02, SST + SD)", 0, True),
    ("ursp_rsd_dnn", "RSD: DNN (0x04)", "DNN type (0x04, internet/ims)", 0, True),
    ("ursp_rsd_pdu_type", "RSD: PDU Type (0x08)", "PDU session type (0x08, IPv4/IPv6)", 16, False),
    ("ursp_rsd_access", "RSD: Access Pref (0x10)", "Preferred access type (0x10, 3GPP/non-3GPP)", 16, False),
    ("ursp_rsd_multi_access", "RSD: ATSSS Pref (0x11)", "Multi-access preference type (0x11)", 8, False),
    ("ursp_rsd_non_seamless", "RSD: NSWO (0x20)", "Non-seamless non-3GPP offload ind (0x20)", 8, False),
    ("ursp_rsd_location", "RSD: Location (0x40)", "Location criteria type (0x40, TAI/NCGI)", 0, True),
    ("ursp_rsd_time_window", "RSD: Time Win (17B)", "Time window type (0x80, 17 B)", 136, False),
]


def _build_ue_policy_layers(policy_key: str, flags: Dict[str, bool]) -> List[PacketLayer]:
    pinfo = UE_POLICY_SPECS.get(policy_key, UE_POLICY_SPECS["ursp"])
    code, pname, pspec, pclause, pdef = pinfo["code"], pinfo["name"], pinfo["spec"], pinfo["clause"], pinfo["def_name"]

    delivery_layer = PacketLayer(
        layer_id="ue_policy_delivery",
        title="UE Policy Delivery — MANAGE UE POLICY COMMAND + Section Management List",
        spec_number="24.501",
        clause="D.5.1.1.1",
        role="ue_policy_delivery",
        fields=[
            PacketField("PTI", "PTI (Proc Transaction ID)", 8, "M", "1..254", "24.501", "9.6", def_name="Procedure transaction identity"),
            PacketField("MANAGE UE POLICY COMMAND message identity", "MsgType = 0x01 (MANAGE UE POLICY CMD)", 8, "M", "0x01", "24.501", "D.6.1", def_name="UE policy delivery service message type"),
            PacketField("UE policy section management list", "Policy List Len (LV-E, 16-bit)", 16, "M", "Len", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("UE policy section management sublist Len", "Sublist Length (16-bit)", 16, "M", "Len", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("PLMN", "PLMN ID (MCC + MNC, 24-bit)", 24, "M", "MCC/MNC", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("Instruction contents Len", "Instruction Length (16-bit)", 16, "M", "Len", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("UPSC", "UPSC (UE Policy Section Code, 16-bit)", 16, "M", "UPSC", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("UE policy part Len", "Policy Part Length (16-bit)", 16, "M", "Len", "24.501", "D.6.2", def_name="UE policy section management list"),
            PacketField("Spare", "Spare", 4, "M", "0000", "24.501", "9.5", is_spare=True),
            PacketField("UE policy part type", f"Part={code} ({pname})", 4, "M", code, "24.501", "D.6.2", def_name="UE policy section management list"),
        ],
    )

    if policy_key == "ursp":
        rule_fields: List[PacketField] = [
            PacketField("Length of UE policy part contents", "Part Contents Len (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
            PacketField("Length of URSP rule", "URSP Rule Length (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
            PacketField("Precedence value of URSP rule", "Rule Precedence (8-bit)", 8, "M", "0..255", pspec, pclause, def_name=pdef),
            PacketField("Traffic descriptor", "Length of Traffic descriptor (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
        ]
        for key, _lbl, short, bits, is_var in _URSP_TD_OPTIONS:
            if flags.get(key):
                rule_fields.append(
                    PacketField(short, short, bits, "O", "", pspec, pclause, flag_key=key, is_variable=is_var, variable_display_octets=3 if is_var else 0, def_name=pdef)
                )
        rule_fields.extend([
            PacketField("Route selection descriptor list", "Length of RSD list (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
            PacketField("Route selection descriptor", "Length of RSD (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
            PacketField("Precedence value of RSD", "RSD Precedence (8-bit)", 8, "M", "0..255", pspec, pclause, def_name=pdef),
            PacketField("Length of RSD contents", "RSD Contents Len (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
        ])
        for key, _lbl, short, bits, is_var in _URSP_RSD_OPTIONS:
            if flags.get(key):
                rule_fields.append(
                    PacketField(short, short, bits, "O", "", pspec, pclause, flag_key=key, is_variable=is_var, variable_display_octets=3 if is_var else 0, def_name=pdef)
                )
    else:
        rule_fields = [
            PacketField(f"{pname} part length", f"{pname} Part Length (16-bit)", 16, "M", "Len", pspec, pclause, def_name=pdef),
            PacketField(pinfo["title"], f"{pinfo['title']} Contents", 0, "M", "TLV", pspec, pclause, is_variable=True, variable_display_octets=8, def_name=pdef),
        ]

    rule_layer = PacketLayer(
        layer_id="ue_policy_rule",
        title=f"{pinfo['title']} (TS {pspec} Cl. {pclause})",
        spec_number=pspec,
        clause=pclause,
        role="ue_policy_rule",
        fields=rule_fields,
    )
    return [delivery_layer, rule_layer]


def _build_nas_ie_fields(ies: List[Dict[str, str]], flags: Dict[str, bool], *, is_5gsm: bool = False, container_type_label: str = "") -> List[PacketField]:
    out: List[PacketField] = []
    for ie in ies:
        pres = ie.get("presence", "M").upper()
        key = f"opt_{_slug(ie['name'])}"
        if pres != "M" and not flags.get(key):
            continue
        bits, is_var, var_oct = _nas_len_spec(ie.get("length", ""))
        label = _shorten_nas_label(ie, container_type_label)
        if is_5gsm and "extended protocol discriminator" in ie["name"].lower():
            label = "EPD = 0x2E (5GSM)"
        out.append(
            PacketField(
                name=ie["name"],
                short_label=label,
                bits=bits,
                presence=pres,
                value_hint=ie.get("iei") or ie.get("format", ""),
                spec_number="24.501",
                clause=ie.get("clause") or "8.2",
                flag_key=key if pres != "M" else "",
                description=ie.get("field_description", ""),
                is_spare=("spare" in ie["name"].lower()),
                is_variable=is_var,
                variable_display_octets=var_oct,
                def_name=ie.get("def_name") or ie["name"],
            )
        )
    return out


def _collect_nas_flags(
    msg_meta: Dict[str, Any],
    preset: Any,
    payload_type: str,
    policy_key: str,
    inner_meta: Optional[Dict[str, Any]],
    hl_low: str,
) -> List[FlagOption]:
    flags: List[FlagOption] = [
        FlagOption("nas_security_hdr", "5GMM Security Hdr (7B)", "Include 7-octet 5GMM Security Protected Header (TS 24.501 Cl. 9.1).", "nas_env", True),
    ]
    is_transport = "NAS TRANSPORT" in msg_meta["message_name"].upper()
    if is_transport:
        flags.append(
            FlagOption("expand_payload_container", "Expand Payload Container", "Decode encapsulated Payload container (URSP / 5GSM / SOR / SMS).", "nas_env", bool(payload_type))
        )

    if payload_type == "ue_policy" and policy_key == "ursp":
        for k, lbl, desc, _b, _v in _URSP_TD_OPTIONS:
            flags.append(FlagOption(k, lbl, f"Include URSP Traffic Descriptor: {desc}.", "nas_inner", preset.default_flags.get(k, False), requires_flag="expand_payload_container"))
        for k, lbl, desc, _b, _v in _URSP_RSD_OPTIONS:
            flags.append(FlagOption(k, lbl, f"Include URSP Route Selection Descriptor: {desc}.", "nas_inner", preset.default_flags.get(k, False), requires_flag="expand_payload_container"))
        return flags

    if payload_type == "5gsm":
        combined_ies = [
            ie for ie in msg_meta.get("ies", [])
            if ie.get("presence", "M").upper() != "M" and "pdu session id" in ie["name"].lower()
        ] + list((inner_meta or {}).get("ies", []))
    else:
        combined_ies = list(msg_meta.get("ies", []))
    opt_ies = [ie for ie in combined_ies if ie.get("presence", "M").upper() != "M"]
    opt_ies.sort(key=lambda x: (0 if hl_low and hl_low in x["name"].lower() else 1))
    seen_keys: set[str] = set()
    for ie in opt_ies:
        key = f"opt_{_slug(ie['name'])}"
        if key in seen_keys:
            continue
        seen_keys.add(key)
        lbl = f"[{ie['iei']}] {ie['name']}" if ie.get("iei") else ie["name"]
        flags.append(
            FlagOption(
                key=key,
                label=lbl[:28],
                tooltip=f"Include {ie['presence']} IE '{ie['name']}' (Cl. {ie['clause']}, {ie['format']}).",
                group="nas_inner",
                default_enabled=bool(preset.default_flags.get(key) or (hl_low and hl_low == ie["name"].lower())),
            )
        )
        if len(seen_keys) >= 12:
            break
    return flags


def compose_nas_packet(config: PacketConfig, db: Any = None) -> ComposedPacket:
    """Compose an end-to-end 5GS NAS stack with encapsulated URSP / UE Policy / 5GSM / SOR / SMS."""
    preset = NAS_PRESET_MAP.get(config.preset_id, NAS_PRESET_MAP["nas_ue_policy_ursp"])
    if config.message_name:
        msg_name = config.message_name.strip()
    else:
        msg_name = "REGISTRATION REQUEST" if preset.preset_id == "nas_plain_5gmm" else "DL NAS TRANSPORT"

    msg_meta = load_nas_message_from_db(db, msg_name, config.spec_version_24501)
    is_transport = "NAS TRANSPORT" in msg_meta["message_name"].upper()
    payload_type = config.payload_container_type
    if is_transport and not payload_type:
        payload_type = {"nas_5gsm_in_5gmm": "5gsm", "nas_sor_container": "sor", "nas_sms_over_nas": "sms"}.get(preset.preset_id, "ue_policy")
    if not is_transport:
        payload_type = ""

    policy_key = config.policy_part_type if config.policy_part_type in UE_POLICY_SPECS else "ursp"
    sm_name = config.nas_message_name or ("PDU SESSION ESTABLISHMENT REQUEST" if "UL" in msg_meta["message_name"].upper() else "PDU SESSION ESTABLISHMENT ACCEPT")
    inner_meta = load_nas_message_from_db(db, sm_name, config.spec_version_24501) if payload_type == "5gsm" else None

    hl_low = (config.highlight_field or "").strip().lower()
    available_flags = _collect_nas_flags(msg_meta, preset, payload_type, policy_key, inner_meta, hl_low)

    effective: Dict[str, bool] = {}
    for opt in available_flags:
        effective[opt.key] = bool(config.flags[opt.key]) if opt.key in config.flags else opt.default_enabled
    for opt in available_flags:
        if opt.requires_flag and effective.get(opt.key):
            effective[opt.requires_flag] = True

    show_sec = bool(effective.get("nas_security_hdr", True))
    expand_inner = bool(is_transport and effective.get("expand_payload_container", True))

    type_map = {
        "ue_policy": "Type = 0101 (UE policy)",
        "5gsm": "Type = 0001 (N1 SM info)",
        "sor": "Type = 0100 (SOR)",
        "sms": "Type = 0010 (SMS)",
    }
    msg_ies = msg_meta["ies"]
    trailing_5gmm_fields: List[PacketField] = []
    if expand_inner and payload_type:
        # Split 5GMM IEs around Payload container so optional IEs (e.g. [12] PDU session ID) follow the inner payload
        pc_idx = next((i for i, ie in enumerate(msg_ies) if ie["name"].lower() == "payload container"), -1)
        if pc_idx >= 0:
            pre_ies = msg_ies[:pc_idx]
            post_ies = msg_ies[pc_idx + 1 :]
        else:
            pre_ies = list(msg_ies)
            post_ies = []
        outer_fields = _build_nas_ie_fields(pre_ies, effective, container_type_label=type_map.get(payload_type, ""))
        outer_fields.append(
            PacketField("Payload container", "Payload Container Len (LV-E, 16-bit)", 16, "M", "Len", "24.501", "9.11.3.39", def_name="Payload container")
        )
        trailing_5gmm_fields = _build_nas_ie_fields(post_ies, effective)
    else:
        outer_fields = _build_nas_ie_fields(msg_ies, effective, container_type_label=type_map.get(payload_type, ""))

    is_ul = "UE →" in msg_meta["direction"] or "UL " in msg_meta["message_name"].upper()
    prefix_layers, suffix_layers = build_nas_carrier_layers(config.transport_mode, is_uplink=is_ul)

    nas_layers: List[PacketLayer] = []
    if show_sec:
        nas_layers.append(build_nas_security_layer())
    nas_layers.append(
        PacketLayer(
            layer_id="nas_outer_5gmm",
            title=f"5GMM Message — {msg_meta['message_name']}",
            spec_number="24.501",
            clause=msg_meta["clause"],
            role="nas_message",
            fields=outer_fields,
        )
    )

    spec_refs = [f"TS 24.501{' v' + msg_meta['version'] if msg_meta.get('version') else ''} Cl. {msg_meta['clause']}"]
    chain: List[str] = (["5GMM Security (7B)"] if show_sec else []) + [msg_meta["message_name"]]
    title_suffix = ""

    if expand_inner and payload_type == "ue_policy":
        pinfo = UE_POLICY_SPECS[policy_key]
        nas_layers.extend(_build_ue_policy_layers(policy_key, effective))
        spec_refs.append(f"TS {pinfo['spec']} Cl. {pinfo['clause']}")
        chain.extend(["MANAGE UE POLICY CMD (D.5.1)", f"{pinfo['name']} (TS {pinfo['spec']})"])
        title_suffix = f" → UE Policy ({pinfo['name']})"
    elif expand_inner and payload_type == "5gsm" and inner_meta is not None:
        sm_fields = _build_nas_ie_fields(inner_meta["ies"], effective, is_5gsm=True)
        nas_layers.append(
            PacketLayer(
                layer_id="nas_inner_5gsm",
                title=f"5GSM Payload — {inner_meta['message_name']}",
                spec_number="24.501",
                clause=inner_meta["clause"],
                role="ue_policy_rule",
                fields=sm_fields,
            )
        )
        chain.append(inner_meta["message_name"])
        title_suffix = f" → {inner_meta['message_name']}"
    elif expand_inner and payload_type == "sor":
        nas_layers.append(
            PacketLayer(
                layer_id="nas_inner_sor",
                title="SOR Transparent Container (TS 24.501 Cl. 9.11.3.51)",
                spec_number="24.501",
                clause="9.11.3.51",
                role="ue_policy_delivery",
                fields=[
                    PacketField("Spare", "Spare", 4, "M", "0", "24.501", "9.11.3.51", is_spare=True),
                    PacketField("ACK", "ACK", 1, "M", "1", "24.501", "9.11.3.51"),
                    PacketField("List Type", "ListType", 1, "M", "1", "24.501", "9.11.3.51"),
                    PacketField("List Ind", "ListInd", 1, "M", "1", "24.501", "9.11.3.51"),
                    PacketField("Data Type", "SOR Type", 1, "M", "0", "24.501", "9.11.3.51"),
                    PacketField("SOR-MAC-IAUSF", "SOR-MAC-IAUSF (128-bit)", 128, "M", "128b", "24.501", "9.11.3.51", def_name="SOR transparent container"),
                    PacketField("Counter_SOR", "Counter_SOR (16-bit)", 16, "M", "16b", "24.501", "9.11.3.51", def_name="SOR transparent container"),
                    PacketField("PLMN ID & Access Tech List", "PLMN ID / Access Tech List", 0, "O", "Var", "24.501", "9.11.3.51", is_variable=True, variable_display_octets=6, def_name="SOR transparent container"),
                ],
            )
        )
        chain.append("SOR Container (9.11.3.51)")
        title_suffix = " → SOR Container"
    elif expand_inner and payload_type == "sms":
        nas_layers.append(
            PacketLayer(
                layer_id="nas_inner_sms",
                title="SMS over NAS — CP-DATA (TS 24.011)",
                spec_number="24.011",
                clause="7.2",
                role="ue_policy_delivery",
                fields=[
                    PacketField("TI / PD", "TI | PD = 0x9 (SMS)", 8, "M", "0x09", "24.011", "8.1"),
                    PacketField("CP Message Type", "MT = 0x01 (CP-DATA)", 8, "M", "0x01", "24.011", "8.4"),
                    PacketField("CP-User-Data Length", "RPDU Length (8-bit)", 8, "M", "Len", "24.011", "8.1.4"),
                    PacketField("RP-DATA / SMS TPDU", "RP-DATA + SMS-DELIVER/SUBMIT", 0, "M", "TPDU", "24.011", "7.3", is_variable=True, variable_display_octets=8),
                ],
            )
        )
        spec_refs.append("TS 24.011 Cl. 7.2")
        chain.append("CP-DATA / RP-DATA (24.011)")
        title_suffix = " → SMS CP-DATA"

    if trailing_5gmm_fields:
        nas_layers.append(
            PacketLayer(
                layer_id="nas_outer_5gmm_tail",
                title=f"5GMM Message — {msg_meta['message_name']} (Trailing Optional IEs)",
                spec_number="24.501",
                clause=msg_meta["clause"],
                role="nas_message",
                fields=trailing_5gmm_fields,
            )
        )

    mode = (config.transport_mode or "n1_standalone").lower()
    if mode == "over_n2_ngap":
        spec_refs.extend(["TS 38.413 (NGAP)", "TS 38.412 (SCTP)"])
        chain = ["IP/SCTP → NGAP"] + chain
    elif mode == "over_uu_rrc":
        spec_refs.extend(["TS 38.331 (RRC)", "TS 38.323 (PDCP)", "TS 38.322 (RLC)", "TS 38.321 (MAC)"])
        chain = ["MAC → RLC → PDCP → RRC"] + chain + ["PDCP MAC-I"]

    all_layers = prefix_layers + nas_layers + suffix_layers
    ov_nas = sum(l.fixed_octets for l in nas_layers)
    ov_car = sum(l.fixed_octets for l in prefix_layers + suffix_layers)
    carrier_lbl = {"over_n2_ngap": "N2 (NGAP + SCTP)", "over_uu_rrc": "Uu (MAC → RLC → PDCP → RRC)"}.get(mode, preset.interface)

    return ComposedPacket(
        title=f"5GS NAS Stack — {msg_meta['message_name']}{title_suffix}",
        subtitle=preset.description,
        direction=msg_meta["direction"],
        interface=carrier_lbl,
        spec_refs=spec_refs,
        layers=all_layers,
        available_flags=available_flags,
        available_presets=NAS_SCENARIO_PRESETS,
        config=PacketConfig(
            stack_family="nas",
            target_frame="nas_pdu",
            preset_id=preset.preset_id,
            transport_mode=mode,
            flags=effective,
            include_tpdu=False,
            message_name=msg_meta["message_name"],
            nas_message_name=sm_name,
            spec_version_24501=msg_meta.get("version", ""),
            spec_version_policy=config.spec_version_policy,
            payload_container_type=payload_type,
            policy_part_type=policy_key,
            highlight_field=config.highlight_field,
        ),
        overhead_3gpp_octets=ov_nas,
        overhead_transport_octets=ov_car,
        overhead_total_octets=ov_nas + ov_car,
        ext_header_summaries=chain,
        warnings=[],
    )
