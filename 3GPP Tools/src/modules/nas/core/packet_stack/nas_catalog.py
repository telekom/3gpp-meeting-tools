"""5GS NAS (TS 24.501) and UE Policy (TS 24.526/24.588/24.555/24.514/24.578) packet catalog & DB loaders."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from .models import PacketScenarioPreset

UE_POLICY_SPECS: Dict[str, Dict[str, str]] = {
    "ursp": {
        "code": "0001",
        "label": "0001 URSP (TS 24.526)",
        "name": "URSP",
        "spec": "24.526",
        "clause": "5.2",
        "def_name": "Encoding of UE policy part type URSP",
        "title": "UE Route Selection Policy (URSP) Rule",
    },
    "andsp": {
        "code": "0010",
        "label": "0010 ANDSP (TS 24.526)",
        "name": "ANDSP",
        "spec": "24.526",
        "clause": "5.3.1",
        "def_name": "Encoding of UE policy part type ANDSP — General",
        "title": "Access Network Discovery & Selection Policy (ANDSP)",
    },
    "v2xp": {
        "code": "0011",
        "label": "0011 V2XP (TS 24.588)",
        "name": "V2XP",
        "spec": "24.588",
        "clause": "5.2",
        "def_name": "Encoding of UE policy part type V2XP",
        "title": "V2X Policy (V2XP)",
    },
    "prosep": {
        "code": "0100",
        "label": "0100 ProSeP (TS 24.555)",
        "name": "ProSeP",
        "spec": "24.555",
        "clause": "5.2",
        "def_name": "Encoding of UE policy part type ProSeP",
        "title": "5G ProSe Policy (ProSeP)",
    },
    "a2xp": {
        "code": "0101",
        "label": "0101 A2XP (TS 24.578)",
        "name": "A2XP",
        "spec": "24.578",
        "clause": "5.2",
        "def_name": "Encoding of UE policy part type A2XP",
        "title": "A2X Policy (A2XP)",
    },
    "rslpp": {
        "code": "0110",
        "label": "0110 RSLPP (TS 24.514)",
        "name": "RSLPP",
        "spec": "24.514",
        "clause": "12.2",
        "def_name": "Encoding of UE policy part type RSLPP",
        "title": "Ranging & Sidelink Positioning Policy (RSLPP)",
    },
}

POLICY_SPEC_TO_KEY: Dict[str, str] = {
    "24.526": "ursp",
    "24.588": "v2xp",
    "24.555": "prosep",
    "24.578": "a2xp",
    "24.514": "rslpp",
}

SAMPLE_5GSM_MESSAGES: List[str] = [
    "PDU SESSION ESTABLISHMENT ACCEPT",
    "PDU SESSION ESTABLISHMENT REQUEST",
    "PDU SESSION MODIFICATION COMMAND",
    "PDU SESSION RELEASE COMMAND",
]

NAS_SCENARIO_PRESETS: List[PacketScenarioPreset] = [
    PacketScenarioPreset(
        preset_id="nas_ue_policy_ursp",
        label="DL NAS Transport → UE Policy (URSP / TS 24.526)",
        interface="N1 (PCF → UE via AMF)",
        description=(
            "5-layer policy delivery stack: 5GMM Security (7B) → DL NAS TRANSPORT (0101 UE policy) "
            "→ MANAGE UE POLICY COMMAND (Annex D.5/D.6) → URSP / ANDSP / V2XP / ProSeP / A2XP / RSLPP rule."
        ),
        default_flags={
            "nas_security_hdr": True,
            "expand_payload_container": True,
            "ursp_td_dnn": True,
            "ursp_rsd_snssai": True,
            "ursp_rsd_dnn": True,
            "ursp_rsd_ssc": True,
            "ursp_rsd_pdu_type": True,
        },
        transport_mode="n1_standalone",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="nas_5gsm_in_5gmm",
        label="NAS Transport → 5GSM PDU Session Signalling",
        interface="N1 (UE ↔ SMF via AMF)",
        description=(
            "Encapsulates a 5GSM PDU Session message (EPD=0x2E, Clause 8.3) inside DL/UL NAS TRANSPORT "
            "(Payload container type = 0001 N1 SM information, IEI PDU session ID)."
        ),
        default_flags={
            "nas_security_hdr": True,
            "expand_payload_container": True,
            "opt_pdu_session_id": True,
        },
        transport_mode="n1_standalone",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="nas_plain_5gmm",
        label="5GMM Signalling (Registration / Service / Auth)",
        interface="N1 (UE ↔ AMF)",
        description=(
            "Standalone 5GMM NAS message (TS 24.501 Clause 8.2, e.g. REGISTRATION REQUEST/ACCEPT, SERVICE REQUEST) "
            "with optional 7-octet ciphered/integrity-protected 5GMM Security Header."
        ),
        default_flags={"nas_security_hdr": True, "expand_payload_container": False},
        transport_mode="n1_standalone",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="nas_sor_container",
        label="DL NAS Transport → SOR Transparent Container",
        interface="N1 (AUSF/UDM → UE via AMF)",
        description=(
            "Steering of Roaming (SOR) delivery inside DL NAS TRANSPORT (Payload container type = 0100, "
            "TS 24.501 Cl. 9.11.3.51) with SOR-MAC-IAUSF, Counter_SOR, and preferred PLMN list."
        ),
        default_flags={"nas_security_hdr": True, "expand_payload_container": True},
        transport_mode="n1_standalone",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="nas_sms_over_nas",
        label="NAS Transport → SMS over NAS (CP-DATA / RP-DATA)",
        interface="N1 (SMSF ↔ UE via AMF)",
        description=(
            "SMS over NAS encapsulated inside DL/UL NAS TRANSPORT (Payload container type = 0010 SMS, "
            "TS 24.011 CP-DATA carrying RP-DATA / SMS TPDU)."
        ),
        default_flags={"nas_security_hdr": True, "expand_payload_container": True},
        transport_mode="n1_standalone",
        include_tpdu=False,
    ),
]

NAS_PRESET_MAP: Dict[str, PacketScenarioPreset] = {p.preset_id: p for p in NAS_SCENARIO_PRESETS}


def _policy_flags_from_hint(text: str) -> Dict[str, bool]:
    low = (text or "").lower()
    flags: Dict[str, bool] = {}
    if "match-all" in low or "match all" in low:
        flags["ursp_td_match_all"] = True
    if "os app" in low or "os id" in low:
        flags["ursp_td_os_app"] = True
    if "ipv4" in low:
        flags["ursp_td_ipv4"] = True
    if "ipv6" in low:
        flags["ursp_td_ipv6"] = True
    if "fqdn" in low:
        flags["ursp_td_fqdn"] = True
    if "connection cap" in low:
        flags["ursp_td_conn_cap"] = True
    if "remote port" in low:
        flags["ursp_td_port"] = True
    if "location" in low:
        flags["ursp_rsd_location"] = True
    if "time window" in low:
        flags["ursp_rsd_time_window"] = True
    if "preferred access" in low:
        flags["ursp_rsd_access"] = True
    if "multi-access" in low:
        flags["ursp_rsd_multi_access"] = True
    if "non-seamless" in low:
        flags["ursp_rsd_non_seamless"] = True
    return flags


def resolve_nas_packet_target(
    spec_number: Optional[str],
    clause: str = "",
    ie_name: str = "",
    message_name: str = "",
) -> Optional[Dict[str, Any]]:
    """Resolve TS 24.501 / UE Policy specs (24.526/588/555/514/578) into a `"nas"` packet target."""
    clean_spec = str(spec_number or "").replace("TS", "").strip()
    if clean_spec in POLICY_SPEC_TO_KEY:
        policy_key = POLICY_SPEC_TO_KEY[clean_spec]
        combined_low = f"{ie_name} {message_name}".lower()
        if clean_spec == "24.526" and (
            clause.strip().startswith("5.3")
            or any(k in combined_low for k in ("andsp", "wlansp", "n3an", "n3iwf", "epdg"))
        ):
            policy_key = "andsp"
        highlight = ie_name.strip()
        if highlight.lower() in (
            "ursp rule (ue route selection policy)",
            "andsp (access network discovery and selection policy)",
            "v2xp (v2x policy)",
            "prosep (5g prose policy)",
            "rslpp (ranging/sidelink positioning policy)",
            "a2xp (a2x policy)",
        ):
            highlight = ""
        auto_flags: Dict[str, bool] = {"nas_security_hdr": True, "expand_payload_container": True}
        auto_flags.update(_policy_flags_from_hint(f"{ie_name} {message_name}"))
        return {
            "stack_family": "nas",
            "target_frame": "nas_pdu",
            "preset_id": "nas_ue_policy_ursp",
            "message_name": "DL NAS TRANSPORT",
            "nas_message_name": "PDU SESSION ESTABLISHMENT ACCEPT",
            "payload_container_type": "ue_policy",
            "policy_part_type": policy_key,
            "highlight_field": highlight,
            "auto_enable_flags": auto_flags,
        }

    if clean_spec != "24.501":
        return None

    msg = (message_name or "").strip()
    ie_clean = (ie_name or "").strip()
    cl_clean = (clause or "").strip()
    if not msg and (cl_clean.startswith(("8.2", "8.3")) or cl_clean.upper().startswith("D.5")) and ie_clean:
        msg = ie_clean

    msg_u = msg.upper()
    ie_u = ie_clean.upper()
    highlight = ie_clean if ie_clean and ie_u != msg_u and "MESSAGE IDENTITY" not in ie_u else ""

    # 1. Annex D UE Policy messages/IEs
    if cl_clean.upper().startswith("D.") or "UE POLICY" in msg_u or "UE POLICY" in ie_u:
        carrier = "UL NAS TRANSPORT" if any(k in msg_u for k in ("COMPLETE", "REJECT", "INDICATION")) else "DL NAS TRANSPORT"
        return {
            "stack_family": "nas",
            "target_frame": "nas_pdu",
            "preset_id": "nas_ue_policy_ursp",
            "message_name": carrier,
            "nas_message_name": "PDU SESSION ESTABLISHMENT ACCEPT",
            "payload_container_type": "ue_policy",
            "policy_part_type": "ursp",
            "highlight_field": highlight,
            "auto_enable_flags": {"nas_security_hdr": True, "expand_payload_container": True},
        }

    # 2. 5GSM messages (Clause 8.3 or PDU SESSION ...)
    if cl_clean.startswith("8.3") or msg_u.startswith("PDU SESSION"):
        is_ul = any(k in msg_u for k in ("REQUEST", "COMPLETE"))
        return {
            "stack_family": "nas",
            "target_frame": "nas_pdu",
            "preset_id": "nas_5gsm_in_5gmm",
            "message_name": "UL NAS TRANSPORT" if is_ul else "DL NAS TRANSPORT",
            "nas_message_name": msg or "PDU SESSION ESTABLISHMENT ACCEPT",
            "payload_container_type": "5gsm",
            "policy_part_type": "ursp",
            "highlight_field": highlight,
            "auto_enable_flags": {"nas_security_hdr": True, "expand_payload_container": True, "opt_pdu_session_id": True},
        }

    # 3. SOR Transparent Container (9.11.3.51)
    if cl_clean == "9.11.3.51" or "SOR TRANSPARENT" in ie_u:
        return {
            "stack_family": "nas",
            "target_frame": "nas_pdu",
            "preset_id": "nas_sor_container",
            "message_name": "DL NAS TRANSPORT",
            "nas_message_name": "PDU SESSION ESTABLISHMENT ACCEPT",
            "payload_container_type": "sor",
            "policy_part_type": "ursp",
            "highlight_field": highlight,
            "auto_enable_flags": {"nas_security_hdr": True, "expand_payload_container": True},
        }

    # 4. DL/UL NAS TRANSPORT or Payload container (9.11.3.39 / 9.11.3.40)
    if msg_u in ("DL NAS TRANSPORT", "UL NAS TRANSPORT") or cl_clean in ("9.11.3.39", "9.11.3.40") or "PAYLOAD CONTAINER" in ie_u:
        carrier = msg or "DL NAS TRANSPORT"
        is_ul = "UL" in carrier.upper()
        return {
            "stack_family": "nas",
            "target_frame": "nas_pdu",
            "preset_id": "nas_5gsm_in_5gmm" if is_ul else "nas_ue_policy_ursp",
            "message_name": carrier,
            "nas_message_name": "PDU SESSION ESTABLISHMENT REQUEST" if is_ul else "PDU SESSION ESTABLISHMENT ACCEPT",
            "payload_container_type": "5gsm" if is_ul else "ue_policy",
            "policy_part_type": "ursp",
            "highlight_field": highlight,
            "auto_enable_flags": {"nas_security_hdr": True, "expand_payload_container": True},
        }

    # 5. Any other 5GMM message / IE (default to selected message or REGISTRATION REQUEST)
    return {
        "stack_family": "nas",
        "target_frame": "nas_pdu",
        "preset_id": "nas_plain_5gmm",
        "message_name": msg or "REGISTRATION REQUEST",
        "nas_message_name": "PDU SESSION ESTABLISHMENT ACCEPT",
        "payload_container_type": "",
        "policy_part_type": "ursp",
        "highlight_field": highlight,
        "auto_enable_flags": {"nas_security_hdr": True, "expand_payload_container": False},
    }


_DEFAULT_NAS_MESSAGES: Dict[str, Dict[str, Any]] = {
    "DL NAS TRANSPORT": {
        "clause": "8.2.11.1.1",
        "direction": "network → UE",
        "ies": [
            {"iei": "", "name": "Extended protocol discriminator", "clause": "9.2", "def_name": "Extended protocol discriminator", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Security header type", "clause": "9.3", "def_name": "Security header type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Spare half octet", "clause": "9.5", "def_name": "Spare half octet", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "DL NAS TRANSPORT message identity", "clause": "9.7", "def_name": "Message type", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Payload container type", "clause": "9.11.3.40", "def_name": "Payload container type", "presence": "M", "format": "V", "length": "1/2", "field_description": "Identifies encapsulated payload type (0001 N1 SM, 0010 SMS, 0100 SOR, 0101 UE policy)."},
            {"iei": "", "name": "Spare half octet", "clause": "9.5", "def_name": "Spare half octet", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Payload container", "clause": "9.11.3.39", "def_name": "Payload container", "presence": "M", "format": "LV-E", "length": "3-65537", "field_description": "2-octet length + encapsulated NAS payload."},
            {"iei": "12", "name": "PDU session ID", "clause": "9.11.3.41", "def_name": "PDU session identity 2", "presence": "C", "format": "TV", "length": "2", "field_description": "Present if Payload container carries 5GSM signalling."},
            {"iei": "24", "name": "Additional information", "clause": "9.11.2.1", "def_name": "Additional information", "presence": "O", "format": "TLV", "length": "3-n", "field_description": ""},
            {"iei": "58", "name": "5GMM cause", "clause": "9.11.3.2", "def_name": "5GMM cause", "presence": "O", "format": "TV", "length": "2", "field_description": ""},
            {"iei": "37", "name": "Back-off timer value", "clause": "9.11.2.5", "def_name": "GPRS timer 3", "presence": "O", "format": "TLV", "length": "3", "field_description": ""},
        ],
    },
    "UL NAS TRANSPORT": {
        "clause": "8.2.10.1.1",
        "direction": "UE → network",
        "ies": [
            {"iei": "", "name": "Extended protocol discriminator", "clause": "9.2", "def_name": "Extended protocol discriminator", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Security header type", "clause": "9.3", "def_name": "Security header type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Spare half octet", "clause": "9.5", "def_name": "Spare half octet", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "UL NAS TRANSPORT message identity", "clause": "9.7", "def_name": "Message type", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Payload container type", "clause": "9.11.3.40", "def_name": "Payload container type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Spare half octet", "clause": "9.5", "def_name": "Spare half octet", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Payload container", "clause": "9.11.3.39", "def_name": "Payload container", "presence": "M", "format": "LV-E", "length": "3-65537", "field_description": ""},
            {"iei": "12", "name": "PDU session ID", "clause": "9.11.3.41", "def_name": "PDU session identity 2", "presence": "C", "format": "TV", "length": "2", "field_description": ""},
            {"iei": "8-", "name": "Request type", "clause": "9.11.3.47", "def_name": "Request type", "presence": "O", "format": "TV", "length": "1", "field_description": ""},
            {"iei": "22", "name": "S-NSSAI", "clause": "9.11.2.8", "def_name": "S-NSSAI", "presence": "O", "format": "TLV", "length": "3-10", "field_description": ""},
            {"iei": "25", "name": "DNN", "clause": "9.11.2.1B", "def_name": "DNN", "presence": "O", "format": "TLV", "length": "3-102", "field_description": ""},
        ],
    },
    "REGISTRATION REQUEST": {
        "clause": "8.2.6.1.1",
        "direction": "UE → network",
        "ies": [
            {"iei": "", "name": "Extended protocol discriminator", "clause": "9.2", "def_name": "Extended protocol discriminator", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Security header type", "clause": "9.3", "def_name": "Security header type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Spare half octet", "clause": "9.5", "def_name": "Spare half octet", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "REGISTRATION REQUEST message identity", "clause": "9.7", "def_name": "Message type", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "5GS registration type", "clause": "9.11.3.7", "def_name": "5GS registration type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "ngKSI", "clause": "9.11.3.32", "def_name": "NAS key set identifier", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "5GS mobile identity", "clause": "9.11.3.4", "def_name": "5GS mobile identity", "presence": "M", "format": "LV-E", "length": "6-n", "field_description": "SUCI or 5G-GUTI."},
            {"iei": "10", "name": "5GMM capability", "clause": "9.11.3.1", "def_name": "5GMM capability", "presence": "O", "format": "TLV", "length": "3-15", "field_description": ""},
            {"iei": "2E", "name": "UE security capability", "clause": "9.11.3.54", "def_name": "UE security capability", "presence": "O", "format": "TLV", "length": "4-10", "field_description": ""},
            {"iei": "2F", "name": "Requested NSSAI", "clause": "9.11.3.37", "def_name": "NSSAI", "presence": "O", "format": "TLV", "length": "4-74", "field_description": ""},
        ],
    },
    "PDU SESSION ESTABLISHMENT ACCEPT": {
        "clause": "8.3.2.1.1",
        "direction": "network → UE",
        "ies": [
            {"iei": "", "name": "Extended protocol discriminator", "clause": "9.2", "def_name": "Extended protocol discriminator", "presence": "M", "format": "V", "length": "1", "field_description": "0x2E (5GSM)."},
            {"iei": "", "name": "PDU session ID", "clause": "9.4", "def_name": "PDU session identity", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "PTI", "clause": "9.6", "def_name": "Procedure transaction identity", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "PDU SESSION ESTABLISHMENT ACCEPT message identity", "clause": "9.7", "def_name": "Message type", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Selected PDU session type", "clause": "9.11.4.11", "def_name": "PDU session type", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Selected SSC mode", "clause": "9.11.4.16", "def_name": "SSC mode", "presence": "M", "format": "V", "length": "1/2", "field_description": ""},
            {"iei": "", "name": "Authorized QoS rules", "clause": "9.11.4.13", "def_name": "QoS rules", "presence": "M", "format": "LV-E", "length": "6-65538", "field_description": ""},
            {"iei": "", "name": "Session AMBR", "clause": "9.11.4.14", "def_name": "Session-AMBR", "presence": "M", "format": "LV", "length": "7", "field_description": ""},
            {"iei": "29", "name": "PDU address", "clause": "9.11.4.10", "def_name": "PDU address", "presence": "O", "format": "TLV", "length": "7-31", "field_description": ""},
            {"iei": "22", "name": "S-NSSAI", "clause": "9.11.2.8", "def_name": "S-NSSAI", "presence": "O", "format": "TLV", "length": "3-10", "field_description": ""},
            {"iei": "25", "name": "DNN", "clause": "9.11.2.1B", "def_name": "DNN", "presence": "O", "format": "TLV", "length": "3-102", "field_description": ""},
        ],
    },
    "PDU SESSION ESTABLISHMENT REQUEST": {
        "clause": "8.3.1.1.1",
        "direction": "UE → network",
        "ies": [
            {"iei": "", "name": "Extended protocol discriminator", "clause": "9.2", "def_name": "Extended protocol discriminator", "presence": "M", "format": "V", "length": "1", "field_description": "0x2E (5GSM)."},
            {"iei": "", "name": "PDU session ID", "clause": "9.4", "def_name": "PDU session identity", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "PTI", "clause": "9.6", "def_name": "Procedure transaction identity", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "PDU SESSION ESTABLISHMENT REQUEST message identity", "clause": "9.7", "def_name": "Message type", "presence": "M", "format": "V", "length": "1", "field_description": ""},
            {"iei": "", "name": "Integrity protection maximum data rate", "clause": "9.11.4.7", "def_name": "Integrity protection maximum data rate", "presence": "M", "format": "V", "length": "2", "field_description": ""},
            {"iei": "9-", "name": "PDU session type", "clause": "9.11.4.11", "def_name": "PDU session type", "presence": "O", "format": "TV", "length": "1", "field_description": ""},
            {"iei": "A-", "name": "SSC mode", "clause": "9.11.4.16", "def_name": "SSC mode", "presence": "O", "format": "TV", "length": "1", "field_description": ""},
        ],
    },
}


def _split_type_ref(type_reference: str, fallback_clause: str) -> Tuple[str, str]:
    tr = (type_reference or "").strip()
    m = re.search(r"\b((?:9|D)(?:\.[0-9A-Za-z]+)+)\b", tr)
    cl = m.group(1) if m else fallback_clause
    def_name = re.sub(r"\s+(?:9|D)\.[0-9A-Za-z\.]+.*$", "", tr).strip()
    return cl, def_name


def load_nas_message_from_db(db: Any, message_name: str, version_24501: str = "") -> Dict[str, Any]:
    """Load TS 24.501 message metadata and depth=0 IEs from DB (or clean fallback)."""
    target = (message_name or "DL NAS TRANSPORT").strip()
    if db is not None and hasattr(db, "_get_connection"):
        try:
            with db._get_connection() as conn:
                cur = conn.cursor()
                if version_24501:
                    row = cur.execute(
                        """
                        SELECT m.id, m.clause, m.message_name, m.direction, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '24.501' AND sv.version = ? AND UPPER(m.message_name) = UPPER(?)
                        LIMIT 1
                        """,
                        (version_24501, target),
                    ).fetchone()
                else:
                    row = None
                if not row:
                    row = cur.execute(
                        """
                        SELECT m.id, m.clause, m.message_name, m.direction, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '24.501' AND UPPER(m.message_name) = UPPER(?)
                        ORDER BY sv.id DESC LIMIT 1
                        """,
                        (target,),
                    ).fetchone()
                if row:
                    ie_rows = cur.execute(
                        """
                        SELECT iei, ie_name, type_reference, presence, format, length, field_description
                        FROM message_ies
                        WHERE message_id = ? AND COALESCE(depth, 0) = 0
                        ORDER BY order_index ASC
                        """,
                        (row["id"],),
                    ).fetchall()
                    msg_cl = str(row["clause"] or "8.2")
                    ies: List[Dict[str, str]] = []
                    seen: set[tuple[str, str]] = set()
                    for r in ie_rows:
                        name = str(r["ie_name"] or "").strip()
                        iei = str(r["iei"] or "").strip()
                        sig = (iei, name.lower())
                        if not name or sig in seen:
                            continue
                        seen.add(sig)
                        cl, def_name = _split_type_ref(str(r["type_reference"] or ""), msg_cl)
                        ies.append({
                            "iei": iei,
                            "name": name,
                            "clause": cl,
                            "def_name": def_name or name,
                            "presence": str(r["presence"] or "M").strip() or "M",
                            "format": str(r["format"] or "V").strip(),
                            "length": str(r["length"] or "").strip(),
                            "field_description": str(r["field_description"] or "").strip(),
                        })
                    if ies:
                        return {
                            "message_name": str(row["message_name"]),
                            "clause": msg_cl,
                            "version": str(row["version"] or ""),
                            "direction": str(row["direction"] or "AMF ↔ UE"),
                            "ies": ies,
                        }
        except Exception:
            pass

    fb = _DEFAULT_NAS_MESSAGES.get(target.upper()) or _DEFAULT_NAS_MESSAGES["DL NAS TRANSPORT"]
    m_name = target.upper() if target.upper() in _DEFAULT_NAS_MESSAGES else "DL NAS TRANSPORT"
    return {
        "message_name": m_name,
        "clause": fb["clause"],
        "version": version_24501,
        "direction": fb["direction"],
        "ies": [dict(x) for x in fb["ies"]],
    }
