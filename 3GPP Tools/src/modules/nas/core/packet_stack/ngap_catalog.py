"""N2 / NGAP (TS 38.413 + TS 38.412) and inline 5GS NAS-PDU (TS 24.501) catalog & DB loaders."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from .models import PacketField, PacketScenarioPreset

NGAP_SCENARIO_PRESETS: List[PacketScenarioPreset] = [
    PacketScenarioPreset(
        preset_id="n2_ue_signalling",
        label="N2 UE-Associated Signalling (SCTP Stream ≠ 0)",
        interface="N2 (NG-C)",
        description="Per-UE N2 NGAP procedure (gNB ↔ AMF) over SCTP (Port 38412, PPID = 60) on a dedicated non-zero UE SCTP stream (TS 38.412 Cl. 7).",
        default_flags={"sctp_non_ue_stream": False, "sctp_sack": False, "expand_nas_pdu": False, "nas_security_hdr": True},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="n2_ue_with_nas",
        label="N2 UE Signalling + Expanded 5GS NAS-PDU (TS 24.501)",
        interface="N1 / N2",
        description="Expands NGAP NAS-PDU (IE ID 38) inline with optional 7-octet 5GMM Security Header (EPD=0x7E, MAC, SQN) and TS 24.501 NAS fields.",
        default_flags={"sctp_non_ue_stream": False, "sctp_sack": False, "expand_nas_pdu": True, "nas_security_hdr": True},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="n2_non_ue",
        label="N2 Non-UE-Associated Signalling (SCTP Stream 0)",
        interface="N2 (NG-C)",
        description="Interface-management NGAP procedure (NGSetup, NGReset, AMFConfigUpdate, Paging) carried on common SCTP Stream 0 (TS 38.412 Cl. 7).",
        default_flags={"sctp_non_ue_stream": True, "sctp_sack": False, "expand_nas_pdu": False, "nas_security_hdr": False},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="n2_sctp_bundled_sack",
        label="N2 with Bundled SCTP SACK Chunk",
        interface="N2 (NG-C)",
        description="Demonstrates SCTP chunk bundling (RFC 9260): piggybacking a 16-octet SACK chunk alongside the NGAP DATA chunk in one IP packet.",
        default_flags={"sctp_non_ue_stream": False, "sctp_sack": True, "expand_nas_pdu": False, "nas_security_hdr": True},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
]

NGAP_PRESET_MAP: Dict[str, PacketScenarioPreset] = {p.preset_id: p for p in NGAP_SCENARIO_PRESETS}

NGAP_NON_UE_MESSAGE_PREFIXES: Tuple[str, ...] = (
    "NGSetup",
    "NGReset",
    "AMFConfigurationUpdate",
    "RANConfigurationUpdate",
    "ErrorIndication",
    "Overload",
    "Paging",
    "AMFStatusIndication",
    "WriteReplaceWarning",
    "PWSCancel",
    "RerouteNASRequest",
    "CellTrafficTrace",
    "UplinkRANConfigurationTransfer",
    "DownlinkRANConfigurationTransfer",
    "UplinkRIMInformationTransfer",
    "DownlinkRIMInformationTransfer",
)

SAMPLE_NAS_MESSAGES: List[str] = [
    "DL NAS TRANSPORT",
    "UL NAS TRANSPORT",
    "REGISTRATION REQUEST",
    "REGISTRATION ACCEPT",
    "AUTHENTICATION REQUEST",
    "SERVICE REQUEST",
    "PDU SESSION ESTABLISHMENT REQUEST",
    "PDU SESSION ESTABLISHMENT ACCEPT",
]

_DEFAULT_NGAP_MESSAGES: Dict[str, Dict[str, Any]] = {
    "DownlinkNASTransport": {
        "clause": "9.2.5.2",
        "direction": "AMF → NG-RAN node",
        "procedure_code": "4",
        "msg_criticality": "ignore",
        "pdu_choice": 0,
        "ies": [
            {"id": "10", "name": "AMF-UE-NGAP-ID", "type_reference": "AMF-UE-NGAP-ID", "presence": "M", "criticality": "reject", "field_description": "Uniquely identifies the UE association over NG within the AMF."},
            {"id": "85", "name": "RAN-UE-NGAP-ID", "type_reference": "RAN-UE-NGAP-ID", "presence": "M", "criticality": "reject", "field_description": "Uniquely identifies the UE association over NG within the NG-RAN node."},
            {"id": "48", "name": "OldAMF", "type_reference": "AMFName", "presence": "O", "criticality": "reject", "field_description": "Old AMF name."},
            {"id": "83", "name": "RANPagingPriority", "type_reference": "RANPagingPriority", "presence": "O", "criticality": "ignore", "field_description": "Paging priority for RRC Inactive."},
            {"id": "38", "name": "NAS-PDU", "type_reference": "NAS-PDU", "presence": "M", "criticality": "reject", "field_description": "Encapsulated 5GS NAS message (TS 24.501) carried transparently across N2."},
            {"id": "36", "name": "MobilityRestrictionList", "type_reference": "MobilityRestrictionList", "presence": "O", "criticality": "ignore", "field_description": "Roaming / area restrictions for the UE."},
        ],
    },
    "InitialUEMessage": {
        "clause": "9.2.5.1",
        "direction": "NG-RAN node → AMF",
        "procedure_code": "15",
        "msg_criticality": "ignore",
        "pdu_choice": 0,
        "ies": [
            {"id": "85", "name": "RAN-UE-NGAP-ID", "type_reference": "RAN-UE-NGAP-ID", "presence": "M", "criticality": "reject", "field_description": "Uniquely identifies the UE association within the NG-RAN node."},
            {"id": "38", "name": "NAS-PDU", "type_reference": "NAS-PDU", "presence": "M", "criticality": "reject", "field_description": "Initial 5GS NAS message from UE (e.g. REGISTRATION REQUEST / SERVICE REQUEST)."},
            {"id": "121", "name": "UserLocationInformation", "type_reference": "UserLocationInformation", "presence": "M", "criticality": "reject", "field_description": "NR CGI + TAI of the UE."},
            {"id": "90", "name": "RRCEstablishmentCause", "type_reference": "RRCEstablishmentCause", "presence": "M", "criticality": "ignore", "field_description": "Cause value from RRCSetupComplete."},
            {"id": "26", "name": "FiveG-S-TMSI", "type_reference": "FiveG-S-TMSI", "presence": "O", "criticality": "reject", "field_description": "5G-S-TMSI if provided by UE."},
            {"id": "112", "name": "UEContextRequest", "type_reference": "UEContextRequest", "presence": "O", "criticality": "ignore", "field_description": "Indicates NG-RAN requests UE context setup."},
        ],
    },
    "PDUSessionResourceSetupRequest": {
        "clause": "9.2.1.1",
        "direction": "AMF → NG-RAN node",
        "procedure_code": "29",
        "msg_criticality": "reject",
        "pdu_choice": 0,
        "ies": [
            {"id": "10", "name": "AMF-UE-NGAP-ID", "type_reference": "AMF-UE-NGAP-ID", "presence": "M", "criticality": "reject", "field_description": ""},
            {"id": "85", "name": "RAN-UE-NGAP-ID", "type_reference": "RAN-UE-NGAP-ID", "presence": "M", "criticality": "reject", "field_description": ""},
            {"id": "83", "name": "RANPagingPriority", "type_reference": "RANPagingPriority", "presence": "O", "criticality": "ignore", "field_description": ""},
            {"id": "38", "name": "NAS-PDU", "type_reference": "NAS-PDU", "presence": "O", "criticality": "reject", "field_description": "Optional NAS-PDU (e.g. DL NAS TRANSPORT carrying PDU SESSION ESTABLISHMENT ACCEPT)."},
            {"id": "74", "name": "PDUSessionResourceSetupListSUReq", "type_reference": "PDUSessionResourceSetupListSUReq", "presence": "M", "criticality": "reject", "field_description": "List of PDU Sessions to set up (contains UPF GTP-U TEID & QoS Flows)."},
        ],
    },
    "NGSetupRequest": {
        "clause": "9.2.6.1",
        "direction": "NG-RAN node → AMF",
        "procedure_code": "21",
        "msg_criticality": "reject",
        "pdu_choice": 0,
        "ies": [
            {"id": "27", "name": "GlobalRANNodeID", "type_reference": "GlobalRANNodeID", "presence": "M", "criticality": "reject", "field_description": "Global identifier of the NG-RAN node."},
            {"id": "82", "name": "RANNodeName", "type_reference": "RANNodeName", "presence": "O", "criticality": "ignore", "field_description": "Human-readable name of the NG-RAN node."},
            {"id": "102", "name": "SupportedTAList", "type_reference": "SupportedTAList", "presence": "M", "criticality": "reject", "field_description": "List of Tracking Areas and broadcast PLMNs / slices."},
            {"id": "21", "name": "DefaultPagingDRX", "type_reference": "PagingDRX", "presence": "M", "criticality": "ignore", "field_description": "Default paging DRX cycle."},
        ],
    },
}


def is_non_ue_ngap_message(message_name: str) -> bool:
    clean = (message_name or "").strip()
    return any(clean.startswith(prefix) for prefix in NGAP_NON_UE_MESSAGE_PREFIXES)


def infer_ngap_pdu_choice(message_name: str) -> int:
    clean = (message_name or "").strip()
    if clean.endswith(("Response", "Acknowledge", "Confirm", "Complete")):
        return 1  # successfulOutcome
    if clean.endswith("Failure"):
        return 2  # unsuccessfulOutcome
    return 0  # initiatingMessage


def default_nas_message_for_ngap(ngap_message_name: str) -> str:
    low = (ngap_message_name or "").lower()
    if "initialue" in low:
        return "REGISTRATION REQUEST"
    if "uplinknas" in low:
        return "UL NAS TRANSPORT"
    if "pdusessionresourcesetup" in low:
        return "PDU SESSION ESTABLISHMENT ACCEPT"
    return "DL NAS TRANSPORT"


def _normalize_direction(raw_dir: str, pdu_choice: int) -> str:
    clean = re.sub(r"\s+", " ", (raw_dir or "").strip())
    if "→" in clean or "↔" in clean:
        return clean
    m_dir = re.match(
        r"^(A-IoT CN node|NG-RAN node|NG-RAN|AMF)\s+(A-IoT CN node|NG-RAN node|NG-RAN|AMF)$",
        clean,
        re.IGNORECASE,
    )
    if m_dir:
        def _canon(ep: str) -> str:
            low = ep.strip().lower()
            if low == "amf":
                return "AMF"
            if low in ("ng-ran", "ng-ran node"):
                return "NG-RAN node"
            if low == "a-iot cn node":
                return "A-IoT CN node"
            return ep.strip()

        return f"{_canon(m_dir.group(1))} → {_canon(m_dir.group(2))}"
    if clean:
        return clean
    return "NG-RAN node → AMF" if pdu_choice == 1 else "AMF ↔ NG-RAN node"


def resolve_ngap_packet_target(
    spec_number: Optional[str],
    clause: str = "",
    ie_name: str = "",
    message_name: str = "",
) -> Optional[Dict[str, Any]]:
    """Resolve a Protocols selection in TS 38.413 / TS 38.412 into an N2 NGAP Packet Stack target."""
    clean_spec = str(spec_number or "").replace("TS", "").strip()
    if clean_spec not in ("38.413", "38.412"):
        return None

    if clean_spec == "38.412":
        return {
            "stack_family": "ngap",
            "target_frame": "ngap_pdu",
            "message_name": "NGSetupRequest",
            "preset_id": "n2_non_ue",
            "auto_enable_flags": {"sctp_non_ue_stream": True},
            "highlight_field": "",
        }

    msg = (message_name or "").strip()
    ie_clean = (ie_name or "").strip()
    if not msg:
        # In All IEs mode, if ie_name looks like an NGAP message (Clause 9.2.*), use it; else default
        if clause.startswith("9.2") and ie_clean:
            msg = ie_clean
        else:
            msg = "DownlinkNASTransport"

    highlight = ""
    if ie_clean and ie_clean.lower() != msg.lower() and ie_clean.lower() != "messagetype":
        highlight = ie_clean

    is_non_ue = is_non_ue_ngap_message(msg)
    preset_id = "n2_non_ue" if is_non_ue else "n2_ue_signalling"
    auto_flags: Dict[str, bool] = {"sctp_non_ue_stream": is_non_ue}
    if highlight.upper() == "NAS-PDU":
        preset_id = "n2_ue_with_nas"
        auto_flags["expand_nas_pdu"] = True
        auto_flags["nas_security_hdr"] = True

    return {
        "stack_family": "ngap",
        "target_frame": "ngap_pdu",
        "message_name": msg,
        "preset_id": preset_id,
        "auto_enable_flags": auto_flags,
        "highlight_field": highlight,
    }


def load_ngap_message_from_db(
    db: Any,
    message_name: str,
    version_38413: str = "",
) -> Dict[str, Any]:
    """Load NGAP message metadata, procedureCode, criticality, and depth=0 IEs from DB."""
    target = (message_name or "DownlinkNASTransport").strip()
    if db is not None and hasattr(db, "_get_connection"):
        try:
            with db._get_connection() as conn:
                cur = conn.cursor()
                if version_38413:
                    row = cur.execute(
                        """
                        SELECT m.id, m.version_id, m.clause, m.message_name, m.direction, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '38.413' AND sv.version = ? AND LOWER(m.message_name) = LOWER(?)
                        LIMIT 1
                        """,
                        (version_38413, target),
                    ).fetchone()
                else:
                    row = None
                if not row:
                    row = cur.execute(
                        """
                        SELECT m.id, m.version_id, m.clause, m.message_name, m.direction, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '38.413' AND LOWER(m.message_name) = LOWER(?)
                        ORDER BY sv.id DESC LIMIT 1
                        """,
                        (target,),
                    ).fetchone()
                if row:
                    ies_rows = cur.execute(
                        """
                        SELECT iei, ie_name, type_reference, presence, length, field_description
                        FROM message_ies
                        WHERE message_id = ? AND COALESCE(depth, 0) = 0
                        ORDER BY order_index ASC
                        """,
                        (row["id"],),
                    ).fetchall()
                    ie_clauses_by_name: Dict[str, str] = {}
                    try:
                        def_rows = cur.execute(
                            "SELECT ie_name, clause FROM ie_definitions WHERE version_id = ?",
                            (row["version_id"],),
                        ).fetchall()
                        for dr in def_rows:
                            k = str(dr["ie_name"] or "").strip().lower()
                            cl_val = str(dr["clause"] or "").strip()
                            if k and cl_val:
                                ie_clauses_by_name[k] = cl_val
                    except Exception:
                        ie_clauses_by_name = {}
                    proc_code = "0"
                    msg_crit = "reject"
                    ies: List[Dict[str, str]] = []
                    for r in ies_rows:
                        name = str(r["ie_name"] or "").strip()
                        if name == "MessageType":
                            proc_code = str(r["iei"] or "0").strip() or "0"
                            msg_crit = str(r["length"] or "reject").strip() or "reject"
                            continue
                        tref_raw = str(r["type_reference"] or name).strip()
                        tref = re.sub(r"\s*\(.*?\)\s*$", "", tref_raw).strip()
                        ie_clause = ""
                        def_name = ""
                        if tref and tref.lower() in ie_clauses_by_name:
                            ie_clause = ie_clauses_by_name[tref.lower()]
                            def_name = tref
                        elif name and name.lower() in ie_clauses_by_name:
                            ie_clause = ie_clauses_by_name[name.lower()]
                            def_name = name
                        else:
                            def_name = tref or name
                        ies.append({
                            "id": str(r["iei"] or "").strip(),
                            "name": name,
                            "type_reference": tref_raw,
                            "presence": str(r["presence"] or "M").strip() or "M",
                            "criticality": str(r["length"] or "ignore").strip() or "ignore",
                            "field_description": str(r["field_description"] or "").strip(),
                            "ie_clause": ie_clause,
                            "def_name": def_name,
                        })
                    choice = infer_ngap_pdu_choice(row["message_name"])
                    return {
                        "message_name": str(row["message_name"]),
                        "clause": str(row["clause"] or "9.2"),
                        "version": str(row["version"] or ""),
                        "direction": _normalize_direction(str(row["direction"] or ""), choice),
                        "procedure_code": proc_code,
                        "msg_criticality": msg_crit,
                        "pdu_choice": choice,
                        "ies": ies,
                    }
        except Exception:
            pass

    fb = _DEFAULT_NGAP_MESSAGES.get(target) or _DEFAULT_NGAP_MESSAGES["DownlinkNASTransport"]
    msg_name = target if target in _DEFAULT_NGAP_MESSAGES else "DownlinkNASTransport"
    return {
        "message_name": msg_name,
        "clause": fb["clause"],
        "version": version_38413,
        "direction": fb["direction"],
        "procedure_code": fb["procedure_code"],
        "msg_criticality": fb["msg_criticality"],
        "pdu_choice": fb["pdu_choice"],
        "ies": [dict(x) for x in fb["ies"]],
    }


def _nas_length_to_bits(length_str: str) -> Tuple[int, bool]:
    s = (length_str or "").strip().lower()
    if s in ("1/2", "0.5"):
        return 4, False
    if s.isdigit():
        return int(s) * 8, False
    m = re.match(r"^(\d+)$", s)
    if m:
        return int(m.group(1)) * 8, False
    return 0, True


def load_nas_message_fields_from_db(db: Any, nas_message_name: str) -> Tuple[List[PacketField], str, str]:
    """Load mandatory depth=0 5GS NAS fields from TS 24.501 in DB (or fallback)."""
    target = (nas_message_name or "DL NAS TRANSPORT").strip()
    if db is not None and hasattr(db, "_get_connection"):
        try:
            with db._get_connection() as conn:
                cur = conn.cursor()
                row = cur.execute(
                    """
                    SELECT m.id, m.clause, m.message_name, sv.version
                    FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                    WHERE sv.spec_number = '24.501' AND UPPER(m.message_name) = UPPER(?)
                    ORDER BY sv.id DESC LIMIT 1
                    """,
                    (target,),
                ).fetchone()
                if row:
                    ie_rows = cur.execute(
                        """
                        SELECT ie_name, presence, format, length, type_reference
                        FROM message_ies
                        WHERE message_id = ? AND COALESCE(depth, 0) = 0
                        ORDER BY order_index ASC
                        """,
                        (row["id"],),
                    ).fetchall()
                    fields: List[PacketField] = []
                    for r in ie_rows:
                        pres = str(r["presence"] or "").strip().upper()
                        if pres != "M":
                            continue
                        name = str(r["ie_name"] or "").strip()
                        bits, is_var = _nas_length_to_bits(str(r["length"] or ""))
                        cl_match = re.search(r"((?:9|D\.)[0-9A-Za-z\.]+)", str(r["type_reference"] or ""))
                        cl = cl_match.group(1) if cl_match else str(row["clause"] or "8.2")
                        short = name.replace(" message identity", " Msg Type").replace("Extended protocol discriminator", "EPD (0x7E)")
                        if "Spare half octet" in name:
                            short = "Spare"
                        fields.append(
                            PacketField(
                                name=name,
                                short_label=short,
                                bits=bits,
                                presence="M",
                                value_hint=str(r["format"] or ""),
                                spec_number="24.501",
                                clause=cl,
                                is_spare=("spare" in name.lower()),
                                is_variable=is_var,
                                variable_display_octets=6 if is_var else 0,
                                description=f"Mandatory TS 24.501 IE ({r['format']}, {r['length']}).",
                            )
                        )
                    if fields:
                        return fields, str(row["clause"] or "8.2"), str(row["version"] or "")
        except Exception:
            pass

    # Built-in fallback (DL NAS TRANSPORT)
    return [
        PacketField("Extended protocol discriminator", "EPD (0x7E)", 8, "M", "0x7E", "24.501", "9.2"),
        PacketField("Security header type", "SecHdr", 4, "M", "0000", "24.501", "9.3"),
        PacketField("Spare half octet", "Spare", 4, "M", "0000", "24.501", "9.5", is_spare=True),
        PacketField("DL NAS TRANSPORT message identity", "Msg Type (0x68)", 8, "M", "0x68", "24.501", "9.7"),
        PacketField("Payload container type", "Payload Type", 4, "M", "N1 SM / Policy", "24.501", "9.11.3.40"),
        PacketField("Spare half octet", "Spare", 4, "M", "0000", "24.501", "9.5", is_spare=True),
        PacketField("Payload container", "Payload Container (LV-E)", 0, "M", "3–65537 B", "24.501", "9.11.3.39", is_variable=True, variable_display_octets=8),
    ], "8.2.11.1.1", ""
