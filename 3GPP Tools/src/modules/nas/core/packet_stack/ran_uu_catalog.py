"""5G NR RAN Uu (TS 38.300 / 38.331 / 37.324 / 38.323 / 38.322 / 38.321) catalog & DB loaders."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from .models import PacketScenarioPreset

RAN_UU_SCENARIO_PRESETS: List[PacketScenarioPreset] = [
    PacketScenarioPreset(
        preset_id="uu_cp_srb2_nas",
        label="Uu CP: NAS over RRC (SRB2 · DCCH · RLC AM · PDCP · MAC)",
        interface="Uu CP (SRB2 / DCCH → DL-SCH/UL-SCH)",
        description="5GS NAS carried transparently inside RRC DLInformationTransfer/ULInformationTransfer over SRB2 (LCID 2, RLC AM, PDCP SRB ciphered + MAC-I).",
        default_flags={"expand_nas_pdu": True, "nas_security_hdr": True, "rlc_segmented": False, "mac_ce": False},
        transport_mode="uu_depth_mac",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="uu_cp_srb1_rrc",
        label="Uu CP: Dedicated RRC Signalling (SRB1 · DCCH · RLC AM · PDCP · MAC)",
        interface="Uu CP (SRB1 / DCCH → DL-SCH/UL-SCH)",
        description="Dedicated RRC connection control (e.g. RRCReconfiguration, SecurityModeCommand) on SRB1 (LCID 1, RLC AM, PDCP SRB + 32-bit MAC-I).",
        default_flags={"expand_nas_pdu": False, "nas_security_hdr": True, "rlc_segmented": False, "mac_ce": False},
        transport_mode="uu_depth_mac",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="uu_cp_srb0_ccch",
        label="Uu CP: Initial Access / Pre-Security (SRB0 · CCCH · RLC TM · No PDCP)",
        interface="Uu CP (SRB0 / CCCH → DL-SCH/UL-SCH)",
        description="Pre-security RRC messages (RRCSetupRequest, RRCSetup, RRCReject) on SRB0 (LCID 0): bypasses PDCP and uses 0-byte RLC TM.",
        default_flags={"rlc_segmented": False, "mac_ce": False},
        transport_mode="uu_depth_mac",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="uu_cp_bcch_pcch",
        label="Uu CP: Broadcast / Paging (BCCH / PCCH · RLC TM · No PDCP)",
        interface="Uu CP (BCCH/PCCH → DL-SCH/PCH)",
        description="System Information (MIB/SIB1) and Paging over BCCH/PCCH using Transparent Mode RLC (no PDCP, no MAC subheader on BCH/PCH).",
        default_flags={},
        transport_mode="uu_depth_mac",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="uu_up_drb_dl",
        label="Uu UP: Downlink User Data (DRB · SDAP · PDCP · RLC · MAC)",
        interface="Uu UP (DRB1 / DTCH → DL-SCH)",
        description="NR Uu Downlink User Plane stack: IP/Ethernet T-PDU → SDAP (RDI/RQI/QFI) → PDCP DRB → RLC AM/UM → MAC (with optional DL MAC CE at front).",
        default_flags={"sdap_header": True, "pdcp_sn_18": False, "pdcp_up_integrity": False, "rlc_um": False, "rlc_segmented": False, "mac_ce": False, "mac_padding": False},
        transport_mode="uu_depth_mac",
        include_tpdu=True,
    ),
    PacketScenarioPreset(
        preset_id="uu_up_drb_ul",
        label="Uu UP: Uplink User Data (DRB · SDAP · PDCP · RLC · MAC)",
        interface="Uu UP (DRB1 / DTCH → UL-SCH)",
        description="NR Uu Uplink User Plane stack: IP/Ethernet T-PDU → SDAP (D/C|R|QFI) → PDCP DRB → RLC AM/UM → MAC (UL MAC CE placed after SDU before padding).",
        default_flags={"sdap_header": True, "pdcp_sn_18": False, "pdcp_up_integrity": False, "rlc_um": False, "rlc_segmented": False, "mac_ce": True, "mac_padding": False},
        transport_mode="uu_depth_mac",
        include_tpdu=True,
    ),
]

RAN_UU_PRESET_MAP: Dict[str, PacketScenarioPreset] = {p.preset_id: p for p in RAN_UU_SCENARIO_PRESETS}

_SRB0_MESSAGES = frozenset({
    "rrcsetuprequest", "rrcsetup", "rrcreestablishmentrequest",
    "rrcresumerequest", "rrcresumerequest1", "rrcreject", "rrcsysteminforequest",
})
_SRB2_MESSAGES = frozenset({"dlinformationtransfer", "ulinformationtransfer"})
_BCCH_MESSAGES = frozenset({"mib", "sib1", "systeminformation"})


def infer_rrc_channel_profile(
    message_name: str,
    description: str = "",
    direction: str = "",
) -> Dict[str, Any]:
    """Infer SRB, PDCP usage, RLC mode, Logical Channel (LCID), and Transport Channel per TS 38.300/38.331."""
    msg = (message_name or "RRCReconfiguration").strip()
    low = msg.lower()
    desc = description or ""

    # Defaults for dedicated SRB1 RRC message
    bearer = "SRB1"
    pdcp = True
    rlc_mode = "AM"
    logical = "DCCH"
    lcid: Optional[int] = 1
    supports_nas = low in ("rrcreconfiguration", "rrcsetupcomplete", "rrcresumecomplete", "dlinformationtransfer", "ulinformationtransfer")
    preset_id = "uu_cp_srb1_rrc"

    if low in _SRB2_MESSAGES:
        bearer, pdcp, rlc_mode, logical, lcid = "SRB2", True, "AM", "DCCH", 2
        supports_nas, preset_id = True, "uu_cp_srb2_nas"
    elif low in _SRB0_MESSAGES:
        bearer, pdcp, rlc_mode, logical, lcid = "SRB0", False, "TM", "CCCH", 0
        supports_nas, preset_id = False, "uu_cp_srb0_ccch"
    elif low == "paging":
        bearer, pdcp, rlc_mode, logical, lcid = "N/A", False, "TM", "PCCH", None
        supports_nas, preset_id = False, "uu_cp_bcch_pcch"
    elif low in _BCCH_MESSAGES or low.startswith("sib"):
        bearer, pdcp, rlc_mode, logical, lcid = "N/A", False, "TM", "BCCH", None
        supports_nas, preset_id = False, "uu_cp_bcch_pcch"

    # Override from parsed TS 38.331 clause prose when available
    m_srb = re.search(r"Signalling radio bearer:\s*(SRB[0-4]|N/A)", desc, re.IGNORECASE)
    if m_srb:
        srb_val = m_srb.group(1).upper()
        bearer = srb_val
        if srb_val == "SRB0":
            pdcp, rlc_mode, logical, lcid, preset_id = False, "TM", "CCCH", 0, "uu_cp_srb0_ccch"
        elif srb_val == "SRB2":
            pdcp, rlc_mode, logical, lcid, preset_id = True, "AM", "DCCH", 2, "uu_cp_srb2_nas"
        elif srb_val == "SRB1":
            pdcp, rlc_mode, logical, lcid = True, "AM", "DCCH", 1
        elif srb_val == "N/A":
            pdcp, lcid, preset_id = False, None, "uu_cp_bcch_pcch"

    m_rlc = re.search(r"RLC-SAP:\s*(TM|UM|AM)", desc, re.IGNORECASE)
    if m_rlc:
        rlc_mode = m_rlc.group(1).upper()
    m_ch = re.search(r"Logical channel:\s*(CCCH|DCCH|BCCH|PCCH)", desc, re.IGNORECASE)
    if m_ch:
        logical = m_ch.group(1).upper()

    dir_low = (direction or "").lower()
    is_ul = (
        "ue to network" in dir_low
        or "ue →" in dir_low
        or low.startswith("ul")
        or low.endswith(("request", "request1", "complete", "response", "report", "information"))
        and "systeminformation" not in low
        and "dl" not in low[:2]
    )
    norm_dir = "UE → NG-RAN node" if is_ul else "NG-RAN node → UE"
    if logical == "PCCH":
        transport = "PCH"
    elif logical == "BCCH" and low == "mib":
        transport = "BCH"
    elif logical == "BCCH":
        transport = "DL-SCH"
    else:
        transport = "UL-SCH" if is_ul else "DL-SCH"

    return {
        "plane": "cp",
        "bearer": bearer,
        "pdcp": pdcp,
        "rlc_mode": rlc_mode,
        "logical": logical,
        "lcid": lcid,
        "transport": transport,
        "direction": norm_dir,
        "is_uplink": is_ul,
        "supports_nas": supports_nas,
        "preset_id": preset_id,
    }


def resolve_ran_uu_packet_target(
    spec_number: Optional[str],
    clause: str = "",
    ie_name: str = "",
    message_name: str = "",
) -> Optional[Dict[str, Any]]:
    """Resolve TS 38.331 / 37.324 / 38.323 / 38.322 / 38.321 selection into a `ran_uu` packet target."""
    clean_spec = str(spec_number or "").replace("TS", "").strip()
    if clean_spec not in ("38.331", "37.324", "38.323", "38.322", "38.321"):
        return None

    low_all = f"{ie_name} {message_name}".lower()
    if clean_spec == "38.331":
        cl_clean = (clause or "").strip()
        # Clause 7.x UE variables (VarMeasConfig, etc.) & 6.6.x PC5 sidelink are not Uu air-interface frames
        if cl_clean.startswith(("7.", "6.6")):
            return None
        msg = (message_name or "").strip()
        ie_clean = (ie_name or "").strip()
        # Clause 6.3.x non-SIB definitions (RadioBearerConfig, CellGroupConfig, MeasConfig, etc.) are IEs carried inside RRCReconfiguration
        if cl_clean.startswith("6.3") and not (msg or ie_clean).lower().startswith("sib"):
            highlight_ie = ie_clean or msg
            msg = "RRCReconfiguration"
            ie_clean = highlight_ie
        elif not msg and cl_clean.startswith("6.2.2") and ie_clean:
            msg = ie_clean
        if not msg:
            msg = "RRCReconfiguration"
        prof = infer_rrc_channel_profile(msg)
        highlight = ie_clean if ie_clean and ie_clean.lower() != msg.lower() else ""
        auto_flags: Dict[str, bool] = {}
        if prof["preset_id"] == "uu_cp_srb2_nas":
            auto_flags["expand_nas_pdu"] = True
            auto_flags["nas_security_hdr"] = True
        elif prof["supports_nas"] and ("nas" in highlight.lower()):
            auto_flags["expand_nas_pdu"] = True
            auto_flags["nas_security_hdr"] = True
        return {
            "stack_family": "ran_uu",
            "target_frame": "uu_cp",
            "preset_id": prof["preset_id"],
            "message_name": msg,
            "nas_message_name": "UL NAS TRANSPORT" if prof["is_uplink"] else "DL NAS TRANSPORT",
            "highlight_field": highlight,
            "auto_enable_flags": auto_flags,
        }

    if clean_spec == "37.324":
        is_ul = "ul " in low_all or "uplink" in low_all
        return {
            "stack_family": "ran_uu",
            "target_frame": "uu_up",
            "preset_id": "uu_up_drb_ul" if is_ul else "uu_up_drb_dl",
            "message_name": "",
            "nas_message_name": "",
            "highlight_field": ie_name.strip(),
            "auto_enable_flags": {"sdap_header": True},
        }

    # L2 PDCP (38.323) / RLC (38.322) / MAC (38.321)
    if "srb" in low_all:
        return {
            "stack_family": "ran_uu",
            "target_frame": "uu_cp",
            "preset_id": "uu_cp_srb1_rrc",
            "message_name": "RRCReconfiguration",
            "nas_message_name": "DL NAS TRANSPORT",
            "highlight_field": ie_name.strip(),
            "auto_enable_flags": {},
        }

    flags_up: Dict[str, bool] = {}
    if "18 bit" in low_all:
        flags_up["pdcp_sn_18"] = True
    if "mac-i" in low_all or "integrity" in low_all:
        flags_up["pdcp_up_integrity"] = True
    if "umd" in low_all:
        flags_up["rlc_um"] = True
    if "segment" in low_all or low_all.strip() in ("si", "so"):
        flags_up["rlc_segmented"] = True
    if "ce" in low_all.split() or "control element" in low_all:
        flags_up["mac_ce"] = True

    return {
        "stack_family": "ran_uu",
        "target_frame": "uu_up",
        "preset_id": "uu_up_drb_dl",
        "message_name": "",
        "nas_message_name": "",
        "highlight_field": ie_name.strip(),
        "auto_enable_flags": flags_up,
    }


_DEFAULT_RRC_MESSAGES: Dict[str, Dict[str, Any]] = {
    "DLInformationTransfer": {
        "clause": "6.2.2",
        "direction": "Network to UE",
        "description": "Signalling radio bearer: SRB2 or SRB1 RLC-SAP: AM Logical channel: DCCH Direction: Network to UE",
        "ies": [
            {"name": "rrc-TransactionIdentifier", "type_reference": "RRC-TransactionIdentifier", "presence": "M"},
            {"name": "dedicatedNAS-Message", "type_reference": "DedicatedNAS-Message", "presence": "O"},
        ],
    },
    "ULInformationTransfer": {
        "clause": "6.2.2",
        "direction": "UE to Network",
        "description": "Signalling radio bearer: SRB2 or SRB1 RLC-SAP: AM Logical channel: DCCH Direction: UE to Network",
        "ies": [
            {"name": "dedicatedNAS-Message", "type_reference": "DedicatedNAS-Message", "presence": "O"},
        ],
    },
    "RRCReconfiguration": {
        "clause": "6.2.2",
        "direction": "Network to UE",
        "description": "Signalling radio bearer: SRB1 RLC-SAP: AM Logical channel: DCCH Direction: Network to UE",
        "ies": [
            {"name": "rrc-TransactionIdentifier", "type_reference": "RRC-TransactionIdentifier", "presence": "M"},
            {"name": "radioBearerConfig", "type_reference": "RadioBearerConfig", "presence": "O"},
            {"name": "secondaryCellGroup", "type_reference": "CellGroupConfig", "presence": "O"},
            {"name": "measConfig", "type_reference": "MeasConfig", "presence": "O"},
            {"name": "dedicatedNAS-MessageList", "type_reference": "DedicatedNAS-Message", "presence": "O"},
        ],
    },
    "RRCSetup": {
        "clause": "6.2.2",
        "direction": "Network to UE",
        "description": "Signalling radio bearer: SRB0 RLC-SAP: TM Logical channel: CCCH Direction: Network to UE",
        "ies": [
            {"name": "rrc-TransactionIdentifier", "type_reference": "RRC-TransactionIdentifier", "presence": "M"},
            {"name": "radioBearerConfig", "type_reference": "RadioBearerConfig", "presence": "M"},
            {"name": "masterCellGroup", "type_reference": "CellGroupConfig", "presence": "M"},
        ],
    },
    "RRCSetupRequest": {
        "clause": "6.2.2",
        "direction": "UE to Network",
        "description": "Signalling radio bearer: SRB0 RLC-SAP: TM Logical channel: CCCH Direction: UE to Network",
        "ies": [
            {"name": "ue-Identity", "type_reference": "InitialUE-Identity", "presence": "M"},
            {"name": "establishmentCause", "type_reference": "EstablishmentCause", "presence": "M"},
            {"name": "spare", "type_reference": "BIT STRING (SIZE (1))", "presence": "M"},
        ],
    },
    "Paging": {
        "clause": "6.2.2",
        "direction": "Network to UE",
        "description": "Signalling radio bearer: N/A RLC-SAP: TM Logical channel: PCCH Direction: Network to UE",
        "ies": [{"name": "pagingRecordList", "type_reference": "PagingRecordList", "presence": "O"}],
    },
    "SIB1": {
        "clause": "6.2.2",
        "direction": "Network to UE",
        "description": "Signalling radio bearer: N/A RLC-SAP: TM Logical channel: BCCH Direction: Network to UE",
        "ies": [
            {"name": "cellSelectionInfo", "type_reference": "CellSelectionInfo", "presence": "C"},
            {"name": "cellAccessRelatedInfo", "type_reference": "CellAccessRelatedInfo", "presence": "M"},
        ],
    },
}


def load_rrc_message_from_db(db: Any, message_name: str, version_38331: str = "") -> Dict[str, Any]:
    """Load TS 38.331 RRC message metadata and depth=0 IEs from DB (or clean fallback)."""
    target = (message_name or "RRCReconfiguration").strip()
    if db is not None and hasattr(db, "_get_connection"):
        try:
            with db._get_connection() as conn:
                cur = conn.cursor()
                if version_38331:
                    row = cur.execute(
                        """
                        SELECT m.id, m.clause, m.message_name, m.direction, m.description, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '38.331' AND sv.version = ? AND LOWER(m.message_name) = LOWER(?)
                        LIMIT 1
                        """,
                        (version_38331, target),
                    ).fetchone()
                else:
                    row = None
                if not row:
                    row = cur.execute(
                        """
                        SELECT m.id, m.clause, m.message_name, m.direction, m.description, sv.version
                        FROM nas_messages m JOIN spec_versions sv ON m.version_id = sv.id
                        WHERE sv.spec_number = '38.331' AND LOWER(m.message_name) = LOWER(?)
                        ORDER BY sv.id DESC LIMIT 1
                        """,
                        (target,),
                    ).fetchone()
                if row:
                    ie_rows = cur.execute(
                        """
                        SELECT ie_name, type_reference, presence, field_description, COALESCE(depth, 0) AS depth
                        FROM message_ies
                        WHERE message_id = ? AND COALESCE(depth, 0) <= 1
                        ORDER BY order_index ASC
                        """,
                        (row["id"],),
                    ).fetchall()
                    d0 = [r for r in ie_rows if int(r["depth"]) == 0 and str(r["ie_name"] or "").strip()]
                    # Only unwrap depth=1 when depth=0 is a single `<msg>-IEs` wrapper (e.g. RRCSetupRequest)
                    if len(d0) == 1 and str(d0[0]["type_reference"] or "").strip().lower().endswith("-ies"):
                        chosen_rows = [r for r in ie_rows if int(r["depth"]) == 1 and str(r["ie_name"] or "").strip()] or d0
                    else:
                        chosen_rows = d0
                    ies = [
                        {
                            "name": str(r["ie_name"] or "").strip(),
                            "type_reference": str(r["type_reference"] or r["ie_name"] or "").strip(),
                            "presence": str(r["presence"] or "M").strip() or "M",
                            "description": str(r["field_description"] or "").strip(),
                        }
                        for r in chosen_rows
                    ]
                    if ies or str(row["message_name"]) in _DEFAULT_RRC_MESSAGES:
                        fb_ies = (_DEFAULT_RRC_MESSAGES.get(str(row["message_name"])) or {}).get("ies", [])
                        return {
                            "message_name": str(row["message_name"]),
                            "clause": str(row["clause"] or "6.2.2"),
                            "version": str(row["version"] or ""),
                            "direction": str(row["direction"] or ""),
                            "description": str(row["description"] or ""),
                            "ies": ies or [dict(x) for x in fb_ies],
                        }
                # Fallback: check ie_definitions in 38.331 (e.g. DLInformationTransfer, Paging, MIB)
                def_row = cur.execute(
                    """
                    SELECT d.clause, d.ie_name, d.raw_description, sv.version
                    FROM ie_definitions d JOIN spec_versions sv ON d.version_id = sv.id
                    WHERE sv.spec_number = '38.331' AND LOWER(d.ie_name) = LOWER(?)
                    ORDER BY sv.id DESC LIMIT 1
                    """,
                    (target,),
                ).fetchone()
                if def_row:
                    plain_desc = re.sub(r"<[^>]+>", " ", str(def_row["raw_description"] or ""))
                    plain_desc = " ".join(plain_desc.split())
                    fb_entry = _DEFAULT_RRC_MESSAGES.get(str(def_row["ie_name"])) or _DEFAULT_RRC_MESSAGES.get(target) or {}
                    return {
                        "message_name": str(def_row["ie_name"]),
                        "clause": str(def_row["clause"] or "6.2.2"),
                        "version": version_38331 or str(def_row["version"] or ""),
                        "direction": str(fb_entry.get("direction") or ""),
                        "description": plain_desc or str(fb_entry.get("description") or ""),
                        "ies": [dict(x) for x in (fb_entry.get("ies") or [])],
                    }
        except Exception:
            pass

    fb = _DEFAULT_RRC_MESSAGES.get(target) or _DEFAULT_RRC_MESSAGES["RRCReconfiguration"]
    m_name = target if target in _DEFAULT_RRC_MESSAGES else "RRCReconfiguration"
    return {
        "message_name": m_name,
        "clause": fb["clause"],
        "version": version_38331,
        "direction": fb["direction"],
        "description": fb["description"],
        "ies": [dict(x) for x in fb["ies"]],
    }
