"""5G User Plane catalog, target resolution, release capability lookup, and TS 38.415 frame builders."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

from .models import FlagOption, PacketField, PacketScenarioPreset

# -----------------------------------------------------------------------------
# Scenario Presets (5G UP Interfaces: N3/N9, Xn-U, F1-U, Dummy G-PDU, End Marker)
# -----------------------------------------------------------------------------

UP_SCENARIO_PRESETS: List[PacketScenarioPreset] = [
    PacketScenarioPreset(
        preset_id="n3_n9",
        label="N3 / N9 (UPF ↔ NG-RAN)",
        interface="N3, N9",
        description="Standard 5GS User Plane G-PDU with GTP-U Header and PDU Session Container (0x85).",
        default_flags={"gtpu_seq": False, "ext_pdu_session": True},
        transport_mode="collapsed",
        include_tpdu=True,
    ),
    PacketScenarioPreset(
        preset_id="n3_n9_pdu_set",
        label="N3 / N9 + PDU Set Chained (0x85 → 0x04)",
        interface="N3, N9",
        description="Rel-18/19 XR/media G-PDU chaining PDU Session Container (0x85) and PDU Set Information Container (0x04).",
        default_flags={"gtpu_seq": False, "ext_pdu_session": True, "ext_pdu_set": True},
        transport_mode="collapsed",
        include_tpdu=True,
    ),
    PacketScenarioPreset(
        preset_id="xnu_forwarding",
        label="Xn-U (Inter-gNB / Data Forwarding)",
        interface="Xn-U",
        description="Xn-U user plane packet carrying Long PDCP PDU Number (0x03) and NR RAN Container (0x84, TS 38.425).",
        default_flags={"gtpu_seq": False, "ext_pdcp_sn": True, "ext_nr_ran": True, "ext_pdu_session": False},
        transport_mode="collapsed",
        include_tpdu=True,
    ),
    PacketScenarioPreset(
        preset_id="f1u_cu_du",
        label="F1-U (gNB-CU-UP ↔ gNB-DU)",
        interface="F1-U",
        description="Split gNB F1-U packet carrying NR RAN Container (0x84, TS 38.425) and optional PDU Set Information (0x04).",
        default_flags={"gtpu_seq": False, "ext_nr_ran": True, "ext_pdu_session": False},
        transport_mode="collapsed",
        include_tpdu=True,
    ),
    PacketScenarioPreset(
        preset_id="dummy_qmp",
        label="QoS Monitoring Dummy G-PDU (No T-PDU)",
        interface="N3, N9",
        description="Control G-PDU without T-PDU payload carrying QoS Monitoring timestamps (QMP=1) in PDU Session Container (0x85).",
        default_flags={"gtpu_seq": False, "ext_pdu_session": True, "qmp": True},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
    PacketScenarioPreset(
        preset_id="end_marker",
        label="End Marker (MsgType 0xFE / 254)",
        interface="N3, N9, Xn-U",
        description="GTP-U End Marker signalling packet (Message Type 254) with optional PDU Session Container (0x85) and no T-PDU.",
        default_flags={"gtpu_seq": False, "ext_pdu_session": True},
        transport_mode="collapsed",
        include_tpdu=False,
    ),
]

UP_PRESET_MAP: Dict[str, PacketScenarioPreset] = {p.preset_id: p for p in UP_SCENARIO_PRESETS}


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


# -----------------------------------------------------------------------------
# Target & Field Resolution from Selection in NASTab
# -----------------------------------------------------------------------------

_TS38415_CLAUSE_TO_TARGET: Dict[str, Tuple[str, str, Dict[str, bool]]] = {
    "5.5.2.1": ("dl_pdu_session", "n3_n9", {}),
    "5.5.2.2": ("ul_pdu_session", "n3_n9", {}),
    "6.5.2.1": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True}),
    # DL PDU SESSION INFORMATION IEs (5.5.3.x)
    "5.5.3.4": ("dl_pdu_session", "n3_n9", {"rqi": True}),
    "5.5.3.6": ("dl_pdu_session", "n3_n9", {"ppp": True}),
    "5.5.3.7": ("dl_pdu_session", "n3_n9", {"ppp": True}),
    "5.5.3.8": ("dl_pdu_session", "n3_n9", {"qmp": True}),
    "5.5.3.9": ("dl_pdu_session", "n3_n9", {"qmp": True}),
    "5.5.3.17": ("dl_pdu_session", "n3_n9", {"snp": True}),
    "5.5.3.18": ("dl_pdu_session", "n3_n9", {"snp": True}),
    "5.5.3.23": ("dl_pdu_session", "n3_n9", {"msnp": True}),
    "5.5.3.24": ("dl_pdu_session", "n3_n9", {"msnp": True}),
    "5.5.3.29": ("dl_pdu_session", "n3_n9", {"bssi": True}),
    "5.5.3.30": ("dl_pdu_session", "n3_n9", {"bssi": True}),
    "5.5.3.31": ("dl_pdu_session", "n3_n9", {"ttnbi": True}),
    "5.5.3.32": ("dl_pdu_session", "n3_n9", {"ttnbi": True}),
    # UL PDU SESSION INFORMATION IEs (5.5.3.x)
    "5.5.3.10": ("ul_pdu_session", "n3_n9", {"qmp": True}),
    "5.5.3.11": ("ul_pdu_session", "n3_n9", {"qmp": True}),
    "5.5.3.12": ("ul_pdu_session", "n3_n9", {"qmp": True}),
    "5.5.3.13": ("ul_pdu_session", "n3_n9", {"dl_delay_ind": True}),
    "5.5.3.14": ("ul_pdu_session", "n3_n9", {"dl_delay_ind": True}),
    "5.5.3.15": ("ul_pdu_session", "n3_n9", {"ul_delay_ind": True}),
    "5.5.3.16": ("ul_pdu_session", "n3_n9", {"ul_delay_ind": True}),
    "5.5.3.19": ("ul_pdu_session", "n3_n9", {"snp": True}),
    "5.5.3.20": ("ul_pdu_session", "n3_n9", {"n3_n9_delay_ind": True}),
    "5.5.3.21": ("ul_pdu_session", "n3_n9", {"n3_n9_delay_ind": True}),
    "5.5.3.22": ("ul_pdu_session", "n3_n9", {"new_ie_flag": True, "new_ie_0": True}),
    "5.5.3.25": ("ul_pdu_session", "n3_n9", {"new_ie_flag": True, "new_ie_2": True}),
    "5.5.3.26": ("ul_pdu_session", "n3_n9", {"new_ie_flag": True, "new_ie_1": True}),
    "5.5.3.27": ("ul_pdu_session", "n3_n9", {"new_ie_flag": True, "new_ie_3": True}),
    "5.5.3.28": ("ul_pdu_session", "n3_n9", {"new_ie_flag": True, "new_ie_4": True}),
    # DL PDU SET INFORMATION IEs (6.5.3.x)
    "6.5.3.5": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True, "pssi": True}),
    "6.5.3.7": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True, "epdu": True}),
    "6.5.3.8": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True, "edb": True}),
    "6.5.3.9": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True}),
    "6.5.3.10": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True}),
    "6.5.3.11": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True}),
    "6.5.3.12": ("dl_pdu_set", "n3_n9_pdu_set", {"ext_pdu_set": True, "pssi": True}),
}


def resolve_up_packet_target(
    spec_number: Optional[str],
    clause: str = "",
    ie_name: str = "",
    message_name: str = "",
) -> Optional[Dict[str, Any]]:
    """
    Determine whether the current Inspector selection belongs to a supported
    5G User Plane stack (`TS 29.281` / `TS 38.415`) and return initial target settings.
    Returns `None` for unrelated specifications (NAS, RRC, NGAP, PFCP, etc.).
    """
    clean_spec = str(spec_number or "").replace("TS", "").strip()
    if clean_spec not in ("29.281", "38.415"):
        return None

    clean_clause = re.sub(r"\s*\(.*?\)\s*", "", str(clause or "")).strip()
    norm_ie = _norm(ie_name)
    norm_msg = _norm(message_name)

    if clean_spec == "38.415":
        if clean_clause in _TS38415_CLAUSE_TO_TARGET:
            tgt, preset, flags = _TS38415_CLAUSE_TO_TARGET[clean_clause]
            highlight = "" if clean_clause in ("5.5.2.1", "5.5.2.2", "6.5.2.1") else (ie_name or "")
            return {
                "target_frame": tgt,
                "preset_id": preset,
                "auto_enable_flags": dict(flags),
                "highlight_field": highlight,
            }

        # Fallback by message/IE name
        if "pduset" in norm_msg or "pduset" in norm_ie or clean_clause.startswith("6.5"):
            flags = {"ext_pdu_set": True}
            if "pssize" in norm_ie or "pssi" in norm_ie:
                flags["pssi"] = True
            return {
                "target_frame": "dl_pdu_set",
                "preset_id": "n3_n9_pdu_set",
                "auto_enable_flags": flags,
                "highlight_field": ie_name if "pdusetinformation" not in norm_ie else "",
            }
        if "ulpdu" in norm_msg or "ulpdu" in norm_ie or any(
            k in norm_ie for k in ("uldelay", "dldelay", "n3n9", "newieflag", "congestion", "availablebitrate", "ulqfiseq", "dlreceived", "ulsending", "dlsendingtimestamprepeated")
        ):
            flags = {}
            if "newieflag" in norm_ie or "congestion" in norm_ie or "bitrate" in norm_ie or "pdcpdelay" in norm_ie:
                flags["new_ie_flag"] = True
            if "ulcongestion" in norm_ie:
                flags["new_ie_1"] = True
            elif "dlcongestion" in norm_ie:
                flags["new_ie_2"] = True
            elif "ulavailablebitrate" in norm_ie:
                flags["new_ie_3"] = True
            elif "dlavailablebitrate" in norm_ie:
                flags["new_ie_4"] = True
            elif "d1ulpdcp" in norm_ie:
                flags["new_ie_0"] = True
            return {
                "target_frame": "ul_pdu_session",
                "preset_id": "n3_n9",
                "auto_enable_flags": flags,
                "highlight_field": ie_name if "ulpdusessioninformation" not in norm_ie else "",
            }
        return {
            "target_frame": "dl_pdu_session",
            "preset_id": "n3_n9",
            "auto_enable_flags": {},
            "highlight_field": ie_name if "dlpdusessioninformation" not in norm_ie else "",
        }

    # TS 29.281 (GTPv1-U)
    if "endmarker" in norm_ie or "endmarker" in norm_msg or clean_clause.startswith("7.3.2"):
        return {
            "target_frame": "end_marker",
            "preset_id": "end_marker",
            "auto_enable_flags": {"ext_pdu_session": True},
            "highlight_field": "",
        }
    if "pdusetinformation" in norm_ie or clean_clause == "5.2.2.8":
        return {
            "target_frame": "dl_pdu_set",
            "preset_id": "n3_n9_pdu_set",
            "auto_enable_flags": {"ext_pdu_set": True},
            "highlight_field": "",
        }
    if "pdcppdu" in norm_ie or clean_clause in ("5.2.2.2", "5.2.2.2A", "5.2.2.2a"):
        return {
            "target_frame": "dl_pdu_session",
            "preset_id": "xnu_forwarding",
            "auto_enable_flags": {"ext_pdcp_sn": True, "ext_nr_ran": True},
            "highlight_field": "",
        }
    if "nrrancontainer" in norm_ie or clean_clause == "5.2.2.6":
        return {
            "target_frame": "dl_pdu_session",
            "preset_id": "f1u_cu_du",
            "auto_enable_flags": {"ext_nr_ran": True},
            "highlight_field": "",
        }
    if "ulpdusession" in norm_ie:
        return {
            "target_frame": "ul_pdu_session",
            "preset_id": "n3_n9",
            "auto_enable_flags": {"ext_pdu_session": True},
            "highlight_field": "",
        }
    if "gpdu" in norm_ie or "gpdu" in norm_msg or "pdusessioncontainer" in norm_ie or "gtpu" in norm_ie or clean_clause.startswith("5."):
        return {
            "target_frame": "dl_pdu_session",
            "preset_id": "n3_n9",
            "auto_enable_flags": {"ext_pdu_session": True},
            "highlight_field": "",
        }
    return None


# -----------------------------------------------------------------------------
# Release-Aware DB Capability Inspection (3GPP Protocol DB)
# -----------------------------------------------------------------------------

def load_up_release_capabilities(
    db: Any = None,
    version_38415: str = "",
    version_29281: str = "",
) -> Dict[str, Any]:
    """
    Inspect `3gpp_protocol_data.db` for TS 38.415 / TS 29.281 version capabilities.
    Falls back cleanly to Rel-19 defaults when DB is absent or a spec is unimported.
    """
    caps: Dict[str, Any] = {
        "version_38415": version_38415,
        "version_29281": version_29281,
        "field_names_by_clause": {},
        "ie_descriptions": {},
        "has_db_38415": False,
        "has_db_29281": False,
    }
    if db is None or not hasattr(db, "_get_connection"):
        return caps

    try:
        with db._get_connection() as conn:
            cur = conn.cursor()
            if not caps["version_38415"]:
                row = cur.execute(
                    "SELECT id, version FROM spec_versions WHERE spec_number = '38.415' ORDER BY id DESC LIMIT 1"
                ).fetchone()
            else:
                row = cur.execute(
                    "SELECT id, version FROM spec_versions WHERE spec_number = '38.415' AND version = ?",
                    (caps["version_38415"],),
                ).fetchone()
            if row:
                ver_id_38415 = row["id"]
                caps["version_38415"] = row["version"]
                caps["has_db_38415"] = True
                for m_row in cur.execute(
                    "SELECT id, clause FROM nas_messages WHERE version_id = ?", (ver_id_38415,)
                ).fetchall():
                    cl = str(m_row["clause"] or "").strip()
                    names: Set[str] = set()
                    for ie_row in cur.execute(
                        "SELECT ie_name, field_path FROM message_ies WHERE message_id = ?", (m_row["id"],)
                    ).fetchall():
                        names.add(_norm(ie_row["ie_name"] or ie_row["field_path"]))
                    caps["field_names_by_clause"][cl] = names

                for d_row in cur.execute(
                    "SELECT clause, ie_name, raw_description FROM ie_definitions WHERE version_id = ?",
                    (ver_id_38415,),
                ).fetchall():
                    plain = re.sub(r"<[^>]+>", " ", d_row["raw_description"] or "")
                    plain = " ".join(plain.split())
                    if plain:
                        caps["ie_descriptions"][_norm(d_row["ie_name"])] = plain[:280]

            if not caps["version_29281"]:
                r29 = cur.execute(
                    "SELECT version FROM spec_versions WHERE spec_number = '29.281' ORDER BY id DESC LIMIT 1"
                ).fetchone()
            else:
                r29 = cur.execute(
                    "SELECT version FROM spec_versions WHERE spec_number = '29.281' AND version = ?",
                    (caps["version_29281"],),
                ).fetchone()
            if r29:
                caps["version_29281"] = r29["version"]
                caps["has_db_29281"] = True
    except Exception:
        pass
    return caps


def _has_field(caps: Dict[str, Any], clause: str, token: str) -> bool:
    field_sets = caps.get("field_names_by_clause") or {}
    if clause not in field_sets or not field_sets[clause]:
        return True
    norm_tok = _norm(token)
    return any(norm_tok in name for name in field_sets[clause])


def get_flags_for_target(
    target_frame: str,
    preset_id: str,
    caps: Dict[str, Any],
) -> List[FlagOption]:
    """Return interactive FlagOption pills applicable to the active target frame & release."""
    flags: List[FlagOption] = [
        FlagOption("gtpu_seq", "GTP-U S (Seq)", "Include GTP-U Sequence Number (Octets 9–10, S=1).", "gtpu", False),
    ]
    if target_frame != "end_marker":
        flags.extend([
            FlagOption("ext_pdcp_sn", "0x03 PDCP SN", "Chain Long PDCP PDU Number extension header (0x03).", "ext", preset_id == "xnu_forwarding"),
            FlagOption("ext_nr_ran", "0x84 NR RAN", "Chain NR RAN Container extension header (0x84, TS 38.425).", "ext", preset_id in ("xnu_forwarding", "f1u_cu_du")),
        ])

    if target_frame in ("dl_pdu_session", "ul_pdu_session", "dl_pdu_set", "end_marker"):
        flags.append(
            FlagOption(
                "ext_pdu_session",
                "0x85 PDU Session",
                "Include PDU Session Container extension header (0x85).",
                "ext",
                preset_id in ("n3_n9", "n3_n9_pdu_set", "dummy_qmp", "end_marker"),
            )
        )
    if target_frame in ("dl_pdu_session", "dl_pdu_set", "f1u_cu_du") or preset_id == "n3_n9_pdu_set":
        flags.append(
            FlagOption("ext_pdu_set", "0x04 PDU Set", "Chain PDU Set Information Container (0x04, TS 38.415 Cl. 6.5.2.1).", "ext", target_frame == "dl_pdu_set")
        )

    if target_frame == "dl_pdu_session":
        cl = "5.5.2.1"
        flags.extend([
            FlagOption("qmp", "QMP", "QoS Monitoring Packet: adds 8-octet DL Sending Time Stamp (Cl. 5.5.3.9).", "frame"),
            FlagOption("snp", "SNP", "Sequence Number Presence: adds 3-octet DL QFI Sequence Number (Cl. 5.5.3.18).", "frame"),
            FlagOption("msnp", "MSNP", "MBS Sequence Number Presence: adds 4-octet DL MBS QFI Sequence Number (Cl. 5.5.3.24).", "frame", is_release_present=_has_field(caps, cl, "msnp")),
            FlagOption("ppp", "PPP", "Paging Policy Presence: activates Octet 3 and 3-bit PPI (Cl. 5.5.3.7).", "frame"),
            FlagOption("rqi", "RQI", "Reflective QoS Indicator bit in Octet 2 (Cl. 5.5.3.4).", "frame"),
            FlagOption("bssi", "BSSI", "Burst Size Indicator (Rel-19): activates Octet 3 + 3-octet BSSize (Cl. 5.5.3.30).", "frame", is_release_present=_has_field(caps, cl, "bssi")),
            FlagOption("ttnbi", "TTNBI", "TTNB Indicator (Rel-19): activates Octet 3 + 2-octet TTNB (Cl. 5.5.3.32).", "frame", is_release_present=_has_field(caps, cl, "ttnbi")),
        ])
    elif target_frame == "ul_pdu_session":
        cl = "5.5.2.2"
        flags.extend([
            FlagOption("qmp", "QMP", "QoS Monitoring Packet: adds 3×8-octet timestamps (Cl. 5.5.3.10–12).", "frame"),
            FlagOption("dl_delay_ind", "DL Delay", "Adds 4-octet DL Delay Result (Cl. 5.5.3.14).", "frame"),
            FlagOption("ul_delay_ind", "UL Delay", "Adds 4-octet UL Delay Result (Cl. 5.5.3.16).", "frame"),
            FlagOption("snp", "SNP", "Adds 3-octet UL QFI Sequence Number (Cl. 5.5.3.19).", "frame"),
            FlagOption("n3_n9_delay_ind", "N3/N9 Delay", "Adds 4-octet N3/N9 Delay Result (Cl. 5.5.3.21).", "frame"),
            FlagOption("new_ie_flag", "New IE Flag", "Includes New IE Flags Octet (Rel-16+ extensibility).", "frame", is_release_present=_has_field(caps, cl, "newieflag")),
            FlagOption("new_ie_0", "IE0: D1 PDCP", "New IE Flag 0: adds D1 UL PDCP Delay Result Ind octet (Cl. 5.5.3.22).", "frame", requires_flag="new_ie_flag", is_release_present=_has_field(caps, cl, "pdcpdelay")),
            FlagOption("new_ie_1", "IE1: UL Cong", "New IE Flag 1: adds 2-octet UL Congestion Information (Cl. 5.5.3.26).", "frame", requires_flag="new_ie_flag", is_release_present=_has_field(caps, cl, "ulcongestion")),
            FlagOption("new_ie_2", "IE2: DL Cong", "New IE Flag 2: adds 2-octet DL Congestion Information (Cl. 5.5.3.25).", "frame", requires_flag="new_ie_flag", is_release_present=_has_field(caps, cl, "dlcongestion")),
            FlagOption("new_ie_3", "IE3: UL Bitrate", "New IE Flag 3: adds 4-octet UL Available Bitrate (Cl. 5.5.3.27).", "frame", requires_flag="new_ie_flag", is_release_present=_has_field(caps, cl, "ulavailablebitrate")),
            FlagOption("new_ie_4", "IE4: DL Bitrate", "New IE Flag 4: adds 4-octet DL Available Bitrate (Cl. 5.5.3.28).", "frame", requires_flag="new_ie_flag", is_release_present=_has_field(caps, cl, "dlavailablebitrate")),
        ])
    elif target_frame == "dl_pdu_set":
        flags.extend([
            FlagOption("edb", "EDB", "End of Data Burst indicator bit (Cl. 6.5.3.8).", "frame"),
            FlagOption("epdu", "EPDU", "End PDU of the PDU Set indicator bit (Cl. 6.5.3.7).", "frame"),
            FlagOption("pssi", "PSSI", "PDU Set Size Indicator: adds 3-octet PSSize field (Cl. 6.5.3.12).", "frame"),
        ])

    return flags


# -----------------------------------------------------------------------------
# TS 38.415 & Companion Extension Header Field Builders
# -----------------------------------------------------------------------------

def _desc(caps: Dict[str, Any], name: str, fallback: str) -> str:
    return (caps.get("ie_descriptions") or {}).get(_norm(name), fallback)


def build_dl_pdu_session_fields(flags: Dict[str, bool], caps: Dict[str, Any]) -> List[PacketField]:
    """Build exact octet-aligned fields for TS 38.415 Clause 5.5.2.1 (DL PDU SESSION INFORMATION)."""
    qmp = bool(flags.get("qmp"))
    snp = bool(flags.get("snp"))
    msnp = bool(flags.get("msnp"))
    ppp = bool(flags.get("ppp"))
    rqi = bool(flags.get("rqi"))
    bssi = bool(flags.get("bssi"))
    ttnbi = bool(flags.get("ttnbi"))

    fields: List[PacketField] = [
        PacketField("PDU Type", "PDU Type=0", 4, "M", "0", "38.415", "6.5.3.1", description="PDU Type = 0 (DL PDU SESSION INFORMATION)."),
        PacketField("QMP", "QMP", 1, "M", "1" if qmp else "0", "38.415", "5.5.3.8", flag_key="qmp", description=_desc(caps, "qmp", "Indicates whether DL Sending Time Stamp is present.")),
        PacketField("SNP", "SNP", 1, "M", "1" if snp else "0", "38.415", "5.5.3.17", flag_key="snp", description=_desc(caps, "snp", "Indicates whether DL QFI Sequence Number is present.")),
        PacketField("MSNP", "MSNP", 1, "M", "1" if msnp else "0", "38.415", "5.5.3.23", flag_key="msnp", description=_desc(caps, "msnp", "Indicates whether DL MBS QFI Sequence Number is present.")),
        PacketField("Spare (Octet 1)", "Spare", 1, "M", "0", "38.415", "6.5.3.2", is_spare=True, description="Spare bit, set to 0."),
        PacketField("PPP", "PPP", 1, "M", "1" if ppp else "0", "38.415", "5.5.3.6", flag_key="ppp", description=_desc(caps, "ppp", "Paging Policy Presence indicator.")),
        PacketField("RQI", "RQI", 1, "M", "1" if rqi else "0", "38.415", "5.5.3.4", flag_key="rqi", description=_desc(caps, "rqi", "Reflective QoS Indicator.")),
        PacketField("QoS Flow Identifier (QFI)", "QFI", 6, "M", "0..63", "38.415", "6.5.3.3", description=_desc(caps, "qosflowidentifier", "6-bit QoS Flow Identifier.")),
    ]

    if ppp or bssi or ttnbi:
        fields.extend([
            PacketField("PPI", "PPI", 3, "C", "0..7" if ppp else "0", "38.415", "5.5.3.7", condition="PPP=1", flag_key="ppp", description="Paging Policy Indicator (evaluated when PPP=1)."),
            PacketField("Spare (Octet 3)", "Spare", 3, "C", "0", "38.415", "6.5.3.2", is_spare=True, description="Spare bits in Octet 3."),
            PacketField("BSSI", "BSSI", 1, "C", "1" if bssi else "0", "38.415", "5.5.3.29", flag_key="bssi", description=_desc(caps, "bssi", "Burst Size Indicator (1 = BSSize present).")),
            PacketField("TTNBI", "TTNBI", 1, "C", "1" if ttnbi else "0", "38.415", "5.5.3.31", flag_key="ttnbi", description=_desc(caps, "ttnbi", "Time To Next Burst Indicator (1 = TTNB present).")),
        ])

    if qmp:
        fields.append(PacketField("DL Sending Time Stamp", "DL Sending Time Stamp (64-bit NTP)", 64, "C", "64-bit", "38.415", "5.5.3.9", condition="QMP=1", flag_key="qmp", description=_desc(caps, "dlsendingtimestamp", "64-bit timestamp when UPF sent the DL packet.")))
    if snp:
        fields.append(PacketField("DL QFI Sequence Number", "DL QFI Sequence Number (24-bit)", 24, "C", "24-bit", "38.415", "5.5.3.18", condition="SNP=1", flag_key="snp", description=_desc(caps, "dlqfisequencenumber", "24-bit per-QFI sequence number.")))
    if msnp:
        fields.append(PacketField("DL MBS QFI Sequence Number", "DL MBS QFI Seq Num (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.24", condition="MSNP=1", flag_key="msnp", description=_desc(caps, "dlmbsqfisequencenumber", "32-bit MBS QFI sequence number.")))
    if bssi:
        fields.append(PacketField("BSSize", "Burst Size (BSSize, 24-bit)", 24, "C", "24-bit", "38.415", "5.5.3.30", condition="BSSI=1", flag_key="bssi", description=_desc(caps, "bssize", "Burst Size in octets.")))
    if ttnbi:
        fields.append(PacketField("TTNB", "Time To Next Burst (TTNB, 16-bit)", 16, "C", "16-bit", "38.415", "5.5.3.32", condition="TTNBI=1", flag_key="ttnbi", description=_desc(caps, "ttnb", "Time To Next Burst in ms.")))

    return fields


def build_ul_pdu_session_fields(flags: Dict[str, bool], caps: Dict[str, Any]) -> List[PacketField]:
    """Build exact octet-aligned fields for TS 38.415 Clause 5.5.2.2 (UL PDU SESSION INFORMATION)."""
    qmp = bool(flags.get("qmp"))
    dl_d = bool(flags.get("dl_delay_ind"))
    ul_d = bool(flags.get("ul_delay_ind"))
    snp = bool(flags.get("snp"))
    n3_d = bool(flags.get("n3_n9_delay_ind"))
    new_ie = bool(flags.get("new_ie_flag")) or any(bool(flags.get(f"new_ie_{i}")) for i in range(5))

    fields: List[PacketField] = [
        PacketField("PDU Type", "PDU Type=1", 4, "M", "1", "38.415", "6.5.3.1", description="PDU Type = 1 (UL PDU SESSION INFORMATION)."),
        PacketField("QMP", "QMP", 1, "M", "1" if qmp else "0", "38.415", "5.5.3.8", flag_key="qmp", description="QoS Monitoring Packet indicator."),
        PacketField("DL Delay Ind.", "DL D", 1, "M", "1" if dl_d else "0", "38.415", "5.5.3.13", flag_key="dl_delay_ind", description="Indicates DL Delay Result presence."),
        PacketField("UL Delay Ind.", "UL D", 1, "M", "1" if ul_d else "0", "38.415", "5.5.3.15", flag_key="ul_delay_ind", description="Indicates UL Delay Result presence."),
        PacketField("SNP", "SNP", 1, "M", "1" if snp else "0", "38.415", "5.5.3.17", flag_key="snp", description="Indicates UL QFI Sequence Number presence."),
        PacketField("N3/N9 Delay Ind.", "N3/9", 1, "M", "1" if n3_d else "0", "38.415", "5.5.3.20", flag_key="n3_n9_delay_ind", description="Indicates N3/N9 Delay Result presence."),
        PacketField("New IE Flag", "NewIE", 1, "M", "1" if new_ie else "0", "38.415", "5.5.2.2", flag_key="new_ie_flag", description="Indicates presence of New IE Flags Octet."),
        PacketField("QoS Flow Identifier (QFI)", "QFI", 6, "M", "0..63", "38.415", "6.5.3.3", description="6-bit QoS Flow Identifier."),
    ]

    if qmp:
        fields.extend([
            PacketField("DL Sending Time Stamp Repeated", "DL Sending Time Stamp Repeated (64-bit)", 64, "C", "64-bit", "38.415", "5.5.3.10", condition="QMP=1", flag_key="qmp"),
            PacketField("DL Received Time Stamp", "DL Received Time Stamp (64-bit)", 64, "C", "64-bit", "38.415", "5.5.3.11", condition="QMP=1", flag_key="qmp"),
            PacketField("UL Sending Time Stamp", "UL Sending Time Stamp (64-bit)", 64, "C", "64-bit", "38.415", "5.5.3.12", condition="QMP=1", flag_key="qmp"),
        ])
    if dl_d:
        fields.append(PacketField("DL Delay Result", "DL Delay Result (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.14", condition="DL Delay Ind.=1", flag_key="dl_delay_ind"))
    if ul_d:
        fields.append(PacketField("UL Delay Result", "UL Delay Result (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.16", condition="UL Delay Ind.=1", flag_key="ul_delay_ind"))
    if snp:
        fields.append(PacketField("UL QFI Sequence Number", "UL QFI Sequence Number (24-bit)", 24, "C", "24-bit", "38.415", "5.5.3.19", condition="SNP=1", flag_key="snp"))
    if n3_d:
        fields.append(PacketField("N3/N9 Delay Result", "N3/N9 Delay Result (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.21", condition="N3/N9 Delay Ind.=1", flag_key="n3_n9_delay_ind"))

    if new_ie:
        fields.extend([
            PacketField("New IE flag 7(E)", "E", 1, "C", "0", "38.415", "5.5.2.2", condition="New IE Flag=1"),
            PacketField("New IE Flag 6", "F6", 1, "C", "0", "38.415", "5.5.2.2", is_spare=True),
            PacketField("New IE Flag 5", "F5", 1, "C", "0", "38.415", "5.5.2.2", is_spare=True),
            PacketField("New IE Flag 4", "F4", 1, "C", "1" if flags.get("new_ie_4") else "0", "38.415", "5.5.2.2", flag_key="new_ie_4"),
            PacketField("New IE Flag 3", "F3", 1, "C", "1" if flags.get("new_ie_3") else "0", "38.415", "5.5.2.2", flag_key="new_ie_3"),
            PacketField("New IE Flag 2", "F2", 1, "C", "1" if flags.get("new_ie_2") else "0", "38.415", "5.5.2.2", flag_key="new_ie_2"),
            PacketField("New IE Flag 1", "F1", 1, "C", "1" if flags.get("new_ie_1") else "0", "38.415", "5.5.2.2", flag_key="new_ie_1"),
            PacketField("New IE Flag 0", "F0", 1, "C", "1" if flags.get("new_ie_0") else "0", "38.415", "5.5.2.2", flag_key="new_ie_0"),
        ])
        if flags.get("new_ie_0"):
            fields.extend([
                PacketField("Spare (D1 PDCP Octet)", "Spare", 7, "C", "0", "38.415", "6.5.3.2", is_spare=True),
                PacketField("D1 UL PDCP Delay Result Ind", "D1", 1, "C", "1", "38.415", "5.5.3.22", condition="New IE Flag 0=1", flag_key="new_ie_0"),
            ])
        if flags.get("new_ie_1"):
            fields.append(PacketField("UL Congestion Information", "UL Congestion Info (16-bit)", 16, "C", "16-bit", "38.415", "5.5.3.26", condition="New IE Flag 1=1", flag_key="new_ie_1"))
        if flags.get("new_ie_2"):
            fields.append(PacketField("DL Congestion Information", "DL Congestion Info (16-bit)", 16, "C", "16-bit", "38.415", "5.5.3.25", condition="New IE Flag 2=1", flag_key="new_ie_2"))
        if flags.get("new_ie_3"):
            fields.append(PacketField("UL Available Bitrate", "UL Available Bitrate (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.27", condition="New IE Flag 3=1", flag_key="new_ie_3"))
        if flags.get("new_ie_4"):
            fields.append(PacketField("DL Available Bitrate", "DL Available Bitrate (32-bit)", 32, "C", "32-bit", "38.415", "5.5.3.28", condition="New IE Flag 4=1", flag_key="new_ie_4"))

    return fields


def build_dl_pdu_set_fields(flags: Dict[str, bool], caps: Dict[str, Any]) -> List[PacketField]:
    """Build exact octet-aligned fields for TS 38.415 Clause 6.5.2.1 (DL PDU SET INFORMATION)."""
    edb = bool(flags.get("edb"))
    epdu = bool(flags.get("epdu"))
    pssi = bool(flags.get("pssi"))

    fields: List[PacketField] = [
        PacketField("PDU Type", "PDU Type=0", 4, "M", "0", "38.415", "6.5.3.1", description="PDU Type = 0 (DL PDU SET INFORMATION)."),
        PacketField("EDB", "EDB", 1, "M", "1" if edb else "0", "38.415", "6.5.3.8", flag_key="edb", description=_desc(caps, "edb", "End of Data Burst indicator.")),
        PacketField("EPDU", "EPDU", 1, "M", "1" if epdu else "0", "38.415", "6.5.3.7", flag_key="epdu", description=_desc(caps, "epdu", "End PDU of the PDU Set indicator.")),
        PacketField("PSSI", "PSSI", 1, "M", "1" if pssi else "0", "38.415", "6.5.3.5", flag_key="pssi", description=_desc(caps, "pssi", "PDU Set Size Indicator.")),
        PacketField("Spare (Octet 1)", "Spare", 1, "M", "0", "38.415", "6.5.3.2", is_spare=True),
        PacketField("QoS Flow Identifier (QFI)", "QFI", 6, "M", "0..63", "38.415", "6.5.3.3"),
        PacketField("PSSN (MSB)", "PSSN (hi)", 2, "M", "10-bit", "38.415", "6.5.3.10", description="PDU Set Sequence Number (2 MSBs in Octet 2)."),
        PacketField("PSSN (LSB)", "PSSN (bits 7..0)", 8, "M", "10-bit", "38.415", "6.5.3.10", description="PDU Set Sequence Number (8 LSBs in Octet 3)."),
        PacketField("Spare (Octet 4)", "Spare", 4, "M", "0", "38.415", "6.5.3.2", is_spare=True),
        PacketField("PSI", "PSI", 4, "M", "0..15", "38.415", "6.5.3.9", description="PDU Set Importance."),
        PacketField("PSN", "PSN (8-bit)", 8, "M", "0..255", "38.415", "6.5.3.11", description="PDU Sequence Number within a PDU Set."),
    ]
    if pssi:
        fields.append(PacketField("PSSize", "PDU Set Size (PSSize, 24-bit)", 24, "C", "24-bit", "38.415", "6.5.3.12", condition="PSSI=1", flag_key="pssi"))
    return fields


def build_pdcp_sn_fields() -> List[PacketField]:
    """Build Long PDCP PDU Number payload fields (TS 29.281 Clause 5.2.2.2A)."""
    return [
        PacketField("Spare (PDCP SN Octet 2)", "Spare", 6, "M", "0", "29.281", "5.2.2.2A", is_spare=True),
        PacketField("Long PDCP PDU Number (18-bit)", "PDCP SN (hi)", 2, "M", "18-bit", "29.281", "5.2.2.2A", description="Top 2 bits of 18-bit PDCP PDU Number."),
        PacketField("Long PDCP PDU Number (LSB)", "PDCP PDU Number (bits 15..0)", 16, "M", "18-bit", "29.281", "5.2.2.2A", description="Lower 16 bits of 18-bit PDCP PDU Number."),
        PacketField("Spare (PDCP SN Octet 5)", "Spare", 8, "M", "0", "29.281", "5.2.2.2A", is_spare=True),
        PacketField("Spare (PDCP SN Octet 6)", "Spare", 8, "M", "0", "29.281", "5.2.2.2A", is_spare=True),
        PacketField("Spare (PDCP SN Octet 7)", "Spare", 8, "M", "0", "29.281", "5.2.2.2A", is_spare=True),
    ]


def build_nr_ran_container_fields() -> List[PacketField]:
    """Build NR RAN Container placeholder (TS 29.281 Cl. 5.2.2.6 -> TS 38.425)."""
    return [
        PacketField(
            "NR User Plane Frame (TS 38.425)",
            "NR UP Frame (DL USER DATA / DL DATA DELIVERY STATUS)",
            0,
            "M",
            "TS 38.425",
            "38.425",
            "5.5",
            is_variable=True,
            variable_display_octets=8,
            description="Encapsulated NR User Plane protocol frame (TS 38.425) — full TS 38.425 frame expansion is planned.",
        )
    ]
