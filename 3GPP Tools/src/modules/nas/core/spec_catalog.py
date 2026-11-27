"""Curated 3GPP protocol specification catalogue with faceted metadata (no Qt dependency).

Facets (`generation`, `domain`, `plane`, `layer`) drive the Specification Import Studio filters and
bundles. Specs flagged `importable=False` are reference-only (shown greyed out in the import dialog).
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple


def _spec(
    acronym: str,
    wg: str,
    title: str,
    full_title: str,
    tech: str,
    description: str,
    *,
    domain: str,
    plane: str,
    layer: str,
    importable: bool = True,
    note: str = "",
) -> Dict[str, object]:
    return {
        "acronym": acronym,
        "wg": wg,
        "title": title,
        "full_title": full_title,
        "tech": tech,
        "description": description,
        "domain": domain,
        "plane": plane,
        "layer": layer,
        "importable": importable,
        "note": note,
    }


PROTOCOL_SPEC_CATALOG: Dict[str, Dict[str, object]] = {
    "24.501": _spec(
        "5GS NAS", "CT1", "5G Non-Access-Stratum (NAS) Protocol",
        "Non-Access-Stratum (NAS) protocol for 5G System (5GS); Stage 3", "5G",
        "Core Network signaling between UE and AMF/SMF for 5G registration, connection, mobility, and PDU sessions.",
        domain="Core", plane="CP", layer="NAS",
    ),
    "24.526": _spec(
        "5GS UE Policies (URSP)", "CT1", "UE Policies for 5G System (URSP & ANDSP); Stage 3",
        "User Equipment (UE) policies for 5G System (5GS); Stage 3", "5G",
        "Stage 3 bit-level encoding of UE Route Selection Policy (URSP) rules, Traffic Descriptors, Route Selection Descriptors, and ANDSP/WLANSP.",
        domain="Core", plane="CP", layer="UE Policy",
    ),
    "24.588": _spec(
        "5GS V2X Policies (V2XP)", "CT1", "V2X Services UE Policies (V2XP); Stage 3",
        "Vehicle-to-Everything (V2X) services in 5G System (5GS); User Equipment (UE) policies; Stage 3", "5G/V2X",
        "Stage 3 bit-level encoding of V2X Policy (V2XP) for PC5 and Uu communications, QoS mapping, SLRB, and NR-PC5 security.",
        domain="Core", plane="CP", layer="UE Policy",
    ),
    "24.555": _spec(
        "5G ProSe Policies (ProSeP)", "CT1", "5G Proximity-Services UE Policies (ProSeP); Stage 3",
        "Proximity-services (ProSe) in 5G System (5GS); User Equipment (UE) policies; Stage 3", "5G/ProSe",
        "Stage 3 encoding of 5G ProSe policies for direct discovery, direct communication over PC5, UE-to-network/UE-to-UE relay, and multi-hop.",
        domain="Core", plane="CP", layer="UE Policy",
    ),
    "24.514": _spec(
        "5GS Ranging & SL Positioning (RSLPP)", "CT1", "Ranging and Sidelink Positioning Policies (RSLPP); Stage 3",
        "Ranging based services and sidelink positioning in 5G system (5GS); Stage 3", "5G/Ranging",
        "Stage 3 encoding of Ranging & Sidelink Positioning Policies (RSLPP) Clause 12 coding, ProSe/V2X QoS mapping, and SLPKMF discovery.",
        domain="Core", plane="CP", layer="UE Policy",
    ),
    "24.578": _spec(
        "5GS Aerial Access Policies (A2XP)", "CT1", "Aircraft-to-Everything UE Policies (A2XP); Stage 3",
        "Aircraft-to-Everything (A2X) services in 5G System (5GS); UE policies; Stage 3", "5G/A2X",
        "Stage 3 encoding of A2X policies for PC5 communication, Broadcast Remote ID (BRID), Direct Detect and Avoid (DDAA), C2, and Uu.",
        domain="Core", plane="CP", layer="UE Policy",
    ),
    "24.301": _spec(
        "EPS NAS", "CT1", "4G EPS Non-Access-Stratum (NAS) Protocol",
        "Non-Access-Stratum (NAS) protocol for Evolved Packet System (EPS); Stage 3", "LTE",
        "Core Network signaling between UE and MME for 4G LTE attach, tracking area update, and EPS bearer establishment.",
        domain="Core", plane="CP", layer="NAS",
    ),
    "24.008": _spec(
        "Core Network L3", "CT1", "Core Network Protocols; Stage 3",
        "Mobile radio interface Layer 3 specification; Core Network protocols; Stage 3", "3G/2G",
        "Legacy 3G/2G Mobility Management (MM), GPRS Mobility Management (GMM), and Call Control (CC).",
        domain="Core", plane="CP", layer="NAS",
    ),
    "29.244": _spec(
        "PFCP", "CT4", "Packet Forwarding Control Protocol (PFCP)",
        "Interface between the Control Plane and the User Plane nodes (PFCP); Stage 3", "5G/Core",
        "N4/Sxa/Sxb interface protocol between SMF/SGW-C and UPF/SGW-U for packet detection, forwarding, and usage reporting.",
        domain="Core", plane="CP", layer="App Protocol",
    ),
    "29.281": _spec(
        "GTPv1-U", "CT4", "GTP User Plane Tunneling (GTPv1-U)",
        "General Packet Radio System (GPRS) Tunnelling Protocol User Plane (GTPv1-U)", "5G/Core",
        "Encapsulated user-plane IP and Ethernet tunneling over UDP between radio access nodes (gNodeB) and UPF/SGW-U.",
        domain="Core", plane="UP", layer="Tunnel / Framing",
    ),
    "38.331": _spec(
        "NR RRC", "RAN2", "5G NR Radio Resource Control (RRC)",
        "NR; Radio Resource Control (RRC) protocol specification", "5G",
        "Air interface signaling: Broadcast of system information (SIBs), RRC connection establishment, reconfiguration, and handover.",
        domain="RAN", plane="CP", layer="L3 / RRC",
    ),
    "36.331": _spec(
        "LTE RRC", "RAN2", "4G LTE Radio Resource Control (RRC)",
        "E-UTRA; Radio Resource Control (RRC); Protocol specification", "LTE",
        "Air interface signaling for 4G LTE: SIB broadcast, paging, connection control, measurement reporting, and mobility.",
        domain="RAN", plane="CP", layer="L3 / RRC",
    ),
    "38.413": _spec(
        "NGAP", "RAN3", "NG Application Protocol (NGAP)",
        "NG-RAN; NG Application Protocol (NGAP)", "5G",
        "N2 interface control protocol between gNodeB and 5G Core (AMF) for PDU session resource management and UE context transfer.",
        domain="RAN", plane="CP", layer="App Protocol",
    ),
    "38.415": _spec(
        "PDU Session UP", "RAN3", "PDU Session User Plane Protocol",
        "NG-RAN; PDU Session user plane protocol", "5G",
        "Framing protocol for user-plane data over N3/N9, conveying QoS Flow Identifiers (QFI) and Reflective QoS Indicators (RQI).",
        domain="RAN", plane="UP", layer="Tunnel / Framing",
    ),
    "38.423": _spec(
        "XnAP", "RAN3", "Xn Application Protocol (XnAP)",
        "NG-RAN; Xn Application Protocol (XnAP)", "5G",
        "Direct inter-gNodeB interface protocol for dual connectivity, Xn-based handover, and inter-node coordination.",
        domain="RAN", plane="CP", layer="App Protocol",
    ),
    "38.473": _spec(
        "F1AP", "RAN3", "F1 Application Protocol (F1AP)",
        "NG-RAN; F1 Application Protocol (F1AP)", "5G",
        "Interface between Central Unit (CU) and Distributed Unit (DU) in disaggregated gNodeB deployments.",
        domain="RAN", plane="CP", layer="App Protocol",
    ),
    "38.463": _spec(
        "E1AP", "RAN3", "E1 Application Protocol (E1AP)",
        "NG-RAN; E1 Application Protocol (E1AP)", "5G",
        "Interface between CU Control Plane (CU-CP) and CU User Plane (CU-UP) nodes.",
        domain="RAN", plane="CP", layer="App Protocol",
    ),
    "38.412": _spec(
        "NG-C Transport", "RAN3", "NG Signaling Transport (NG-C)",
        "NG-RAN; NG signalling transport", "5G",
        "SCTP transport layer specification for NGAP signaling over the N2 reference point.",
        domain="RAN", plane="CP", layer="Transport",
    ),
    "36.413": _spec(
        "S1AP", "RAN3", "S1 Application Protocol (S1AP)",
        "E-UTRA; S1 Application Protocol (S1AP)", "LTE",
        "S1-MME interface control protocol between eNodeB and 4G Core (MME) for bearer management and handover.",
        domain="RAN", plane="CP", layer="App Protocol",
    ),
    "23.501": _spec(
        "5GS Architecture", "SA2", "System Architecture for 5G System (Stage 2)",
        "System architecture for the 5G System (5GS); Stage 2", "5G",
        "High-level Stage 2 architectural requirements, network functions, reference points, and roaming models.",
        domain="Core", plane="CP+UP", layer="Architecture", importable=False,
    ),
    "23.502": _spec(
        "5GS Procedures", "SA2", "Procedures for 5G System (Stage 2)",
        "Procedures for the 5G System (5GS); Stage 2", "5G",
        "Stage 2 call flows, message sequences, registration procedures, and PDU session establishment workflows.",
        domain="Core", plane="CP", layer="Architecture", importable=False,
    ),
    # --- RAN Layer 2 (Uu air interface) ---
    "38.300": _spec(
        "NR Overall Description", "RAN2", "NR and NG-RAN Overall Description; Stage 2",
        "NR; NR and NG-RAN Overall description; Stage-2", "5G",
        "Defines the Uu CP/UP protocol stacks (Fig. 4.4.1-1 / 4.4.2-1), SRB/DRB bearer model, and logical-to-transport channel mapping.",
        domain="RAN", plane="CP+UP", layer="Architecture", importable=False,
        note="Reference only: bearer/channel mapping rules are built into the Packet Stack composer.",
    ),
    "37.324": _spec(
        "SDAP", "RAN2", "Service Data Adaptation Protocol (SDAP)",
        "E-UTRA and NR; Service Data Adaptation Protocol (SDAP) specification", "5G",
        "UP Layer 2: maps QoS flows (QFI) to DRBs and carries the 1-octet SDAP header (D/C, RDI, RQI, QFI).",
        domain="RAN", plane="UP", layer="L2",
    ),
    "38.323": _spec(
        "NR PDCP", "RAN2", "NR Packet Data Convergence Protocol (PDCP)",
        "NR; Packet Data Convergence Protocol (PDCP) specification", "5G",
        "CP/UP Layer 2: SN numbering, ROHC/EHC/UDC header compression, ciphering, integrity (MAC-I), duplication, reordering.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "38.322": _spec(
        "NR RLC", "RAN2", "NR Radio Link Control (RLC)",
        "NR; Radio Link Control (RLC) protocol specification", "5G",
        "CP/UP Layer 2: TM / UM / AM modes, segmentation (SI + SO) and ARQ status reporting.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "38.321": _spec(
        "NR MAC", "RAN2", "NR Medium Access Control (MAC)",
        "NR; Medium Access Control (MAC) protocol specification", "5G",
        "CP/UP Layer 2: MAC subheaders (R/F/LCID/L), MAC CEs, multiplexing, HARQ, random access and scheduling requests.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "36.323": _spec(
        "LTE PDCP", "RAN2", "LTE Packet Data Convergence Protocol (PDCP)",
        "E-UTRA; Packet Data Convergence Protocol (PDCP) specification", "LTE",
        "4G CP/UP Layer 2: PDCP SN, ROHC, ciphering and integrity for E-UTRA radio bearers.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "36.322": _spec(
        "LTE RLC", "RAN2", "LTE Radio Link Control (RLC)",
        "E-UTRA; Radio Link Control (RLC) protocol specification", "LTE",
        "4G CP/UP Layer 2: TM / UM / AM modes, segmentation, concatenation and ARQ.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "36.321": _spec(
        "LTE MAC", "RAN2", "LTE Medium Access Control (MAC)",
        "E-UTRA; Medium Access Control (MAC) protocol specification", "LTE",
        "4G CP/UP Layer 2: MAC subheaders, MAC CEs, multiplexing, HARQ and scheduling.",
        domain="RAN", plane="CP+UP", layer="L2",
    ),
    "38.425": _spec(
        "NR UP Protocol", "RAN3", "NR User Plane Protocol (Xn-U / F1-U / E1)",
        "NG-RAN; NR user plane protocol", "5G",
        "NR UP frames (DL USER DATA, DL DATA DELIVERY STATUS, ASSISTANCE INFORMATION) carried in GTP-U extension header 0x84.",
        domain="RAN", plane="UP", layer="Tunnel / Framing", importable=False,
        note="Not yet supported by the importer (planned for the RAN user plane packet stack).",
    ),
    "25.323": _spec(
        "UMTS PDCP (legacy)", "RAN2", "UMTS Packet Data Convergence Protocol (PDCP)",
        "Packet Data Convergence Protocol (PDCP) specification", "3G/2G",
        "Legacy 3G UTRAN PDCP. Not used in 4G/5G — see TS 36.323 (LTE) and TS 38.323 (NR).",
        domain="RAN", plane="CP+UP", layer="L2", importable=False,
        note="Legacy 3G: use TS 38.323 for 5G NR and TS 36.323 for LTE.",
    ),
}

# Display order of importable specs (grouped by role).
IMPORT_ORDER: List[str] = [
    "38.331", "36.331",
    "37.324", "38.323", "38.322", "38.321", "36.323", "36.322", "36.321",
    "38.413", "38.423", "38.473", "38.463", "36.413", "38.412",
    "38.415", "29.244", "29.281",
    "24.501", "24.526", "24.588", "24.555", "24.514", "24.578", "24.301", "24.008",
    "38.300", "38.425", "25.323",
]

SPEC_BUNDLES: List[Tuple[str, str, List[str]]] = [
    ("⚡ 5G RAN L2 Stack", "SDAP + PDCP + RLC + MAC (Uu user/control plane)", ["37.324", "38.323", "38.322", "38.321"]),
    ("⚡ 5G RAN L3 & NG", "NR RRC + NGAP + NG-C transport", ["38.331", "38.413", "38.412"]),
    ("⚡ 5G NAS & UE Policies", "5GS NAS + URSP / V2XP / ProSeP / RSLPP / A2XP", ["24.501", "24.526", "24.588", "24.555", "24.514", "24.578"]),
    ("⚡ 5G User Plane", "GTP-U + PDU Session UP", ["29.281", "38.415"]),
]


def get_spec_metadata(spec_number: str) -> Dict[str, str]:
    """Returns rich metadata dictionary for a 3GPP specification number."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    if clean_num in PROTOCOL_SPEC_CATALOG:
        return PROTOCOL_SPEC_CATALOG[clean_num]  # type: ignore[return-value]

    # Heuristic fallback based on 3GPP specification series
    wg = "3GPP"
    tech = "Multi"
    domain = "Core"
    if clean_num.startswith("24."):
        wg, tech = "CT1", "Core"
    elif clean_num.startswith("29."):
        wg, tech = "CT4", "Core"
    elif clean_num.startswith("38."):
        wg = "RAN2" if clean_num.startswith("38.3") else "RAN3"
        tech, domain = "5G", "RAN"
    elif clean_num.startswith("36."):
        wg = "RAN2" if clean_num.startswith("36.3") else "RAN3"
        tech, domain = "LTE", "RAN"
    elif clean_num.startswith("23."):
        wg, tech = "SA2", "Architecture"

    return {
        "acronym": f"TS {clean_num}",
        "wg": wg,
        "title": f"TS {clean_num} Specification",
        "full_title": f"3GPP TS {clean_num} Technical Specification",
        "tech": tech,
        "description": f"Technical Specification TS {clean_num} governed by 3GPP {wg}.",
        "domain": domain,
        "plane": "",
        "layer": "",
        "importable": True,
        "note": "",
    }


def get_spec_title(spec_number: str) -> str:
    """Returns backwards-compatible short title string with acronym."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return f"TS {clean_num} ({meta['acronym']})"


def get_spec_display_title(spec_number: str) -> str:
    """Returns a rich single-line display label: e.g. [CT1] TS 24.501 — 5GS NAS."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return f"[{meta['wg']}] TS {clean_num} — {meta['title']}"


def get_spec_tooltip(spec_number: str) -> str:
    """Returns multi-line informative tooltip describing the specification's 3GPP role."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return (
        f"3GPP TS {clean_num}: {meta['full_title']}\n"
        f"Working Group: {meta['wg']}  •  Domain: {meta['tech']}\n\n"
        f"Role: {meta['description']}"
    )


def spec_generation(spec_number: str) -> str:
    """Returns the generation facet ('5G', 'LTE', '3G/2G', or 'Other') for a spec."""
    tech = str(get_spec_metadata(spec_number).get("tech", ""))
    if tech.startswith("5G"):
        return "5G"
    if tech.startswith("LTE"):
        return "LTE"
    if tech.startswith("3G"):
        return "3G/2G"
    return "Other"


def spec_matches_filters(
    spec_number: str,
    *,
    text: str = "",
    generation: Optional[str] = None,
    domain: Optional[str] = None,
    plane: Optional[str] = None,
    layer: Optional[str] = None,
) -> bool:
    """Facet + free-text filter used by the Specification Import Studio (None = any)."""
    meta = get_spec_metadata(spec_number)
    if generation and spec_generation(spec_number) != generation:
        return False
    if domain and meta.get("domain") != domain:
        return False
    if plane:
        spec_plane = str(meta.get("plane", ""))
        if spec_plane != plane and spec_plane != "CP+UP":
            return False
    if layer and meta.get("layer") != layer:
        return False
    needle = (text or "").strip().lower()
    if needle:
        haystack = " ".join(
            str(meta.get(k, ""))
            for k in ("acronym", "title", "full_title", "wg", "tech", "layer", "domain", "description")
        ).lower() + f" {spec_number} ts {spec_number}"
        return all(tok in haystack for tok in needle.split())
    return True


def importable_spec_order() -> List[str]:
    """Catalogue spec numbers in display order (importable first, reference-only last)."""
    ordered = [s for s in IMPORT_ORDER if s in PROTOCOL_SPEC_CATALOG]
    remaining = [s for s in PROTOCOL_SPEC_CATALOG if s not in ordered and PROTOCOL_SPEC_CATALOG[s].get("importable")]
    return ordered + remaining


def is_importable_spec(spec_number: str) -> bool:
    meta: Optional[Dict[str, object]] = PROTOCOL_SPEC_CATALOG.get(str(spec_number).replace("TS", "").strip())
    return bool(meta and meta.get("importable"))
