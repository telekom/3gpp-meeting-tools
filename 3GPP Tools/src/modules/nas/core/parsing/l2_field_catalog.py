"""Curated octet-level field lists for well-known RAN Layer 2 PDU formats (NR SDAP / PDCP / RLC / MAC).

3GPP RAN2 specifications draw PDU formats as embedded Visio/EMF figures, which cannot be parsed from
DOCX. `L2PduDocxParser` therefore extracts the prose/tables of Clause 6 from the document itself and
attaches these curated field lists (keyed by spec + format-clause title) so messages carry meaningful
`message_ies` rows. Lengths and presence follow the figures in TS 37.324 Cl. 6.2, TS 38.323 Cl. 6.2,
TS 38.322 Cl. 6.2 and TS 38.321 Cl. 6.1/6.2. This module is pure data (no Qt, no I/O).
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

L2FieldDef = Dict[str, str]


def _f(name: str, length: str, presence: str = "M", fmt: str = "Bit field", desc: str = "") -> L2FieldDef:
    return {"name": name, "length": length, "presence": presence, "format": fmt, "description": desc}


_DATA = _f("Data", "Variable", "M", "Octet string", "Upper-layer SDU / payload carried by this PDU.")
_MAC_I = _f("MAC-I", "4 octets", "C", "Octet string", "Message authentication code for integrity protection (present when integrity protection is configured).")

# Each entry: spec, title regex (searched, case-insensitive), optional exclusion regex, direction, fields.
L2_FORMAT_CATALOG: List[Dict[str, Any]] = [
    # --- TS 37.324 SDAP ---
    {
        "spec": "37.324", "title": r"Data PDU without SDAP header", "direction": "UE ↔ NG-RAN node",
        "fields": [_DATA],
    },
    {
        "spec": "37.324", "title": r"DL Data PDU with SDAP header", "direction": "NG-RAN node → UE",
        "fields": [
            _f("RDI", "1 bit", desc="Reflective QoS flow to DRB mapping Indication."),
            _f("RQI", "1 bit", desc="Reflective QoS Indication."),
            _f("QFI", "6 bits", desc="QoS Flow ID."),
            _DATA,
        ],
    },
    {
        "spec": "37.324", "title": r"UL Data PDU with SDAP header", "direction": "UE → NG-RAN node",
        "fields": [
            _f("D/C", "1 bit", desc="1 = Data PDU, 0 = Control PDU."),
            _f("R", "1 bit", desc="Reserved (0)."),
            _f("QFI", "6 bits", desc="QoS Flow ID."),
            _DATA,
        ],
    },
    {
        "spec": "37.324", "title": r"End-Marker Control PDU", "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("D/C", "1 bit", desc="0 = Control PDU."),
            _f("R", "1 bit", desc="Reserved (0)."),
            _f("QFI", "6 bits", desc="QoS Flow ID of the flow for which the end-marker is sent."),
        ],
    },
    # --- TS 38.323 NR PDCP ---
    {
        "spec": "38.323", "title": r"Data PDU for SRBs", "exclude": r"sidelink", "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("R", "4 bits", desc="Reserved (0)."),
            _f("PDCP SN", "12 bits", fmt="Integer", desc="PDCP sequence number."),
            _DATA,
            {**_MAC_I, "presence": "M", "description": "Message authentication code (always present for SRBs, integrity-protected)."},
        ],
    },
    {
        "spec": "38.323", "title": r"Data PDU for DRBs and MRBs with 12 bits PDCP SN", "exclude": r"sidelink",
        "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("D/C", "1 bit", desc="1 = Data PDU, 0 = Control PDU."),
            _f("R", "3 bits", desc="Reserved (0)."),
            _f("PDCP SN", "12 bits", fmt="Integer", desc="PDCP sequence number."),
            _DATA,
            _MAC_I,
        ],
    },
    {
        "spec": "38.323", "title": r"Data PDU for DRBs and MRBs with 18 bits PDCP SN", "exclude": r"sidelink",
        "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("D/C", "1 bit", desc="1 = Data PDU, 0 = Control PDU."),
            _f("R", "5 bits", desc="Reserved (0)."),
            _f("PDCP SN", "18 bits", fmt="Integer", desc="PDCP sequence number."),
            _DATA,
            _MAC_I,
        ],
    },
    # --- TS 38.322 NR RLC ---
    {
        "spec": "38.322", "title": r"\bTMD PDU\b", "direction": "UE ↔ NG-RAN node",
        "fields": [_DATA],
    },
    {
        "spec": "38.322", "title": r"\bUMD PDU\b", "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("SI", "2 bits", desc="Segmentation Info: 00 complete SDU, 01 first, 10 last, 11 middle segment."),
            _f("R", "0 or 2 bits", presence="M", desc="Reserved (2 bits with 12-bit SN, absent with 6-bit SN)."),
            _f("SN", "6 or 12 bits", fmt="Integer", desc="RLC sequence number (absent when SI = 00)."),
            _f("SO", "16 bits", "C", "Integer", "Segment Offset (present when SI = 10 or 11)."),
            _DATA,
        ],
    },
    {
        "spec": "38.322", "title": r"\bAMD PDU\b", "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("D/C", "1 bit", desc="1 = Data PDU, 0 = Control PDU."),
            _f("P", "1 bit", desc="Polling bit."),
            _f("SI", "2 bits", desc="Segmentation Info: 00 complete SDU, 01 first, 10 last, 11 middle segment."),
            _f("R", "0 or 2 bits", desc="Reserved (absent with 12-bit SN)."),
            _f("SN", "12 or 18 bits", fmt="Integer", desc="RLC sequence number."),
            _f("SO", "16 bits", "C", "Integer", "Segment Offset (present when SI = 10 or 11)."),
            _DATA,
        ],
    },
    # --- TS 38.321 NR MAC ---
    {
        "spec": "38.321", "title": r"subheader", "exclude": r"\bCE\b|Control Element", "direction": "UE ↔ NG-RAN node",
        "fields": [
            _f("R", "1 bit", desc="Reserved (0)."),
            _f("F", "1 bit", desc="Format: 0 = 8-bit L field, 1 = 16-bit L field."),
            _f("LCID", "6 bits", fmt="Integer", desc="Logical Channel ID (SRB / DRB / CCCH) or MAC CE identifier."),
            _f("L", "8 or 16 bits", "C", "Integer", "Length in octets of the MAC SDU / variable-sized MAC CE."),
        ],
    },
]


def find_l2_format_fields(spec_number: str, title: str) -> Optional[Dict[str, Any]]:
    """Return the curated catalog entry (spec + title regex) for a Clause 6 format section, or None."""
    clean_title = " ".join((title or "").split())
    for entry in L2_FORMAT_CATALOG:
        if entry["spec"] != spec_number:
            continue
        if not re.search(entry["title"], clean_title, re.IGNORECASE):
            continue
        excl = entry.get("exclude")
        if excl and re.search(excl, clean_title, re.IGNORECASE):
            continue
        return entry
    return None
