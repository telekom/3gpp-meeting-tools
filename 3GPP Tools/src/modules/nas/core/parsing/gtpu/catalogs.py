"""Constants, catalogs, interface maps, and fallback diagrams for TS 29.281 (GTPv1-U) parsing."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Tuple

RE_GTPU_CLAUSE_HEADER = re.compile(
    r"^((?:5\.1|5\.2\.1|5\.2\.2|7\.[23]|8)(?:\.[0-9A-Za-z]+)*)\s*(.*)$"
)
RE_GTPU_ANY_HEADING = re.compile(
    r"^(?:(\d+(?:\.[0-9A-Za-z]+)*)\s+(.+)|(Annex\s+[A-Z]\b.*))$",
    re.IGNORECASE,
)
RE_GTPU_TABLE_CAPTION = re.compile(
    r"^Table\s+([5678]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+)$",
    re.IGNORECASE,
)
RE_GTPU_FIG_511 = re.compile(r"^Figure\s+5\.1-1\b", re.IGNORECASE)
RE_GTPU_FIG_5213 = re.compile(r"^Figure\s+5\.2\.1-3\b", re.IGNORECASE)

# Interface applicability mappings based on TS 29.281 Clause 5.2 and Clause 7
GTPU_EXT_HEADER_INTERFACES: Dict[str, str] = {
    "udpport": "All Interfaces",
    "pdcppdunumber": "S1-U, X2, N3, Xn, Iu",
    "longpdcppdunumber": "S1-U, X2, N3, Xn",
    "serviceclassindicator": "Gn, Gp, S5/S8, S4, S1-U, S12, Iu, S11-U",
    "rancontainer": "X2",
    "xwrancontainer": "Xw",
    "nrrancontainer": "X2-U, Xn-U, F1-U",
    "pdusessioncontainer": "N3, N9, N3mb, N19mb",
    "pdusetinformationcontainer": "N3, N9, Xn-U, F1-U",
}

GTPU_MESSAGE_INTERFACES: Dict[str, str] = {
    "echorequest": "All Interfaces",
    "echoresponse": "All Interfaces",
    "supportedextensionheadersnotification": "All Interfaces",
    "errorindication": "All Interfaces",
    "endmarker": "S1-U, S11-U, S5/S8, N3, N9, X2, Xn",
    "tunnelstatus": "S5/S8, N3, N9",
    "gpdu": "N3, N9, N19, S1-U, S5/S8, F1-U, X2-U, Xn-U, W1-U, Gn, Gp, S4, S2a, S2b, S12, M1, Sn",
    "gtpuheadercommon": "All Interfaces",
}

GTPU_MESSAGE_DEFAULT_CONTEXT: Dict[str, Tuple[str, str]] = {
    "echorequest": (
        "An Echo Request shall be sent on a path to another GTP-U peer to find out if it is alive (path management).",
        "GTP-U Peer → GTP-U Peer",
    ),
    "echoresponse": (
        "The message shall be sent as a response to a received Echo Request.",
        "GTP-U Peer → GTP-U Peer",
    ),
    "supportedextensionheadersnotification": (
        "This message indicates a list of supported Extension Header types that the GTP entity on the identified IP address can support.",
        "GTP-U Peer → GTP-U Peer",
    ),
    "errorindication": (
        "When a GTP-U node receives a G-PDU for which no EPS Bearer, PDP context, RAB, or PDU Session exists, the GTP-U node shall send an Error Indication to the originating node.",
        "GTP-U Peer → GTP-U Peer",
    ),
    "endmarker": (
        "The End Marker message(s) shall be sent after the last G-PDU that needs to be sent on a GTP-U tunnel to indicate the end of the payload stream.",
        "Source GTP-U Peer → Target GTP-U Peer",
    ),
    "tunnelstatus": (
        "The Tunnel Status message is sent between GTP-U peers to transfer tunnel status information.",
        "GTP-U Peer ↔ GTP-U Peer",
    ),
    "gpdu": (
        "GTPv1-U user plane packet comprising the GTP-U Header (8–12 octets), optional chained 4n-octet Extension Headers, and the encapsulated T-PDU payload (optional for dummy/control G-PDUs).",
        "GTP-U Peer ↔ GTP-U Peer",
    ),
    "gtpuheadercommon": (
        "Common GTPv1-U header structure (8 mandatory octets plus 4 optional octets when E, S, or PN flags are set) shared by all GTP-U signalling messages and G-PDUs (Clause 5.1).",
        "GTP-U Peer ↔ GTP-U Peer",
    ),
}

DEFAULT_TYPE_REGISTRY: Dict[str, Dict[str, str]] = {
    "recovery": {"type_id": "14", "format": "TV", "clause": "8.2", "length": "1"},
    "tunnelendpointidentifierdatai": {"type_id": "16", "format": "TV", "clause": "8.3", "length": "4"},
    "gtpupeeraddress": {"type_id": "133", "format": "TLV", "clause": "8.4", "length": "4 or 16"},
    "gsnaddress": {"type_id": "133", "format": "TLV", "clause": "8.4", "length": "4 or 16"},
    "extensionheadertypelist": {"type_id": "141", "format": "TLV", "clause": "8.5", "length": "Variable"},
    "gtputunnelstatusinformation": {"type_id": "230", "format": "TLV", "clause": "8.7", "length": "Variable"},
    "recoverytimestamp": {"type_id": "231", "format": "TLV", "clause": "8.8", "length": "Variable"},
    "privateextension": {"type_id": "255", "format": "TLV", "clause": "8.6", "length": "Variable"},
}

DEFAULT_EXT_HEADER_REGISTRY: Dict[str, Dict[str, Any]] = {
    "udpport": {"iei": "0x40", "iei_list": ["0x40"], "clause": "5.2.2.1", "name": "UDP Port"},
    "pdcppdunumber": {"iei": "0xC0", "iei_list": ["0xC0"], "clause": "5.2.2.2", "name": "PDCP PDU Number"},
    "longpdcppdunumber": {
        "iei": "0x03 / 0x82",
        "iei_list": ["0x03", "0x82"],
        "clause": "5.2.2.2A",
        "name": "Long PDCP PDU Number",
    },
    "serviceclassindicator": {
        "iei": "0x20",
        "iei_list": ["0x20"],
        "clause": "5.2.2.3",
        "name": "Service Class Indicator",
    },
    "rancontainer": {"iei": "0x81", "iei_list": ["0x81"], "clause": "5.2.2.4", "name": "RAN Container"},
    "xwrancontainer": {"iei": "0x83", "iei_list": ["0x83"], "clause": "5.2.2.5", "name": "Xw RAN Container"},
    "nrrancontainer": {"iei": "0x84", "iei_list": ["0x84"], "clause": "5.2.2.6", "name": "NR RAN Container"},
    "pdusessioncontainer": {
        "iei": "0x85",
        "iei_list": ["0x85"],
        "clause": "5.2.2.7",
        "name": "PDU Session Container",
    },
    "pdusetinformationcontainer": {
        "iei": "0x04 / 0x86",
        "iei_list": ["0x04", "0x86"],
        "clause": "5.2.2.8",
        "name": "PDU Set Information Container",
    },
}

GTPU_EXT_HEADER_ORDER: List[Tuple[str, str, str]] = [
    ("udpport", "UDP Port", "Clause 5.2.2.1"),
    ("pdcppdunumber", "PDCP PDU Number", "Clause 5.2.2.2"),
    ("longpdcppdunumber", "Long PDCP PDU Number", "Clause 5.2.2.2A"),
    ("serviceclassindicator", "Service Class Indicator", "Clause 5.2.2.3"),
    ("rancontainer", "RAN Container", "Clause 5.2.2.4"),
    ("xwrancontainer", "Xw RAN Container", "Clause 5.2.2.5"),
    ("nrrancontainer", "NR RAN Container", "Clause 5.2.2.6"),
    ("pdusessioncontainer", "PDU Session Container", "Clause 5.2.2.7"),
    ("pdusetinformationcontainer", "PDU Set Information Container", "Clause 5.2.2.8"),
]

GTPU_EXT_HEADER_DESCRIPTIONS: Dict[str, str] = {
    "udpport": (
        "Used in Error Indication messages (All Interfaces) to convey the UDP Source Port "
        "of the G-PDU that triggered the Error Indication (DoS mitigation; Clause 5.2.2.1)."
    ),
    "pdcppdunumber": (
        "Carries 12-bit (NR, TS 38.323) or 15-bit (E-UTRA, TS 36.323) PDCP PDU sequence numbers "
        "during handover data forwarding (X2/S1/Xn/N3) or UTRAN SRNS relocation (Clause 5.2.2.2)."
    ),
    "longpdcppdunumber": (
        "Carries 18-bit PDCP PDU sequence numbers during direct (X2/Xn) or indirect (S1/N3) DL data "
        "forwarding at handover. Codepoint 0x03 (Rel-15+) / legacy 0x82 (Clause 5.2.2.2A)."
    ),
    "serviceclassindicator": (
        "Identifies the GERAN A/Gb Service Class Indicator (SCI) for downlink T-PDUs over Gn/Gp, "
        "S5/S8, and S4 (ignored by eNB/RNC/MME on S1-U/S12/Iu/S11-U; Clause 5.2.2.3)."
    ),
    "rancontainer": (
        "Carries the E-UTRA user plane protocol container defined in TS 36.425 over X2-U between "
        "eNBs; a G-PDU with this header may be sent without a T-PDU (Clause 5.2.2.4)."
    ),
    "xwrancontainer": (
        "Carries the LTE-WLAN Aggregation user plane container defined in TS 36.465 over Xw-U "
        "between eNB and WLAN Termination; may be sent without a T-PDU (Clause 5.2.2.5)."
    ),
    "nrrancontainer": (
        "Carries the NR User Plane Protocol container defined in TS 38.425 over X2-U, Xn-U, and "
        "F1-U; a G-PDU with this header may be sent without a T-PDU (Clause 5.2.2.6)."
    ),
    "pdusessioncontainer": (
        "Mandatory on N3/N9/N3mb/N19mb G-PDUs and 5GS↔EPS forwarding End Markers; carries PDU "
        "Session User Plane protocol frames (QFI, RQI, PPI, QoS Monitoring, delay, congestion) "
        "defined in TS 38.415 Clause 5; may be sent as a dummy G-PDU without a T-PDU (Clause 5.2.2.7)."
    ),
    "pdusetinformationcontainer": (
        "Carries PDU Set Information User Plane protocol frames (QFI, PSSN, PSI, EDB, EPDU, PS Size) "
        "defined in TS 38.415 Clause 6 over N3, N9, Xn-U, and F1-U. Codepoint 0x04 (Rel-19+) / "
        "legacy 0x86 (Clause 5.2.2.8)."
    ),
}

# Common GTP-U Header fields from TS 29.281 Clause 5.1 (Figure 5.1-1)
GTPU_COMMON_HEADER_FIELDS: List[Dict[str, str]] = [
    {
        "iei": "Octet 1 (b8-6)",
        "name": "Version",
        "presence": "M",
        "format": "Bit field",
        "length": "3 bits",
        "field_description": "GTP-U protocol version; set to '001' (1) in GTPv1-U (Octet 1 bits 8–6).",
    },
    {
        "iei": "Octet 1 (b5)",
        "name": "Protocol Type (PT)",
        "presence": "M",
        "format": "Flag",
        "length": "1 bit",
        "field_description": "Protocol discriminator: '1' for GTP (TS 29.281) and '0' for GTP' (TS 32.295).",
    },
    {
        "iei": "Octet 1 (b4)",
        "name": "Spare (*)",
        "presence": "M",
        "format": "Flag",
        "length": "1 bit",
        "field_description": "Spare bit in Octet 1 bit 4; shall be sent as '0' and ignored by receiver.",
    },
    {
        "iei": "Octet 1 (b3)",
        "name": "Extension Header flag (E)",
        "presence": "M",
        "format": "Flag",
        "length": "1 bit",
        "field_description": "Indicates presence of a meaningful Next Extension Header Type field in Octet 12 when set to '1'.",
    },
    {
        "iei": "Octet 1 (b2)",
        "name": "Sequence number flag (S)",
        "presence": "M",
        "format": "Flag",
        "length": "1 bit",
        "field_description": "Indicates presence of a meaningful Sequence Number (Octets 9–10). Mandatory '1' for Echo, Error Indication, and Supported Extension Headers Notification; '0' for End Marker and Tunnel Status.",
    },
    {
        "iei": "Octet 1 (b1)",
        "name": "N-PDU Number flag (PN)",
        "presence": "M",
        "format": "Flag",
        "length": "1 bit",
        "field_description": "Indicates presence of a meaningful N-PDU Number field in Octet 11 when set to '1'.",
    },
    {
        "iei": "Octet 2",
        "name": "Message Type",
        "presence": "M",
        "format": "Integer",
        "length": "1 octet",
        "field_description": "Identifies GTP-U message type (Table 6.1-1): 1=Echo Request, 2=Echo Response, 26=Error Indication, 31=Supported Extension Headers Notification, 253=Tunnel Status, 254=End Marker, 255=G-PDU.",
    },
    {
        "iei": "Octets 3-4",
        "name": "Length",
        "presence": "M",
        "format": "Integer",
        "length": "2 octets",
        "field_description": "Payload length in octets following the first 8 mandatory header octets (includes Sequence Number, N-PDU Number, Extension Headers, and T-PDU/IEs).",
    },
    {
        "iei": "Octets 5-8",
        "name": "Tunnel Endpoint Identifier (TEID)",
        "presence": "M",
        "format": "Octet field",
        "length": "4 octets",
        "field_description": "Unambiguously identifies the tunnel endpoint in the receiving GTP-U entity; set to all-zeros for Echo Request/Response, Error Indication, and Supported Extension Headers Notification.",
    },
    {
        "iei": "Octets 9-10",
        "name": "Sequence Number",
        "presence": "C",
        "format": "Integer",
        "length": "2 octets",
        "field_description": "Present if any of S, PN, or E flag is '1'; evaluated only when S='1' (used for signalling transaction matching or ordered T-PDU delivery).",
    },
    {
        "iei": "Octet 11",
        "name": "N-PDU Number",
        "presence": "C",
        "format": "Integer",
        "length": "1 octet",
        "field_description": "Present if any of S, PN, or E flag is '1'; evaluated only when PN='1' (used in 2G/3G inter-SGSN RAU and inter-system handover acknowledgement).",
    },
    {
        "iei": "Octet 12",
        "name": "Next Extension Header Type",
        "presence": "C",
        "format": "Ext Type",
        "length": "1 octet",
        "field_description": "Present if any of S, PN, or E flag is '1'; evaluated only when E='1'. Identifies the first Extension Header (Figure 5.2.1-3), or 0x00 if none follows.",
    },
]

# Cross-spec unrolling targets into TS 38.415 (PDU Session & PDU Set User Plane Protocols)
GTPU_CROSS_SPEC_CONTAINERS: Dict[str, List[Dict[str, str]]] = {
    "pdusessioncontainer": [
        {
            "iei": "PDU Type 0",
            "name": "DL PDU SESSION INFORMATION (PDU Type 0)",
            "type_reference": "TS 38.415 5.5.2.1",
            "presence": "C",
            "format": "Frame",
            "length": "Variable",
            "applicability": "N3, N9, N3mb, N19mb",
            "field_description": (
                "Downlink PDU Session User Plane frame (QFI, RQI, PPP, PPI, QMP timestamps, "
                "SNP/MSNP) specified in TS 38.415 Clause 5.5.2.1."
            ),
        },
        {
            "iei": "PDU Type 1",
            "name": "UL PDU SESSION INFORMATION (PDU Type 1)",
            "type_reference": "TS 38.415 5.5.2.2",
            "presence": "C",
            "format": "Frame",
            "length": "Variable",
            "applicability": "N3, N9",
            "field_description": (
                "Uplink PDU Session User Plane frame (QFI, QMP timestamps, DL/UL/N3-N9 delay "
                "results, congestion info) specified in TS 38.415 Clause 5.5.2.2."
            ),
        },
    ],
    "pdusetinformationcontainer": [
        {
            "iei": "PDU Type 0",
            "name": "DL PDU SET INFORMATION (PDU Type 0)",
            "type_reference": "TS 38.415 6.5.2.1",
            "presence": "C",
            "format": "Frame",
            "length": "Variable",
            "applicability": "N3, N9, Xn-U, F1-U",
            "field_description": (
                "Downlink PDU Set Information User Plane frame (QFI, PSSN, PSI, EDB, EPDU, "
                "PS Size) specified in TS 38.415 Clause 6.5.2.1."
            ),
        },
    ],
}

# Extension headers applicable to Clause 7 signalling messages
GTPU_SIGNALLING_EXT_HEADERS: Dict[str, List[Tuple[str, str, str]]] = {
    "errorindication": [("udpport", "O", "All Interfaces")],
    "endmarker": [("pdusessioncontainer", "C", "N3, N9, Xn-U")],
}

# HTML octet/bit diagram fallbacks for Clause 8 figures stored as WMF/EMF images in TS 29.281 DOCX
_TBL_STYLE = (
    "border-collapse: collapse; margin: 8px 0; border: 1px solid #CBD5E1; "
    "font-family: Consolas, 'Courier New', monospace; font-size: 11px; width: 100%;"
)
_TH_STYLE = (
    "border: 1px solid #CBD5E1; padding: 4px 6px; text-align: center; "
    "background-color: #F1F5F9; font-weight: bold; color: #1E293B;"
)
_TD_STYLE = "border: 1px solid #CBD5E1; padding: 4px 6px; text-align: center;"
_TD_ALT = "border: 1px solid #CBD5E1; padding: 4px 6px; text-align: center; background-color: #FAFAFA;"


def _octet_diagram_table(rows: List[Tuple[str, str, bool]]) -> str:
    """Build a clean bit-grid HTML table matching _convert_table_to_html(is_figure_diagram=True)."""
    parts = [
        f'<table border="1" cellspacing="0" cellpadding="4" style="{_TBL_STYLE}">',
        f'<tr><th style="{_TH_STYLE}">Octets</th>'
        f'<th colspan="8" style="{_TH_STYLE}">Bits 8 &hellip; 1</th></tr>',
    ]
    for idx, (octet_label, content, shaded) in enumerate(rows):
        style = _TD_ALT if shaded or idx % 2 == 0 else _TD_STYLE
        parts.append(
            f'<tr><td style="{style}">{octet_label}</td>'
            f'<td colspan="8" style="{style}">{content}</td></tr>'
        )
    parts.append("</table>")
    return "".join(parts)


CLAUSE_8_EMBEDDED_FIGURE_HTML: Dict[str, str] = {
    "8.2": _octet_diagram_table([
        ("1", "Type = 14 (Decimal)", True),
        ("2", "Restart Counter (0)", False),
    ]),
    "8.4": _octet_diagram_table([
        ("1", "Type = 133 (Decimal)", True),
        ("2–3", "Length = 4 (IPv4) or 16 (IPv6)", False),
        ("4–(n+3)", "IPv4 or IPv6 Address", True),
    ]),
    "8.5": _octet_diagram_table([
        ("1", "Type = 141 (Decimal)", True),
        ("2", "Length = n (Number of Extension Header Types)", False),
        ("3", "Extension Header Type 1", True),
        ("&hellip;", "&hellip;", False),
        ("n+2", "Extension Header Type n", True),
    ]),
    "8.6": _octet_diagram_table([
        ("1", "Type = 255 (Decimal)", True),
        ("2–3", "Length", False),
        ("4–5", "Extension Identifier (IANA Private Enterprise Number)", True),
        ("6–m", "Extension Value", False),
    ]),
}
