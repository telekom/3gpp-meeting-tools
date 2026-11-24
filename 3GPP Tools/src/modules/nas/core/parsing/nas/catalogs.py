"""Constants, regex patterns, and UE Policy synthetic catalogs for NAS DOCX parsing."""

from __future__ import annotations

import re

from modules.nas.core.parsing.protocol_parser_constants import W_NS

TAG_PSTYLE = f"{W_NS}pStyle"
TAG_INSTR = f"{W_NS}instrText"
ATTR_VAL = f"{W_NS}val"

RE_CAPTION = re.compile(
    r"^Table\s+((?:8|D\.5)\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
    re.IGNORECASE,
)
RE_MSG_HEADING_NAS = re.compile(r"^((?:8\.(?:[2-9]|1[0-9])|D\.5)(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_IE_HEADING_NAS = re.compile(r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_IE_HEADING_POLICY = re.compile(r"^((?:5\.(?:[2-9]|1[0-4])|12\.[1-9])(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_PARENT_SECTION_POLICY = re.compile(r"^((?:5\.(?:[2-9]|1[0-4])|12\.[1-9]))\s+(.+)$")
RE_MAJOR_BOUNDARY_NAS = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-CE-Z])\b")
RE_MAJOR_BOUNDARY_POLICY = re.compile(r"^(?:[1-4]|[6-9]|10|11|13|Annex\s+[A-Z])\b")
RE_TRAILING_PAGE_NUM = re.compile(r"\s+\d{1,4}$")

# Structural umbrella headings in TS 24.501 / TS 24.301 / TS 24.526 that group subclauses rather than defining a single IE
STRUCTURAL_SECTION_CLAUSES = {
    "5.1", "5.3", "5.3.3",
    "9.9", "9.9.1", "9.9.2", "9.9.3", "9.9.4",
    "9.11", "9.11.1", "9.11.2", "9.11.3", "9.11.4",
    "9.12", "D.6",
}

# Encapsulated payload types carried inside TS 24.501 Payload container (9.11.3.39 / Table 9.11.3.40.1)
PAYLOAD_CONTAINER_SUB_IES = [
    ("0100", "SOR transparent container", "9.11.3.51", "C", "TLV-E", "20-n"),
    ("0101", "UE policy section management list", "D.6.2", "C", "TLV-E", "5-65538"),
    ("0110", "UE parameters update transparent container", "9.11.3.53A", "C", "TLV-E", "20-n"),
    ("1000", "CIoT small data container", "9.11.3.18B", "C", "TLV-E", "5-n"),
    ("1001", "Service-level-AA container", "9.11.2.10", "C", "TLV-E", "4-65538"),
]

# UE policy part types carried inside TS 24.501 UE policy section management list (D.6.2 / Table D.6.2.1)
UE_POLICY_PART_SUB_IES = [
    ("0001", "URSP (UE route selection policy)", "TS 24.526 5.2", "C", "LV-E", "3-65535"),
    ("0010", "ANDSP (Access network discovery and selection policy)", "TS 24.526 5.3.1", "C", "LV-E", "3-65535"),
    ("0011", "V2XP (V2X policy)", "TS 24.588 5.2", "C", "LV-E", "3-65535"),
    ("0100", "ProSeP (5G ProSe policy)", "TS 24.555 5.2", "C", "LV-E", "3-65535"),
    ("0101", "A2XP (A2X policy)", "TS 24.578 5.2", "C", "LV-E", "3-65535"),
    ("0110", "RSLPP (Ranging/sidelink positioning policy)", "TS 24.514 12.2", "C", "LV-E", "3-65535"),
]

# Sub-parameters carried inside TS 24.526 UE policy rules (Clause 5.2 URSP & Clause 5.3 ANDSP)
TS24526_POLICY_MESSAGES = [
    {
        "clause": "5.2",
        "message_name": "URSP rule (UE route selection policy)",
        "table_caption": "Table 5.2.1: UE policy part contents including a URSP rule",
        "description": "UE Route Selection Policy (URSP) rule associating application traffic descriptors with a prioritized list of Route Selection Descriptors (RSD) for PDU session binding.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "Length of UE policy part contents", "field_path": "Length of UE policy part contents", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "-", "information_element": "Length of URSP rule", "field_path": "Length of URSP rule", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "-", "information_element": "Precedence value of URSP rule", "field_path": "Precedence value of URSP rule", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "V", "length": "1"},
            {"iei": "-", "information_element": "Traffic descriptor", "field_path": "Traffic descriptor", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "LV-E", "length": "2-65535"},
            {"iei": "00000001", "information_element": "Match-all type", "field_path": "Traffic descriptor > Match-all type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "T", "length": "1"},
            {"iei": "00001000", "information_element": "OS Id + OS App Id type", "field_path": "Traffic descriptor > OS Id + OS App Id type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "19-n"},
            {"iei": "00010000", "information_element": "IPv4 remote address type", "field_path": "Traffic descriptor > IPv4 remote address type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "9"},
            {"iei": "00100001", "information_element": "IPv6 remote address/prefix length type", "field_path": "Traffic descriptor > IPv6 remote address/prefix length type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "18"},
            {"iei": "00110000", "information_element": "Protocol identifier/next header type", "field_path": "Traffic descriptor > Protocol identifier/next header type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "01010000", "information_element": "Single remote port type", "field_path": "Traffic descriptor > Single remote port type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "3"},
            {"iei": "01010001", "information_element": "Remote port range type", "field_path": "Traffic descriptor > Remote port range type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "5"},
            {"iei": "01100000", "information_element": "Security parameter index type", "field_path": "Traffic descriptor > Security parameter index type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "5"},
            {"iei": "01110000", "information_element": "Type of service/traffic class type", "field_path": "Traffic descriptor > Type of service/traffic class type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "3"},
            {"iei": "10000000", "information_element": "Flow label type", "field_path": "Traffic descriptor > Flow label type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "4"},
            {"iei": "10000001", "information_element": "Destination MAC address type", "field_path": "Traffic descriptor > Destination MAC address type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "7"},
            {"iei": "10000011", "information_element": "802.1Q C-TAG VID type", "field_path": "Traffic descriptor > 802.1Q C-TAG VID type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "3"},
            {"iei": "10000100", "information_element": "802.1Q S-TAG VID type", "field_path": "Traffic descriptor > 802.1Q S-TAG VID type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "3"},
            {"iei": "10000101", "information_element": "802.1Q C-TAG PCP/DEI type", "field_path": "Traffic descriptor > 802.1Q C-TAG PCP/DEI type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "10000110", "information_element": "802.1Q S-TAG PCP/DEI type", "field_path": "Traffic descriptor > 802.1Q S-TAG PCP/DEI type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "10000111", "information_element": "Ethertype type", "field_path": "Traffic descriptor > Ethertype type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "3"},
            {"iei": "10001000", "information_element": "DNN type", "field_path": "Traffic descriptor > DNN type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "3-102"},
            {"iei": "10010000", "information_element": "Connection capabilities type", "field_path": "Traffic descriptor > Connection capabilities type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "3-n"},
            {"iei": "10100001", "information_element": "Destination FQDN", "field_path": "Traffic descriptor > Destination FQDN", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "5-256"},
            {"iei": "10100010", "information_element": "Regular expression", "field_path": "Traffic descriptor > Regular expression", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV-E", "length": "4-65535"},
            {"iei": "10110000", "information_element": "OS App Id type", "field_path": "Traffic descriptor > OS App Id type", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "3-257"},
            {"iei": "-", "information_element": "Route selection descriptor list", "field_path": "Route selection descriptor list", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "LV-E", "length": "2-65535"},
            {"iei": "-", "information_element": "Route selection descriptor", "field_path": "Route selection descriptor list > Route selection descriptor", "depth": 1, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "M", "format": "LV-E", "length": "5-65535"},
            {"iei": "00000001", "information_element": "SSC mode type", "field_path": "Route selection descriptor list > Route selection descriptor > SSC mode type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "00000010", "information_element": "S-NSSAI type", "field_path": "Route selection descriptor list > Route selection descriptor > S-NSSAI type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "3-10"},
            {"iei": "00000100", "information_element": "DNN type", "field_path": "Route selection descriptor list > Route selection descriptor > DNN type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "3-102"},
            {"iei": "00001000", "information_element": "PDU session type type", "field_path": "Route selection descriptor list > Route selection descriptor > PDU session type type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "00010000", "information_element": "Preferred access type type", "field_path": "Route selection descriptor list > Route selection descriptor > Preferred access type type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "00010001", "information_element": "Multi-access preference type", "field_path": "Route selection descriptor list > Route selection descriptor > Multi-access preference type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "T", "length": "1"},
            {"iei": "00100000", "information_element": "Non-seamless non-3GPP offload indication type", "field_path": "Route selection descriptor list > Route selection descriptor > Non-seamless non-3GPP offload indication type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "T", "length": "1"},
            {"iei": "01000000", "information_element": "Location criteria type", "field_path": "Route selection descriptor list > Route selection descriptor > Location criteria type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV-E", "length": "7-65535"},
            {"iei": "10000000", "information_element": "Time window type", "field_path": "Route selection descriptor list > Route selection descriptor > Time window type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "17"},
            {"iei": "00100001", "information_element": "Route selection validation criteria type", "field_path": "Route selection descriptor list > Route selection descriptor > Route selection validation criteria type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TLV", "length": "5-21"},
            {"iei": "00100010", "information_element": "PDU session pair ID type", "field_path": "Route selection descriptor list > Route selection descriptor > PDU session pair ID type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "00100011", "information_element": "RSN type", "field_path": "Route selection descriptor list > Route selection descriptor > RSN type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "TV", "length": "2"},
            {"iei": "00100100", "information_element": "ProSe layer-3 UE-to-network relay offload indication type", "field_path": "Route selection descriptor list > Route selection descriptor > ProSe layer-3 UE-to-network relay offload indication type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "T", "length": "1"},
            {"iei": "00100101", "information_element": "5G ProSe multi-path preference type", "field_path": "Route selection descriptor list > Route selection descriptor > 5G ProSe multi-path preference type", "depth": 2, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "T", "length": "1"},
            {"iei": "-", "information_element": "Additional indications (URERI)", "field_path": "Additional indications (URERI)", "depth": 0, "type_reference": "Encoding of UE policy part type URSP 5.2", "presence": "O", "format": "V", "length": "1"},
        ],
    },
    {
        "clause": "5.3",
        "message_name": "ANDSP (Access network discovery and selection policy)",
        "table_caption": "Table 5.3.1.1: ANDSP information format",
        "description": "Access Network Discovery and Selection Policy (ANDSP) used by the UE for selecting non-3GPP access networks (WLANSP and N3AN node configuration).",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "0001", "information_element": "WLANSP (WLAN selection policy)", "field_path": "WLANSP (WLAN selection policy)", "depth": 0, "type_reference": "Encoding of WLANSP 5.3.2", "presence": "C", "format": "LV-E", "length": "3-65535"},
            {"iei": "0010", "information_element": "N3AN node configuration information", "field_path": "N3AN node configuration information", "depth": 0, "type_reference": "Encoding of N3AN node configuration information — General 5.3.3.1", "presence": "C", "format": "LV-E", "length": "3-65535"},
            {"iei": "01", "information_element": "N3AN node selection information", "field_path": "N3AN node configuration information > N3AN node selection information", "depth": 1, "type_reference": "N3AN node selection information 5.3.3.2", "presence": "O", "format": "TLV-E", "length": "9-65535"},
            {"iei": "02", "information_element": "Home N3IWF identifier configuration", "field_path": "N3AN node configuration information > Home N3IWF identifier configuration", "depth": 1, "type_reference": "Home N3IWF identifier configuration 5.3.3.3", "presence": "O", "format": "TLV-E", "length": "10-65535"},
            {"iei": "03", "information_element": "Home ePDG identifier configuration", "field_path": "N3AN node configuration information > Home ePDG identifier configuration", "depth": 1, "type_reference": "Home ePDG identifier configuration 5.3.3.4", "presence": "O", "format": "TLV-E", "length": "10-65535"},
            {"iei": "04", "information_element": "Extended home N3IWF identifier configuration", "field_path": "N3AN node configuration information > Extended home N3IWF identifier configuration", "depth": 1, "type_reference": "Extended home N3IWF identifier configuration 5.3.3.5", "presence": "O", "format": "TLV-E", "length": "11-65535"},
            {"iei": "05", "information_element": "Slice-specific N3IWF prefix configuration", "field_path": "N3AN node configuration information > Slice-specific N3IWF prefix configuration", "depth": 1, "type_reference": "Slice-specific N3IWF prefix configuration 5.3.3.6", "presence": "O", "format": "TLV-E", "length": "13-65535"},
        ],
    },
    {
        "clause": "5.3.2",
        "message_name": "WLANSP (WLAN selection policy)",
        "table_caption": "Table 5.3.2.1: WLANSP information element",
        "description": "WLAN Selection Policy (WLANSP) rule containing prioritized selection criteria, validity area, and time of day for non-3GPP WLAN access selection.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "WLANSP rule priority", "field_path": "WLANSP rule priority", "depth": 0, "type_reference": "Encoding of WLANSP 5.3.2", "presence": "M", "format": "V", "length": "1"},
            {"iei": "-", "information_element": "Selection criteria", "field_path": "Selection criteria", "depth": 0, "type_reference": "Encoding of WLANSP 5.3.2", "presence": "M", "format": "LV-E", "length": "2-65535"},
            {"iei": "-", "information_element": "Validity area", "field_path": "Validity area", "depth": 0, "type_reference": "Encoding of WLANSP 5.3.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "-", "information_element": "Time of day", "field_path": "Time of day", "depth": 0, "type_reference": "Encoding of WLANSP 5.3.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
]

# Sub-parameters carried inside TS 24.588 V2X policy rules (Clause 5.2 V2XP, Clause 5.3 PC5, Clause 5.4 Uu)
TS24588_POLICY_MESSAGES = [
    {
        "clause": "5.2",
        "message_name": "V2XP policy part (V2X policy)",
        "table_caption": "Table 5.2.1.1: V2XP information format",
        "description": "Vehicle-to-Everything Policy (V2XP) UE policy part provisioning V2X communication parameters over PC5 and Uu reference points.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "Length of V2XP contents", "field_path": "Length of V2XP contents", "depth": 0, "type_reference": "Encoding of V2X policy (V2XP) UE policy part 5.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "0001", "information_element": "UE policies for V2X communication over PC5", "field_path": "UE policies for V2X communication over PC5", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0010", "information_element": "UE policies for V2X communication over Uu", "field_path": "UE policies for V2X communication over Uu", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "5.3",
        "message_name": "V2XP rule for PC5 communication",
        "table_caption": "Table 5.3.1.1: UE policies for V2X communication over PC5",
        "description": "UE policies for V2X communication over the PC5 reference point (E-UTRA-PC5 and NR-PC5 radio parameters, privacy, and DRX).",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "Served by E-UTRA or served by NR", "field_path": "Served by E-UTRA or served by NR", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "Not served by E-UTRA and not served by NR", "field_path": "Not served by E-UTRA and not served by NR", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "Privacy configuration", "field_path": "Privacy configuration", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "V2X communication over PC5 in E-UTRA-PC5", "field_path": "V2X communication over PC5 in E-UTRA-PC5", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "V2X communication over PC5 in NR-PC5", "field_path": "V2X communication over PC5 in NR-PC5", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "06", "information_element": "Default PC5 DRX configuration", "field_path": "Default PC5 DRX configuration", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over PC5 5.3", "presence": "O", "format": "TLV", "length": "4"},
        ],
    },
    {
        "clause": "5.4",
        "message_name": "V2XP rule for Uu communication",
        "table_caption": "Table 5.4.1.1: UE policies for V2X communication over Uu",
        "description": "UE policies for V2X communication over the Uu reference point (PLMN infos, V2X Application Server addresses, and PDU session mapping rules).",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "PLMN infos", "field_path": "PLMN infos", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "V2X service identifier unrelated info", "field_path": "V2X service identifier unrelated info", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "V2X service identifier related info", "field_path": "V2X service identifier related info", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "Default V2X AS address infos", "field_path": "Default V2X AS address infos", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "V2X service identifier to PDU session parameters mapping rules", "field_path": "V2X service identifier to PDU session parameters mapping rules", "depth": 0, "type_reference": "Encoding of UE policies for V2X communication over Uu 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
]

# Sub-parameters carried inside TS 24.555 5G ProSe policy rules (Clause 5.2 ProSeP, Clause 5.3 Direct discovery, Clause 5.4 Direct communication)
TS24555_POLICY_MESSAGES = [
    {
        "clause": "5.2",
        "message_name": "5G ProSe policy (ProSeP) part",
        "table_caption": "Table 5.2.1: ProSeP information format",
        "description": "5G Proximity-based Services Policy (ProSeP) UE policy part provisioning parameters for 5G ProSe direct discovery, direct communication, and UE-to-network/UE-to-UE relays.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "Length of ProSeP contents", "field_path": "Length of ProSeP contents", "depth": 0, "type_reference": "Encoding of 5G ProSe policy UE policy part 5.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "0001", "information_element": "UE policies for 5G ProSe direct discovery", "field_path": "UE policies for 5G ProSe direct discovery", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0010", "information_element": "UE policies for 5G ProSe direct communication", "field_path": "UE policies for 5G ProSe direct communication", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0011", "information_element": "UE policies for 5G ProSe UE-to-network relay UE", "field_path": "UE policies for 5G ProSe UE-to-network relay UE", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe UE-to-network relay UE 5.5", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0100", "information_element": "UE policies for 5G ProSe remote UE", "field_path": "UE policies for 5G ProSe remote UE", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe remote UE 5.6", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0101", "information_element": "UE policies for 5G ProSe usage information reporting", "field_path": "UE policies for 5G ProSe usage information reporting", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe usage information reporting 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0110", "information_element": "UE policies for 5G ProSe UE-to-UE relay UE", "field_path": "UE policies for 5G ProSe UE-to-UE relay UE", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe UE-to-UE relay UE 5.8", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0111", "information_element": "UE policies for 5G ProSe end UE", "field_path": "UE policies for 5G ProSe end UE", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe end UE 5.9", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "5.3",
        "message_name": "ProSe policy for direct discovery",
        "table_caption": "Table 5.3.2.1: ProSeP Info = {UE policies for 5G ProSe direct discovery}",
        "description": "UE policies for 5G ProSe direct discovery when served or not served by NG-RAN, including DRX, group member discovery, and 5G DDNMF address.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "Served by NG-RAN", "field_path": "Served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "Not served by NG-RAN", "field_path": "Not served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "Default PC5 DRX configuration", "field_path": "Default PC5 DRX configuration", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV", "length": "4"},
            {"iei": "04", "information_element": "Group member discovery parameters", "field_path": "Group member discovery parameters", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "HPLMN 5G DDNMF address information", "field_path": "HPLMN 5G DDNMF address information", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct discovery 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "5.4",
        "message_name": "ProSe policy for direct communications",
        "table_caption": "Table 5.4.2.1: ProSeP Info = {UE policies for 5G ProSe direct communication}",
        "description": "UE policies for 5G ProSe direct communication over PC5 in NR-PC5, including privacy configuration and groupcast parameters.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "Served by NG-RAN", "field_path": "Served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "Not served by NG-RAN", "field_path": "Not served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "Privacy configuration", "field_path": "Privacy configuration", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "5G ProSe direct communication over PC5 in NR-PC5", "field_path": "5G ProSe direct communication over PC5 in NR-PC5", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "Groupcast parameters", "field_path": "Groupcast parameters", "depth": 0, "type_reference": "Encoding of UE policies for 5G ProSe direct communications 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
]

# Sub-parameters carried inside TS 24.514 Ranging & Sidelink Positioning policy rules (Clause 12.2 RSLPP)
TS24514_POLICY_MESSAGES = [
    {
        "clause": "12.2",
        "message_name": "RSLPP policy part (Ranging & sidelink positioning)",
        "table_caption": "Table 12.2.1: UE policy part when UE policy part type = {RSLPP}",
        "description": "Ranging-based services and Sidelink Positioning Policy (RSLPP) UE policy part provisioning configuration for ranging and PC5 sidelink positioning.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "Length of RSLPP contents", "field_path": "Length of RSLPP contents", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "0001", "information_element": "RSLPP contents (Configuration information)", "field_path": "RSLPP contents (Configuration information)", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "M", "format": "LV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "12.2.2",
        "message_name": "RSLPP contents & configuration",
        "table_caption": "Table 12.2.2: RSLPP contents",
        "description": "Detailed RSLPP configuration parameters including NG-RAN authorization, 5G ProSe and V2X QoS mapping rules, and SLPKMF address information.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "Served by NG-RAN", "field_path": "Served by NG-RAN", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "Not served by NG-RAN", "field_path": "Not served by NG-RAN", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "5G ProSe related mapping rules", "field_path": "5G ProSe related mapping rules", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "V2X service related mapping rules", "field_path": "V2X service related mapping rules", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "SLPKMF address information", "field_path": "SLPKMF address information", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
]

# Sub-parameters carried inside TS 24.578 Aerial Access (A2X) policy rules (Clause 5.2 A2XP, Clause 5.3 PC5, Clause 5.7 Uu)
TS24578_POLICY_MESSAGES = [
    {
        "clause": "5.2",
        "message_name": "A2XP policy part (Aerial access policy)",
        "table_caption": "Table 5.2.1.1: A2XP information format",
        "description": "Aerial-to-Everything Policy (A2XP) UE policy part provisioning UAV parameters for A2X communication over PC5 and Uu, Broadcast Remote ID (BRID), DDAA, and C2.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "-", "information_element": "Length of A2XP contents", "field_path": "Length of A2XP contents", "depth": 0, "type_reference": "Encoding of A2X policy (A2XP) UE policy part 5.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "0001", "information_element": "UE policies for A2X communication over PC5", "field_path": "UE policies for A2X communication over PC5", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0010", "information_element": "UE policies for Broadcast Remote ID (BRID)", "field_path": "UE policies for Broadcast Remote ID (BRID)", "depth": 0, "type_reference": "Encoding of UE policies for Broadcast Remote ID (BRID) 5.4", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0011", "information_element": "UE policies for direct detect and avoid (DDAA) over PC5", "field_path": "UE policies for direct detect and avoid (DDAA) over PC5", "depth": 0, "type_reference": "Encoding of UE policies for direct detect and avoid (DDAA) over PC5 5.5", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0100", "information_element": "UE policies for direct C2 communication over PC5", "field_path": "UE policies for direct C2 communication over PC5", "depth": 0, "type_reference": "Encoding of UE policies for direct C2 communication over PC5 5.6", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0101", "information_element": "UE policies for A2X communication over Uu", "field_path": "UE policies for A2X communication over Uu", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "0110", "information_element": "UE policies for ground based detect and avoid for an area (GBDAAA)", "field_path": "UE policies for ground based detect and avoid for an area (GBDAAA)", "depth": 0, "type_reference": "Encoding of UE policies for ground based detect and avoid for an area (GBDAAA) 5.8", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "5.3",
        "message_name": "A2XP rule for PC5 communication",
        "table_caption": "Table 5.3.2.1: A2XP info = {UE policies for A2X communication over PC5}",
        "description": "UE policies for A2X communication over PC5, including radio parameters per altitude range and geographical area.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "Served by NG-RAN", "field_path": "Served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "Not served by NG-RAN", "field_path": "Not served by NG-RAN", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "Radio parameters per altitude range per geographical area list", "field_path": "Radio parameters per altitude range per geographical area list", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "Privacy configuration", "field_path": "Privacy configuration", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "A2X communication over PC5 in NR-PC5", "field_path": "A2X communication over PC5 in NR-PC5", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "06", "information_element": "Default PC5 DRX configuration", "field_path": "Default PC5 DRX configuration", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over PC5 5.3", "presence": "O", "format": "TLV", "length": "4"},
        ],
    },
    {
        "clause": "5.7",
        "message_name": "A2XP rule for Uu communication",
        "table_caption": "Table 5.7.2.1: A2XP Info = {UE policies for A2X communication over Uu}",
        "description": "UE policies for A2X communication over Uu, including PLMN infos, A2X Application Server addresses, and PDU session mapping rules.",
        "direction": "PCF → UE (via AMF)",
        "ies": [
            {"iei": "01", "information_element": "PLMN infos", "field_path": "PLMN infos", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "02", "information_element": "A2X service identifier unrelated info", "field_path": "A2X service identifier unrelated info", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "03", "information_element": "A2X service identifier related info", "field_path": "A2X service identifier related info", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "04", "information_element": "Default A2X AS address infos", "field_path": "Default A2X AS address infos", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
            {"iei": "05", "information_element": "A2X service identifier to PDU session parameters mapping rules", "field_path": "A2X service identifier to PDU session parameters mapping rules", "depth": 0, "type_reference": "Encoding of UE policies for A2X communication over Uu 5.7", "presence": "O", "format": "TLV-E", "length": "3-65535"},
        ],
    },
]

# Sub-parameters carried inside TS 24.501 Service-level-AA container (9.11.2.10 / Table 9.11.2.10.1)
SERVICE_LEVEL_AA_SUB_IES = [
    ("10", "Service-level device ID", "9.11.2.11", "O", "TLV", "3-n"),
    ("20", "Service-level-AA server address", "9.11.2.12", "O", "TLV", "3-n"),
    ("30", "Service-level-AA response", "9.11.2.14", "O", "TLV", "3"),
    ("40", "Service-level-AA payload type", "9.11.2.15", "O", "TLV", "3"),
    ("70", "Service-level-AA payload", "9.11.2.13", "O", "TLV-E", "4-65538"),
    ("A-", "Service-level-AA pending indication", "9.11.2.17", "O", "TV", "1"),
    ("50", "Service-level-AA service status indication", "9.11.2.18", "O", "TLV", "3"),
]

# Sub-parameters carried inside Extended protocol configuration options (9.11.4.6) defined in Clause 9
EPCO_SUB_IES = [
    ("0031", "ECS address", "9.11.4.34", "O", "TLV", "4-n"),
]
