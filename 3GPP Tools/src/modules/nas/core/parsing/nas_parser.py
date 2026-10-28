import html
import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.protocol_parser_utils import extract_document_root, _extract_p_text, _extract_tc_text, \
    _convert_table_to_html
from modules.nas.core.parsing.protocol_parser_constants import TAG_BODY, TAG_P, TAG_PPR, TAG_TBL, TAG_TR, TAG_TC, W_NS

TAG_PSTYLE = f"{W_NS}pStyle"
TAG_INSTR = f"{W_NS}instrText"
ATTR_VAL = f"{W_NS}val"

RE_CAPTION = re.compile(
    r"^Table\s+((?:8|D\.5)\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
    re.IGNORECASE,
)
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
        "ies": [
            {"iei": "-", "information_element": "Length of RSLPP contents", "field_path": "Length of RSLPP contents", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "M", "format": "V", "length": "2"},
            {"iei": "0001", "information_element": "RSLPP contents (Configuration information)", "field_path": "RSLPP contents (Configuration information)", "depth": 0, "type_reference": "RSLPP Information Elements Coding 12.2", "presence": "M", "format": "LV-E", "length": "3-65535"},
        ],
    },
    {
        "clause": "12.2.2",
        "message_name": "RSLPP contents & configuration",
        "table_caption": "Table 12.2.2: RSLPP contents",
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


class NASDocxParser:
    """Dedicated parser for standard 3GPP NAS & UE Policy specifications (TS 24.501, TS 24.301, TS 24.526, TS 24.588, TS 24.555, TS 24.514, TS 24.578)."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "24.501"):
        self.docx_paths = docx_paths
        self.spec_number = str(spec_number or "24.501").replace("TS", "").strip()
        self.is_policy_spec = self.spec_number in ("24.526", "24.588", "24.555", "24.578", "24.514")
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _is_toc_paragraph(p_elem: ET.Element) -> bool:
        """Detects Word Table of Contents (TOC) paragraphs to avoid indexing TOC lines as IE definitions."""
        ppr = p_elem.find(TAG_PPR)
        if ppr is not None:
            ps = ppr.find(TAG_PSTYLE)
            if ps is not None:
                val = (ps.get(ATTR_VAL) or "").lower()
                if val.startswith("toc"):
                    return True
        for instr in p_elem.iter(TAG_INSTR):
            if instr.text and ("PAGEREF" in instr.text or "TOC" in instr.text):
                return True
        return False

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []
        total_files = len(valid_paths)

        is_policy_spec = self.spec_number in ("24.526", "24.588", "24.555", "24.578", "24.514")
        re_ie_heading = RE_IE_HEADING_POLICY if is_policy_spec else RE_IE_HEADING_NAS
        re_major_boundary = RE_MAJOR_BOUNDARY_POLICY if is_policy_spec else RE_MAJOR_BOUNDARY_NAS

        parent_section_titles: Dict[str, str] = {}

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                base_progress = 10 + int((file_idx / total_files) * 80)
                progress_callback(f"Reading {docx_path.name} ({file_idx + 1}/{total_files})...", base_progress)

            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            last_caption_info: Optional[Tuple[str, str, str]] = None
            last_paragraph_text: str = ""
            current_ie_def: Optional[Dict[str, Any]] = None

            for elem in list(body):
                if elem.tag == TAG_P:
                    if self._is_toc_paragraph(elem):
                        continue

                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_paragraph_text = p_text

                    if is_policy_spec:
                        mp_sec = RE_PARENT_SECTION_POLICY.match(p_text)
                        if mp_sec:
                            sec_cl = mp_sec.group(1).strip()
                            sec_title = RE_TRAILING_PAGE_NUM.sub("", mp_sec.group(2).strip()).strip()
                            parent_section_titles[sec_cl] = sec_title

                    if p_text.startswith(("Table 8.", "Table D.5.")):
                        match_cap = RE_CAPTION.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", match_cap.group(2)).strip()
                            last_caption_info = (clause, name, p_text)
                        else:
                            last_caption_info = None
                    else:
                        last_caption_info = None

                    match_ie = re_ie_heading.match(p_text)
                    if match_ie:
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)

                        cl = match_ie.group(1).strip()
                        ie_name = RE_TRAILING_PAGE_NUM.sub("", match_ie.group(2).strip()).strip()
                        if is_policy_spec:
                            parts = cl.split(".")
                            parent_cl = ".".join(parts[:-1]) if len(parts) > 2 else cl
                            parent_title = parent_section_titles.get(parent_cl) or parent_section_titles.get(parts[0])
                            if ie_name.lower() in ("general", "overview", "information elements coding"):
                                if parent_title:
                                    ie_name = f"{parent_title} — {ie_name}"
                                elif cl == "5.3.1":
                                    ie_name = "Encoding of UE policy part type ANDSP — General"
                                elif cl == "5.3.3.1":
                                    ie_name = "Encoding of N3AN node configuration information — General"
                                elif cl.startswith("12.2"):
                                    ie_name = "RSLPP Information Elements Coding"
                                else:
                                    ie_name = f"Clause {cl} — {ie_name}"

                        current_ie_def = {
                            "clause": cl,
                            "ie_name": ie_name,
                            "html_content": [
                                f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif;">{html.escape(ie_name)} (Clause {html.escape(cl)})</h2>'
                            ],
                        }
                    elif re_major_boundary.match(p_text) and not p_text.startswith(("9.", "D.6", "5.", "12.")):
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        if p_text.startswith(("Figure 9.", "Figure D.", "Figure 5.", "Figure 12.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith(("Table 9.", "Table D.", "Table 5.", "Table 12.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #1A237E; margin-top: 12px; margin-bottom: 4px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_def["html_content"].append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{html.escape(p_text)}</div>'
                            )
                        else:
                            current_ie_def["html_content"].append(
                                f'<p style="margin: 4px 0; line-height: 1.4; color: #263238;">{html.escape(p_text)}</p>'
                            )

                elif elem.tag == TAG_TBL:
                    handled_as_message = False
                    if last_caption_info:
                        clause, name, full_caption = last_caption_info
                        ies = self._parse_clause_8_table(elem)
                        if ies:
                            messages.append({
                                "clause": clause,
                                "message_name": name,
                                "table_caption": full_caption,
                                "ies": ies,
                            })
                            handled_as_message = True
                        last_caption_info = None

                    if not handled_as_message and current_ie_def:
                        is_diagram = "figure" in last_paragraph_text.lower() or any(
                            "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC)
                        )
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_diagram)
                        if tbl_html:
                            current_ie_def["html_content"].append(tbl_html)

            if current_ie_def:
                self._finalize_ie_def(current_ie_def, ie_definitions)

        if self.is_policy_spec and ie_definitions and not messages:
            if self.spec_number == "24.526":
                messages.extend(TS24526_POLICY_MESSAGES)
            elif self.spec_number == "24.588":
                messages.extend(TS24588_POLICY_MESSAGES)
            elif self.spec_number == "24.555":
                messages.extend(TS24555_POLICY_MESSAGES)
            elif self.spec_number == "24.514":
                messages.extend(TS24514_POLICY_MESSAGES)
            elif self.spec_number == "24.578":
                messages.extend(TS24578_POLICY_MESSAGES)

        self._unroll_container_sub_ies(messages, ie_definitions)
        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        clause = current_ie_def.get("clause", "").strip()
        ie_name = current_ie_def.get("ie_name", "").strip()
        html_blocks = current_ie_def.get("html_content", [])

        # Skip umbrella section headers, Void clauses, or empty heading stubs
        if not clause or not ie_name or len(html_blocks) <= 1:
            return
        if self.is_policy_spec:
            if clause in ("5.1", "12.1"):
                return
        else:
            if clause in STRUCTURAL_SECTION_CLAUSES:
                return
        lower_name = ie_name.lower()
        if lower_name.startswith("void") or lower_name in ("general", "overview"):
            return
        if not self.is_policy_spec and lower_name == "information elements coding":
            return

        html_str = "".join(html_blocks)
        ie_definitions.append({
            "clause": clause,
            "ie_name": ie_name,
            "raw_description": html_str,
            "structure_table": json.dumps([]),
        })

    def _unroll_container_sub_ies(
        self,
        messages: List[Dict[str, Any]],
        ie_definitions: List[Dict[str, Any]],
    ):
        """
        Unrolls encapsulated NAS container payloads (Payload container, Service-level-AA container,
        Type 6 IE container, ePCO) into depth > 0 child rows when the target Clause 9 / Annex D
        IE definitions exist in the parsed specification version.
        """
        defs_by_clause: Dict[str, Dict[str, Any]] = {
            d["clause"]: d for d in ie_definitions if d.get("clause")
        }
        if not defs_by_clause:
            return

        # Find helper table 8.2.7.54.1.1 (Registration accept type 6 IE container contents) if present
        type6_sub_ies: List[Dict[str, Any]] = []
        for m in messages:
            if m.get("clause", "").startswith("8.2.7.54"):
                type6_sub_ies = m.get("ies", [])
                break

        def _build_slaa_children(parent_path: str, base_depth: int) -> List[Dict[str, Any]]:
            children = []
            for sub_iei, fallback_name, sub_clause, pres, fmt, length in SERVICE_LEVEL_AA_SUB_IES:
                if sub_clause in defs_by_clause:
                    resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                    children.append({
                        "iei": sub_iei,
                        "information_element": resolved_name,
                        "field_path": f"{parent_path} > {resolved_name}",
                        "depth": base_depth,
                        "type_reference": f"{resolved_name} {sub_clause}",
                        "presence": pres,
                        "format": fmt,
                        "length": length,
                    })
            return children

        def _build_ue_policy_children(parent_path: str, base_depth: int) -> List[Dict[str, Any]]:
            children = []
            for sub_iei, part_name, sub_ref, pres, fmt, length in UE_POLICY_PART_SUB_IES:
                children.append({
                    "iei": sub_iei,
                    "information_element": part_name,
                    "field_path": f"{parent_path} > {part_name}",
                    "depth": base_depth,
                    "type_reference": sub_ref,
                    "presence": pres,
                    "format": fmt,
                    "length": length,
                })
            return children

        for msg in messages:
            # Skip unrolling inside the helper table itself
            if msg.get("clause", "").startswith("8.2.7.54"):
                continue

            expanded_ies: List[Dict[str, Any]] = []
            for ie in msg.get("ies", []):
                expanded_ies.append(ie)
                type_ref = ie.get("type_reference", "")
                parent_name = ie.get("information_element", "")

                # 1. Payload container (Clause 9.11.3.39)
                if "9.11.3.39" in type_ref:
                    for sub_iei, fallback_name, sub_clause, pres, fmt, length in PAYLOAD_CONTAINER_SUB_IES:
                        if sub_clause in defs_by_clause:
                            resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                            sub_path = f"{parent_name} > {resolved_name}"
                            expanded_ies.append({
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": sub_path,
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            })
                            if sub_clause == "9.11.2.10":
                                expanded_ies.extend(_build_slaa_children(sub_path, base_depth=2))
                            elif sub_clause == "D.6.2":
                                expanded_ies.extend(_build_ue_policy_children(sub_path, base_depth=2))

                # 2. Top-level Service-level-AA container (Clause 9.11.2.10)
                elif "9.11.2.10" in type_ref:
                    expanded_ies.extend(_build_slaa_children(parent_name, base_depth=1))

                # 3. Top-level UE policy section management list (Clause D.6.2)
                elif "D.6.2" in type_ref and "D.6.2" in defs_by_clause:
                    expanded_ies.extend(_build_ue_policy_children(parent_name, base_depth=1))

                # 4. Registration accept type 6 IE container (Clause 9.11.3.98)
                elif "9.11.3.98" in type_ref and type6_sub_ies:
                    for sub in type6_sub_ies:
                        sub_name = sub.get("information_element", "")
                        expanded_ies.append({
                            "iei": sub.get("iei", ""),
                            "information_element": sub_name,
                            "field_path": f"{parent_name} > {sub_name}",
                            "depth": 1,
                            "type_reference": sub.get("type_reference", ""),
                            "presence": sub.get("presence", ""),
                            "format": sub.get("format", ""),
                            "length": sub.get("length", ""),
                        })

                # 5. Extended protocol configuration options (Clause 9.11.4.6)
                elif "9.11.4.6" in type_ref:
                    for sub_iei, fallback_name, sub_clause, pres, fmt, length in EPCO_SUB_IES:
                        if sub_clause in defs_by_clause:
                            resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                            expanded_ies.append({
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": f"{parent_name} > {resolved_name}",
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            })

            msg["ies"] = expanded_ies

    def _parse_clause_8_table(self, tbl_elem: ET.Element) -> List[Dict[str, str]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_row_idx = -1
        for r_idx in range(min(3, len(rows))):
            cells_text = [_extract_tc_text(tc).lower().strip() for tc in rows[r_idx].findall(TAG_TC)]
            if len(cells_text) < 4:
                continue
            has_ie_col = any(c in ("information element", "information elements", "iei") for c in cells_text)
            has_meta_col = any("reference" in c or "presence" in c or "format" in c for c in cells_text)
            if has_ie_col and has_meta_col:
                header_row_idx = r_idx
                break

        if header_row_idx == -1:
            return []

        ies = []
        for row in rows[header_row_idx + 1:]:
            cells = [_extract_tc_text(tc) for tc in row.findall(TAG_TC)]
            if len(cells) < 6:
                continue

            iei, ie_name, type_ref, presence, fmt, length = cells[:6]
            clean_name = ie_name.lower().replace(" ", "")
            if not clean_name or "informationelement" in clean_name:
                continue

            type_ref = type_ref.strip().rstrip(".")

            ies.append({
                "iei": iei,
                "information_element": ie_name,
                "field_path": ie_name,
                "depth": 0,
                "type_reference": type_ref,
                "presence": presence,
                "format": fmt,
                "length": length,
            })

        return ies
