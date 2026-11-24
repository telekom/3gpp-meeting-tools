"""3GPP working-group FTP/DynaReport sources and meeting folder helpers."""

from __future__ import annotations

import re

MEETING_SOURCES = {
    "RAN": {
        "ftp": [
            "https://www.3gpp.org/ftp/tsg_ran/TSG_RAN/",
            "https://www.3gpp.org/ftp/tsg_ran/TSG_RAN/TSGR_AHs/",
        ],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-RP.htm",
    },
    "RAN1": {
        "ftp": [
            "https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/",
            "https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_AH/",
        ],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R1.htm",
    },
    "RAN2": {
        "ftp": [
            "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/",
            "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_AHs/",
        ],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R2.htm",
    },
    "RAN3": {
        "ftp": [
            "https://www.3gpp.org/ftp/tsg_ran/WG3_Iu/",
            "https://www.3gpp.org/ftp/tsg_ran/WG3_Iu/TSGR3_AHGs/",
        ],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R3.htm",
    },
    "RAN4": {
        "ftp": [
            "https://www.3gpp.org/ftp/tsg_ran/WG4_Radio/",
            "https://www.3gpp.org/ftp/tsg_ran/WG4_Radio/TSGR4_AHs/",
        ],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R4.htm",
    },
    "SA": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/TSG_SA/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-SP.htm",
    },
    "SA1": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG1_Serv/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S1.htm",
    },
    "SA2": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S2.htm",
    },
    "SA3": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG3_Security/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S3.htm",
    },
    "SA4": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S4.htm",
    },
    "SA5": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG5_TM/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S5.htm",
    },
    "SA6": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG6_MissionCritical/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S6.htm",
    },
    "CT": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/TSG_CT/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-CP.htm",
    },
    "CT1": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG1_mm-cc-sm_ex-CN1/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C1.htm",
    },
    "CT2": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG2_capability_ex-T2/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C2.htm",
    },
    "CT3": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG3_interworking_ex-CN3/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C3.htm",
    },
    "CT4": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG4_protocollars_ex-CN4/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C4.htm",
    },
    "CT5": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG5_osa_ex-CN5/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C5.htm",
    },
    "CT6": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG6_Smartcard_Ex-T3/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C6.htm",
    },
}

MEETING_NAME_PATTERN = re.compile(r"^[A-Z0-9]+_(?:\d+|AH)", re.IGNORECASE)
MEETING_NUM_PATTERN = re.compile(
    r"(?:^|_)(AH\d*|\d+(?:[a-z]+)?(?:-?AH)?(?:-?e)?)(?:_|-|$)", re.IGNORECASE
)
HREF_PATTERN = re.compile(r'href=["\']([^"\'>]+)["\']', re.IGNORECASE)
TDOC_LINK_PATTERN = re.compile(r">\s*([^<]+\.(?:zip|doc|docx|pdf))\s*</a>", re.IGNORECASE)
FILE_EXTENSION_PATTERN = re.compile(r"\.[a-z0-9]{2,4}$", re.IGNORECASE)

MEETING_DENY_LIST = {
    "cr_implementation",
    "tor",
    "tool_automation_6g",
    "specifications",
    "r2_tss_trs_early_versions",
    "outgoing_liaisons",
    "_doc_list_archive",
    "approved_reports",
    "docs",
    "latest_sa2_specs",
}


def is_meeting_name(
    folder_name: str,
    deny_list=MEETING_DENY_LIST,
    pattern=MEETING_NAME_PATTERN,
) -> bool:
    """Return True if an FTP folder name represents a 3GPP meeting directory."""
    if folder_name.lower() in deny_list:
        return False
    if FILE_EXTENSION_PATTERN.search(folder_name):
        return False
    return bool(pattern.match(folder_name))


def extract_meeting_number(folder_name: str, pattern=MEETING_NUM_PATTERN) -> str:
    """Normalize a 3GPP meeting number (or AH tag) from a directory name."""
    match = pattern.search(folder_name)
    if match:
        raw = match.group(1).replace("-", "").upper()
        return re.sub(r"(?<!\d)0+(\d)", r"\1", raw)
    return folder_name
