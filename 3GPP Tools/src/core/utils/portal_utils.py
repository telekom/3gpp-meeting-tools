"""
3GPP Portal URL and Resource Utilities.

Provides a centralized, single source of truth for constructing and launching
official 3GPP (3GU) portal URLs:
- TDoc / Contribution view: CreateTdoc.aspx?mode=view&contributionUid={uid}
- TDoc reservation / creation: CreateTdoc.aspx?mode=create&meetingId={meeting_id}
- Meeting portal view: Home.aspx#/meeting?MtgId={meeting_id}
- Meeting document list generation: GenerateDocumentList.Aspx?meetingId={meeting_id}
- Work Item details: desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}
- Work Item specifications: Specifications.aspx?q=1&WiUid={wi_code}
- Work Item CRs: ChangeRequests.aspx?q=1&workitem={wi_code}
"""
from __future__ import annotations

import logging
import re
import webbrowser
from typing import Any, Optional, Union

logger = logging.getLogger(__name__)

PORTAL_BASE_URL = "https://portal.3gpp.org"


def extract_base_tdoc(tdoc_id: str) -> str:
    """Extracts the base TDoc identifier, removing revision suffixes.

    Examples:
        'S2-2611250'     -> 'S2-2611250'
        'S2-2611250r01'  -> 'S2-2611250'
        's2-2611250R2'   -> 'S2-2611250'
        'SP-230456_rev1' -> 'SP-230456'
    """
    if not tdoc_id:
        return ""
    clean = str(tdoc_id).strip()
    match = re.match(r"^([A-Za-z0-9]+-\d+)", clean)
    if match:
        return match.group(1).upper()
    base = re.split(r"(?:r|rev)\d+", clean, maxsplit=1, flags=re.IGNORECASE)[0].rstrip("-_")
    return (base or clean).upper()


def get_portal_tdoc_url(tdoc_id: str, base_only: bool = True) -> str:
    """Returns the official 3GU Portal view URL for a 3GPP contribution (TDoc).

    Note: The 3GPP portal requires 'contributionUid', not 'tdocId'.
    Example:
        https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S2-2611250
    """
    clean = str(tdoc_id or "").strip()
    if not clean:
        return ""
    uid = extract_base_tdoc(clean) if base_only else clean
    return f"{PORTAL_BASE_URL}/ngppapp/CreateTdoc.aspx?mode=view&contributionUid={uid}"


def get_portal_create_tdoc_url(meeting_id: Union[int, str]) -> str:
    """Returns the official 3GU Portal URL to create or reserve a TDoc for a meeting.

    Example:
        https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=create&meetingId=3210
    """
    clean_m = str(meeting_id or "").strip()
    if not clean_m:
        return ""
    return f"{PORTAL_BASE_URL}/ngppapp/CreateTdoc.aspx?mode=create&meetingId={clean_m}"


def get_portal_meeting_url(meeting_id: Union[int, str]) -> str:
    """Returns the official 3GU Portal meeting workspace URL.

    Example:
        https://portal.3gpp.org/Home.aspx#/meeting?MtgId=3210
    """
    clean_m = str(meeting_id or "").strip()
    if not clean_m:
        return ""
    return f"{PORTAL_BASE_URL}/Home.aspx#/meeting?MtgId={clean_m}"


def get_portal_doc_list_url(meeting_id: Union[int, str]) -> str:
    """Returns the official 3GU Portal document list generation URL.

    Example:
        https://portal.3gpp.org/ngppapp/GenerateDocumentList.Aspx?meetingId=3210
    """
    clean_m = str(meeting_id or "").strip()
    if not clean_m:
        return ""
    return f"{PORTAL_BASE_URL}/ngppapp/GenerateDocumentList.Aspx?meetingId={clean_m}"


def get_portal_work_item_url(work_item_id: Union[int, str]) -> str:
    """Returns the official 3GU Portal Work Item details URL.

    Example:
        https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123
    """
    clean_code = str(work_item_id or "").strip().lstrip("#")
    if not clean_code:
        return ""
    return f"{PORTAL_BASE_URL}/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={clean_code}"


def get_portal_wi_specs_url(work_item_id: Union[int, str]) -> str:
    """Returns the 3GU Portal URL filtering specifications impacted by a Work Item."""
    clean_code = str(work_item_id or "").strip().lstrip("#")
    if not clean_code:
        return ""
    return f"{PORTAL_BASE_URL}/Specifications.aspx?q=1&WiUid={clean_code}"


def get_portal_wi_crs_url(work_item_id: Union[int, str]) -> str:
    """Returns the 3GU Portal URL filtering Change Requests impacted by a Work Item."""
    clean_code = str(work_item_id or "").strip().lstrip("#")
    if not clean_code:
        return ""
    return f"{PORTAL_BASE_URL}/ChangeRequests.aspx?q=1&workitem={clean_code}"


def open_portal_tdoc(tdoc_id: str, base_only: bool = True) -> bool:
    """Opens the official 3GU portal view for a TDoc in the default browser."""
    url = get_portal_tdoc_url(tdoc_id, base_only=base_only)
    if url:
        logger.info(f"🌐 [Portal] Opening 3GU TDoc view: {url}")
        return webbrowser.open(url)
    return False


def open_portal_meeting(meeting_id: Union[int, str]) -> bool:
    """Opens the official 3GU meeting workspace in the default browser."""
    url = get_portal_meeting_url(meeting_id)
    if url:
        logger.info(f"🖥️ [Portal] Opening 3GU Meeting view: {url}")
        return webbrowser.open(url)
    return False


def open_portal_create_tdoc(meeting_id: Union[int, str]) -> bool:
    """Opens the official 3GU TDoc creation/reservation view for a meeting."""
    url = get_portal_create_tdoc_url(meeting_id)
    if url:
        logger.info(f"📝 [Portal] Opening 3GU Reserve/Create TDoc: {url}")
        return webbrowser.open(url)
    return False


def open_portal_work_item(work_item_id: Union[int, str]) -> bool:
    """Opens the official 3GU Work Item details page in the default browser."""
    url = get_portal_work_item_url(work_item_id)
    if url:
        logger.info(f"📋 [Portal] Opening 3GU Work Item: {url}")
        return webbrowser.open(url)
    return False
