"""
Alias for core.utils.portal_utils providing seamless access to 3GPP Portal URL generators.
"""
from core.utils.portal_utils import (
    PORTAL_BASE_URL,
    extract_base_tdoc,
    get_portal_tdoc_url,
    get_portal_create_tdoc_url,
    get_portal_meeting_url,
    get_portal_doc_list_url,
    get_portal_work_item_url,
    get_portal_wi_specs_url,
    get_portal_wi_crs_url,
    open_portal_tdoc,
    open_portal_meeting,
    open_portal_create_tdoc,
    open_portal_work_item,
)

__all__ = [
    "PORTAL_BASE_URL",
    "extract_base_tdoc",
    "get_portal_tdoc_url",
    "get_portal_create_tdoc_url",
    "get_portal_meeting_url",
    "get_portal_doc_list_url",
    "get_portal_work_item_url",
    "get_portal_wi_specs_url",
    "get_portal_wi_crs_url",
    "open_portal_tdoc",
    "open_portal_meeting",
    "open_portal_create_tdoc",
    "open_portal_work_item",
]
