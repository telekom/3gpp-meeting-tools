"""
Unit tests for core.utils.portal_utils (and core.utils.portal_urls).
Verifies correct construction of 3GPP (3GU) Portal URLs, including the critical
fix from 'tdocId' to 'contributionUid'.
"""
import unittest
from unittest.mock import patch

from core.utils.portal_utils import (
    PORTAL_BASE_URL,
    extract_base_tdoc,
    get_portal_create_tdoc_url,
    get_portal_doc_list_url,
    get_portal_meeting_url,
    get_portal_tdoc_url,
    get_portal_wi_crs_url,
    get_portal_wi_specs_url,
    get_portal_work_item_url,
    open_portal_create_tdoc,
    open_portal_meeting,
    open_portal_tdoc,
    open_portal_work_item,
)
import core.utils.portal_urls as portal_urls_alias


class TestPortalUtils(unittest.TestCase):
    def test_extract_base_tdoc(self):
        self.assertEqual(extract_base_tdoc("S2-2611250"), "S2-2611250")
        self.assertEqual(extract_base_tdoc("S2-2611250r01"), "S2-2611250")
        self.assertEqual(extract_base_tdoc("s2-2611250R2"), "S2-2611250")
        self.assertEqual(extract_base_tdoc("SP-230456_rev1"), "SP-230456")
        self.assertEqual(extract_base_tdoc("RP-240123-r02"), "RP-240123")
        self.assertEqual(extract_base_tdoc("  C1-240999r05  "), "C1-240999")
        self.assertEqual(extract_base_tdoc(""), "")
        self.assertEqual(extract_base_tdoc(None), "")

    def test_get_portal_tdoc_url_uses_contribution_uid(self):
        """Verifies that the URL uses contributionUid and not the obsolete/broken tdocId parameter."""
        url = get_portal_tdoc_url("S2-2611250")
        self.assertEqual(
            url,
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S2-2611250",
        )
        self.assertNotIn("tdocId=", url)

    def test_get_portal_tdoc_url_revision_handling(self):
        # Default: base_only=True extracts base TDoc
        url_rev = get_portal_tdoc_url("S2-2611250r01")
        self.assertEqual(
            url_rev,
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S2-2611250",
        )

        # base_only=False preserves original identifier
        url_exact = get_portal_tdoc_url("S2-2611250r01", base_only=False)
        self.assertEqual(
            url_exact,
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S2-2611250r01",
        )

        # Empty returns empty string
        self.assertEqual(get_portal_tdoc_url(""), "")
        self.assertEqual(get_portal_tdoc_url(None), "")

    def test_get_portal_create_tdoc_url(self):
        url = get_portal_create_tdoc_url(3210)
        self.assertEqual(
            url,
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=create&meetingId=3210",
        )
        self.assertEqual(get_portal_create_tdoc_url(""), "")

    def test_get_portal_meeting_url(self):
        url = get_portal_meeting_url(3210)
        self.assertEqual(
            url,
            "https://portal.3gpp.org/Home.aspx#/meeting?MtgId=3210",
        )
        self.assertEqual(get_portal_meeting_url(""), "")

    def test_get_portal_doc_list_url(self):
        url = get_portal_doc_list_url(3210)
        self.assertEqual(
            url,
            "https://portal.3gpp.org/ngppapp/GenerateDocumentList.Aspx?meetingId=3210",
        )
        self.assertEqual(get_portal_doc_list_url(""), "")

    def test_get_portal_work_item_urls(self):
        url_wi = get_portal_work_item_url(1121123)
        self.assertEqual(
            url_wi,
            "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123",
        )
        # Handles '#1121123'
        self.assertEqual(
            get_portal_work_item_url("#1121123"),
            "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123",
        )

        url_specs = get_portal_wi_specs_url("1121123")
        self.assertEqual(
            url_specs,
            "https://portal.3gpp.org/Specifications.aspx?q=1&WiUid=1121123",
        )

        url_crs = get_portal_wi_crs_url("1121123")
        self.assertEqual(
            url_crs,
            "https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem=1121123",
        )

    @patch("webbrowser.open")
    def test_browser_open_helpers(self, mock_open):
        open_portal_tdoc("S2-2611250")
        mock_open.assert_called_with(
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=S2-2611250"
        )

        open_portal_meeting(3210)
        mock_open.assert_called_with("https://portal.3gpp.org/Home.aspx#/meeting?MtgId=3210")

        open_portal_create_tdoc(3210)
        mock_open.assert_called_with(
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=create&meetingId=3210"
        )

        open_portal_work_item("1121123")
        mock_open.assert_called_with(
            "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123"
        )

    def test_portal_urls_alias_module(self):
        self.assertEqual(portal_urls_alias.PORTAL_BASE_URL, PORTAL_BASE_URL)
        self.assertEqual(
            portal_urls_alias.get_portal_tdoc_url("S2-2611250"),
            get_portal_tdoc_url("S2-2611250"),
        )


if __name__ == "__main__":
    unittest.main()
