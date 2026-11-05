# --- File: tests/test_pre_meeting_exporter.py ---
import sys
import unittest
from pathlib import Path

# Ensure src is in sys.path
_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.exporters.pre_meeting_exporter import PreMeetingExporterThread
from modules.meetings.ui.dialogs.pre_meeting_export_dialog import PreMeetingExportDialog


class TestPreMeetingExporter(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        # Comprehensive test dataset featuring:
        # - Followed & Non-followed Work Items
        # - Revision chains (S2-2601001 revised to S2-2601002)
        # - Multi-proposal LS in -> candidate LS out replies
        # - Discussion papers and WIDs
        # - Reserved / unuploaded contributions
        # - Delegate personal stance and notes
        self.rows = [
            # 1. Incoming LS in under TEI18
            {
                "TDoc": "S2-2601000",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "TEI18",
                "Agenda Item": "4.1",
                "Agenda item description": "Incoming LSs",
                "Release": "Rel-18",
                "Title": "LS on Architecture Timelines",
                "Source": "RAN2",
            },
            # 2. Competing candidate reply 1 to S2-2601000 from Ericsson (empty Related WIs -> should inherit TEI18)
            {
                "TDoc": "S2-2601010",
                "Type": "LS out",
                "TDoc Status": "",
                "Related WIs": "",
                "Reply to": "S2-2601000",
                "Agenda Item": "4.1",
                "Agenda item description": "Incoming LSs",
                "Release": "Rel-18",
                "Title": "Proposed Reply LS to RAN2",
                "Source": "Ericsson",
                "_Sanitized_Companies": ["Ericsson"],
                "_All_Sanitized_Companies": ["Ericsson"],
            },
            # 3. Competing candidate reply 2 to S2-2601000 from Huawei (also under AI 4.1)
            {
                "TDoc": "S2-2601020",
                "Type": "LS out",
                "TDoc Status": "",
                "Related WIs": "",
                "Reply to": "S2-2601000",
                "Agenda Item": "4.1",
                "Agenda item description": "Incoming LSs",
                "Release": "Rel-18",
                "Title": "Draft Reply LS on RAN2 Questions",
                "Source": "Huawei",
                "_Sanitized_Companies": ["Huawei"],
                "_All_Sanitized_Companies": ["Huawei"],
            },
            # 4. Old superseded draft pCR (revised to S2-2601002)
            {
                "TDoc": "S2-2601001",
                "Type": "pCR",
                "TDoc Status": "Revised",
                "Revised to": "S2-2601002",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions (TEI18)",
                "Release": "Rel-18",
                "Title": "Original Arch proposal",
                "Source": "Deutsche Telekom",
                "_Sanitized_Companies": ["Deutsche Telekom"],
                "_All_Sanitized_Companies": ["Deutsche Telekom"],
            },
            # 5. Latest active draft pCR (supersedes S2-2601001) with Delegate Notes
            {
                "TDoc": "S2-2601002",
                "Type": "pCR",
                "TDoc Status": "Available",
                "Is revision of": "S2-2601001",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions (TEI18)",
                "Release": "Rel-18",
                "Title": "Updated Arch proposal (rev of S2-2601001)",
                "Source": "Deutsche Telekom",
                "My Status": "🟢 Support",
                "My Notes": "Key priority for our delegation",
                "_Sanitized_Companies": ["Deutsche Telekom"],
                "_All_Sanitized_Companies": ["Deutsche Telekom"],
            },
            # 6. Proposed Change Request (CR) under AI 6.2
            {
                "TDoc": "S2-2601003",
                "Type": "CR",
                "TDoc Status": "For discussion",
                "Related WIs": "TEI18",
                "Agenda Item": "6.2",
                "Agenda item description": "Protocol Procedures",
                "Release": "Rel-18",
                "Title": "Protocol Alignment CR",
                "Source": "Nokia",
                "_Sanitized_Companies": ["Nokia"],
                "_All_Sanitized_Companies": ["Nokia"],
            },
            # 7. Proposed Discussion paper under AI 5.1
            {
                "TDoc": "S2-2601004",
                "Type": "Discussion",
                "TDoc Status": "Available",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions (TEI18)",
                "Release": "Rel-18",
                "Title": "Evaluation of Solution 3",
                "Source": "Qualcomm",
                "_Sanitized_Companies": ["Qualcomm"],
                "_All_Sanitized_Companies": ["Qualcomm"],
            },
            # 8. Reserved TDoc (not uploaded yet)
            {
                "TDoc": "S2-2601005",
                "Type": "pCR",
                "TDoc Status": "Reserved",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions (TEI18)",
                "Release": "Rel-18",
                "Title": "Reserved by Samsung",
                "Source": "Samsung",
                "_Sanitized_Companies": ["Samsung"],
                "_All_Sanitized_Companies": ["Samsung"],
            },
            # 9. Explicitly Withdrawn contribution (must be skipped)
            {
                "TDoc": "S2-2601006",
                "Type": "Discussion",
                "TDoc Status": "Withdrawn",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions (TEI18)",
                "Release": "Rel-18",
                "Title": "Withdrawn paper",
                "Source": "ZTE",
                "_Sanitized_Companies": ["ZTE"],
                "_All_Sanitized_Companies": ["ZTE"],
            },
        ]

        self.followed_tokens = {"tei18"}
        self.wi_db_info = {
            "wi_by_exact_acronym": {
                "tei18": {"acronym": "TEI18", "code": "180001", "name": "Rel-18 Improvements", "release": "Rel-18"},
            },
            "wi_by_code": {},
            "wi_by_token": {},
            "followed_tokens": self.followed_tokens,
            "starred_tokens": set(),
            "tracked_tokens": self.followed_tokens,
        }

    def test_dialog_initialization_and_preselection(self):
        """Verify PreMeetingExportDialog pre-selects followed WIs and sets pre-meeting defaults."""
        dialog = PreMeetingExportDialog(
            parent=None,
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            followed_tokens=self.followed_tokens,
            starred_tokens=set(),
        )

        cfg = dialog.get_export_config()
        self.assertEqual(cfg["scope_mode"], "followed")
        self.assertTrue(cfg["include_disc"], "Discussion papers must be checked by default in Pre-Meeting!")
        self.assertTrue(cfg["include_pcr"])
        self.assertTrue(cfg["include_cr"])
        self.assertTrue(cfg["include_ls_in"])
        self.assertTrue(cfg["include_ls_out"])
        self.assertTrue(cfg["deduplicate_revisions"])
        self.assertTrue(cfg["include_reserved_in_index"])
        self.assertTrue(cfg["include_user_notes"])

    def test_pre_meeting_filtering_and_lineage_deduplication(self):
        """Verify exporter deduplicates earlier revisions and captures pre-meeting proposals."""
        exporter = PreMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/",
            revisions_url="",
            scope_mode="followed",
            followed_tokens=self.followed_tokens,
            deduplicate_revisions=True,
            include_reserved_in_index=True,
            include_disc=True,
            download_missing=False,
            force_reconvert=False,
        )

        tree = exporter._filter_and_group_tdocs(self.wi_db_info)
        self.assertIn("Rel-18", tree)
        self.assertIn("TEI18", tree["Rel-18"])

        tei18_cats = tree["Rel-18"]["TEI18"]["categories"]

        # 1. Deduplication: S2-2601001 (revised) should be superseded; only S2-2601002 must be in pCR
        pcr_ids = [d["tdoc_id"] for d in tei18_cats["pCR"]]
        self.assertNotIn("S2-2601001", pcr_ids, "Superseded predecessor S2-2601001 must not be exported standalone")
        self.assertIn("S2-2601002", pcr_ids, "Latest active revision S2-2601002 must be exported")

        # Verify predecessor provenance is tracked
        s2_doc = next(d for d in tei18_cats["pCR"] if d["tdoc_id"] == "S2-2601002")
        self.assertIn("S2-2601001", s2_doc.get("supersedes", []))
        self.assertEqual(s2_doc.get("my_status"), "🟢 Support")
        self.assertEqual(s2_doc.get("my_notes"), "Key priority for our delegation")

        # 2. Discussion papers are included
        disc_ids = [d["tdoc_id"] for d in tei18_cats["DISC"]]
        self.assertIn("S2-2601004", disc_ids)

        # 3. Withdrawn paper (S2-2601006) must be excluded
        self.assertNotIn("S2-2601006", disc_ids)

        # 4. Reserved TDoc (S2-2601005) is included when include_reserved_in_index=True
        self.assertIn("S2-2601005", pcr_ids)
        res_doc = next(d for d in tei18_cats["pCR"] if d["tdoc_id"] == "S2-2601005")
        self.assertTrue(res_doc.get("is_reserved"))

    def test_multi_candidate_ls_replies_and_wi_inheritance(self):
        """Verify 1-to-many candidate LS reply mapping and WI inheritance for draft reply LSs."""
        exporter = PreMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/",
            revisions_url="",
            scope_mode="followed",
            followed_tokens=self.followed_tokens,
            download_missing=False,
        )

        in_to_replies, out_to_in, _ = exporter._build_ls_reply_map()

        # S2-2601000 should have both S2-2601010 and S2-2601020 as candidate replies
        self.assertIn("S2-2601000", in_to_replies)
        replies = in_to_replies["S2-2601000"]
        self.assertIn("S2-2601010", replies)
        self.assertIn("S2-2601020", replies)

        # Candidate reply S2-2601010 points back to S2-2601000
        self.assertIn("S2-2601000", out_to_in.get("S2-2601010", []))
        self.assertIn("S2-2601000", out_to_in.get("S2-2601020", []))

        # Test WI inheritance: Both S2-2601010 and S2-2601020 had Related WIs=""
        # They should inherit TEI18 from S2-2601000 and appear in TEI18/LS_OUT
        tree = exporter._filter_and_group_tdocs(self.wi_db_info)
        ls_out_ids = [d["tdoc_id"] for d in tree["Rel-18"]["TEI18"]["categories"]["LS_OUT"]]
        self.assertIn("S2-2601010", ls_out_ids, "Candidate LS out S2-2601010 should inherit TEI18 from parent LS in")
        self.assertIn("S2-2601020", ls_out_ids, "Candidate LS out S2-2601020 should inherit TEI18 from parent LS in")

        # Competing sibling replies must be cross-linked
        rep_doc = next(d for d in tree["Rel-18"]["TEI18"]["categories"]["LS_OUT"] if d["tdoc_id"] == "S2-2601010")
        self.assertIn("S2-2601020", rep_doc.get("sibling_replies", []))

    def test_contributor_matrix_calculation(self):
        """Verify Contributor Activity Matrix calculation."""
        exporter = PreMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="",
            revisions_url="",
        )

        test_docs = [
            {"category": "pCR", "sanitized_companies": ["Ericsson"]},
            {"category": "pCR", "sanitized_companies": ["Ericsson", "Nokia"]},
            {"category": "CR", "sanitized_companies": ["Nokia"]},
            {"category": "DISC", "sanitized_companies": ["Qualcomm"]},
            {"category": "LS_OUT", "sanitized_companies": ["Huawei"]},
        ]

        table_md = exporter._render_contributor_matrix(test_docs, max_companies=10)
        self.assertIn("**Ericsson**", table_md)
        self.assertIn("**Nokia**", table_md)
        self.assertIn("**Qualcomm**", table_md)
        self.assertIn("**Huawei**", table_md)
        self.assertIn("| Contributor | Total Docs | pCR | CR | Discussion | LS out |", table_md)

    def test_company_filter(self):
        """Verify PreMeetingExporter respects the company focus filter."""
        exporter = PreMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            company_filter={"nokia"},
            download_missing=False,
        )

        tree = exporter._filter_and_group_tdocs(self.wi_db_info)
        all_ids = [
            d["tdoc_id"]
            for rel in tree.values()
            for wi in rel.values()
            for docs in wi["categories"].values()
            for d in docs
        ]

        # S2-2601003 is from Nokia
        self.assertIn("S2-2601003", all_ids)
        # S2-2601004 is from Qualcomm (should be excluded)
        self.assertNotIn("S2-2601004", all_ids)


if __name__ == "__main__":
    unittest.main()
