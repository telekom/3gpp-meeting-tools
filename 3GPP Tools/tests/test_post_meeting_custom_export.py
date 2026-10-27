# --- File: tests/test_post_meeting_custom_export.py ---
import sys
import unittest
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.meetings.core.post_meeting_exporter import PostMeetingExporterThread
from modules.meetings.ui.post_meeting_export_dialog import PostMeetingExportDialog


class TestPostMeetingCustomExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        # Sample dataset featuring WIs spanning multiple AIs and discussion papers
        self.rows = [
            # TEI18 under AI 5.1 (CR, agreed)
            {
                "TDoc": "S2-2601001",
                "Type": "CR",
                "TDoc Status": "agreed",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "High-Level Architecture (TEI18)",
                "Release": "Rel-18",
                "Title": "TEI18 Arch CR",
            },
            # TEI18 under AI 6.2 (CR, agreed)
            {
                "TDoc": "S2-2601002",
                "Type": "CR",
                "TDoc Status": "agreed",
                "Related WIs": "TEI18",
                "Agenda Item": "6.2",
                "Agenda item description": "Protocols & Procedures (TEI18)",
                "Release": "Rel-18",
                "Title": "TEI18 Protocol CR",
            },
            # FS_eLCS_Ph3 under AI 5.1 (pCR, agreed)
            {
                "TDoc": "S2-2601003",
                "Type": "pCR",
                "TDoc Status": "agreed",
                "Related WIs": "FS_eLCS_Ph3",
                "Agenda Item": "5.1",
                "Agenda item description": "High-Level Architecture",
                "Release": "Rel-19",
                "Title": "eLCS Study Flow pCR",
            },
            # Discussion paper under AI 5.1 (Discussion, noted)
            {
                "TDoc": "S2-2601004",
                "Type": "Discussion",
                "TDoc Status": "noted",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "High-Level Architecture (TEI18)",
                "Release": "Rel-18",
                "Title": "Discussion on TEI18 Architecture",
            },
            # Work Item Description (WID) under AI 3.1 (WID, agreed)
            {
                "TDoc": "S2-2601005",
                "Type": "WID",
                "TDoc Status": "agreed",
                "Related WIs": "TEI18",
                "Agenda Item": "3.1",
                "Agenda item description": "Work Item Status Reports",
                "Release": "Rel-18",
                "Title": "Updated WID for TEI18",
            },
            # General LS IN under AI 4.1 (LS in, no WI)
            {
                "TDoc": "S2-2601006",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "",
                "Agenda Item": "4.1",
                "Agenda item description": "Incoming LSs",
                "Release": "Rel-19",
                "Title": "General Incoming LS",
            },
        ]

        self.followed_tokens = {"tei18"}
        self.wi_db_info = {
            "wi_by_exact_acronym": {
                "tei18": {"acronym": "TEI18", "code": "180001", "name": "Rel-18 Improvements", "release": "Rel-18"},
                "fs_elcs_ph3": {"acronym": "FS_eLCS_Ph3", "code": "190002", "name": "Location Services Ph3", "release": "Rel-19"},
            },
            "wi_by_code": {},
            "wi_by_token": {},
            "followed_tokens": self.followed_tokens,
            "starred_tokens": set(),
            "tracked_tokens": self.followed_tokens,
        }

    def test_dialog_auto_preselects_followed_wis(self):
        """Verify PostMeetingExportDialog pre-checks followed WIs and their AIs automatically."""
        dialog = PostMeetingExportDialog(
            parent=None,
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            followed_tokens=self.followed_tokens,
            starred_tokens=set(),
        )

        # Tracked WI 'TEI18' has AIs 5.1, 6.2, and 3.1
        self.assertIn(("TEI18", "5.1"), dialog.selected_wi_ais)
        self.assertIn(("TEI18", "6.2"), dialog.selected_wi_ais)
        self.assertIn(("TEI18", "3.1"), dialog.selected_wi_ais)

        # Non-tracked WI 'FS_eLCS_Ph3' must NOT be in default selection
        self.assertNotIn(("FS_eLCS_Ph3", "5.1"), dialog.selected_wi_ais)

        cfg = dialog.get_export_config()
        self.assertIn("tei18", cfg["custom_wis"])
        self.assertNotIn("fs_elcs_ph3", cfg["custom_wis"])
        self.assertIn("5.1", cfg["custom_wi_ais"]["tei18"])
        self.assertIn("6.2", cfg["custom_wi_ais"]["tei18"])

    def test_dialog_dual_tree_view_synchronization(self):
        """Verify switching between 'By Work Item' and 'By Agenda Item' keeps selections synchronized."""
        dialog = PostMeetingExportDialog(
            parent=None,
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            followed_tokens=set(),  # All selected initially
            starred_tokens=set(),
        )

        # Uncheck AI 6.2 under TEI18
        dialog.selected_wi_ais.discard(("TEI18", "6.2"))

        # Switch to Agenda Item view
        dialog._switch_custom_view("ai")
        self.assertEqual(dialog.tree_stack.currentWidget(), dialog.tree_ai)

        # In AI view, AI 6.2 should show TEI18 unchecked
        root = dialog.tree_ai.invisibleRootItem()
        found_62 = False
        for i in range(root.childCount()):
            ai_item = root.child(i)
            if "AI 6.2" in ai_item.text(0):
                found_62 = True
                self.assertEqual(ai_item.checkState(0), Qt.Unchecked)
                for j in range(ai_item.childCount()):
                    child = ai_item.child(j)
                    self.assertEqual(child.checkState(0), Qt.Unchecked)
        self.assertTrue(found_62)

        # Switch back to Work Item view
        dialog._switch_custom_view("wi")
        self.assertEqual(dialog.tree_stack.currentWidget(), dialog.tree_wi)

        # TEI18 should now be PartiallyChecked (5.1 and 3.1 checked, 6.2 unchecked)
        root_wi = dialog.tree_wi.invisibleRootItem()
        for i in range(root_wi.childCount()):
            wi_item = root_wi.child(i)
            if "TEI18" in wi_item.text(0):
                self.assertEqual(wi_item.checkState(0), Qt.PartiallyChecked)

    def test_exporter_fine_grained_wi_and_ai_scoping(self):
        """Verify that selecting only specific AIs under a WI exports ONLY those AIs and filters others."""
        # Export only AI 5.1 under TEI18 (exclude AI 6.2)
        thread = PostMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"tei18"},
            custom_wi_ais={"tei18": {"5.1"}},
            include_cr=True,
            include_pcr=True,
            include_ls_in=False,
            include_general_ls=False,
        )

        tree = thread._filter_and_group_tdocs(self.wi_db_info)

        exported_tdocs = set()
        for rel_data in tree.values():
            for wi_data in rel_data.values():
                for doc_list in wi_data["categories"].values():
                    for d in doc_list:
                        exported_tdocs.add(d["tdoc_id"])

        # S2-2601001 (TEI18 in AI 5.1) MUST be included
        self.assertIn("S2-2601001", exported_tdocs)

        # S2-2601002 (TEI18 in AI 6.2) MUST BE EXCLUDED because only AI 5.1 was chosen!
        self.assertNotIn("S2-2601002", exported_tdocs)

        # S2-2601003 (eLCS) was not in custom_wis, so excluded
        self.assertNotIn("S2-2601003", exported_tdocs)

    def test_exporter_discussion_papers_inclusion(self):
        """Verify that discussion papers and WID status reports are exported when include_disc=True."""
        thread = PostMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            include_cr=True,
            include_pcr=True,
            include_ls_in=True,
            include_disc=True,  # Discussion papers ENABLED
            include_general_ls=True,
        )

        tree = thread._filter_and_group_tdocs(self.wi_db_info)

        exported_cats = {}
        for rel_data in tree.values():
            for wi_data in rel_data.values():
                for cat, doc_list in wi_data["categories"].items():
                    for d in doc_list:
                        exported_cats[d["tdoc_id"]] = cat

        # S2-2601004 (Discussion, noted) and S2-2601005 (WID, agreed) MUST be categorized as DISC
        self.assertIn("S2-2601004", exported_cats)
        self.assertEqual(exported_cats["S2-2601004"], "DISC")

        self.assertIn("S2-2601005", exported_cats)
        self.assertEqual(exported_cats["S2-2601005"], "DISC")

    def test_exporter_discussion_papers_excluded_by_default(self):
        """Verify that discussion papers are excluded by default when include_disc=False."""
        thread = PostMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "178"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            include_cr=True,
            include_pcr=True,
            include_ls_in=True,
            include_disc=False,  # Discussion papers DISABLED
        )

        tree = thread._filter_and_group_tdocs(self.wi_db_info)

        exported_ids = set()
        for rel_data in tree.values():
            for wi_data in rel_data.values():
                for doc_list in wi_data["categories"].values():
                    for d in doc_list:
                        exported_ids.add(d["tdoc_id"])

        self.assertNotIn("S2-2601004", exported_ids)
        self.assertNotIn("S2-2601005", exported_ids)


if __name__ == "__main__":
    unittest.main()
