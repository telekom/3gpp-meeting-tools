import sys
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

from modules.meetings.core.post_meeting_exporter import PostMeetingExporterThread
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.ui.post_meeting_export_dialog import PostMeetingExportDialog
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel


class TestUnassignedLsInExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.rows = [
            # 1. LS IN with NO WI assigned and generic/unrelated AI description
            {
                "TDoc": "S2-2609300",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "",
                "Agenda Item": "19.2.2",
                "Agenda item description": "System architecture for next generation real time communication (NG_RTC_Ph2)",
                "Release": "Rel-19",
                "Title": "LS on NG-RTC",
            },
            # 2. LS IN with NO WI assigned and AI description containing parenthetical CR(s)
            {
                "TDoc": "S2-2609338",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "",
                "Agenda Item": "20.50",
                "Agenda item description": "Rel-20 LSs and CAT B/C alignment CR(s) due to the study",
                "Release": "Rel-20",
                "Title": "LS on alignment",
            },
            # 3. LS IN with NO WI assigned and common issues AI 4.1
            {
                "TDoc": "S2-2609305",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "",
                "Agenda Item": "4.1",
                "Agenda item description": "Common issues and Incoming LSs",
                "Release": "Rel-20",
                "Title": "General Incoming LS",
            },
            # 4. LS IN with NO WI assigned but AI description matches followed WI
            {
                "TDoc": "S2-2609303",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "",
                "Agenda Item": "20.6.20",
                "Agenda item description": "Study on Architecture for 6G System (FS_6G_ARC)",
                "Release": "Rel-20",
                "Title": "LS on 6G Arch",
            },
            # 5. LS IN with an ASSIGNED non-followed WI
            {
                "TDoc": "S2-2609399",
                "Type": "LS in",
                "TDoc Status": "",
                "Related WIs": "AmbientIoT_Ph2",
                "Agenda Item": "20.5.2",
                "Agenda item description": "Ambient IoT",
                "Release": "Rel-20",
                "Title": "Assigned LS to Ambient IoT",
            },
            # 6. Followed WI CR (Agreed)
            {
                "TDoc": "S2-2609400",
                "Type": "CR",
                "TDoc Status": "Agreed",
                "Related WIs": "6G_ARC",
                "Agenda Item": "20.6.20",
                "Agenda item description": "Study on Architecture for 6G System (FS_6G_ARC)",
                "Release": "Rel-20",
                "Title": "Agreed CR on 6G",
            },
            # 7. Non-followed WI CR (Agreed)
            {
                "TDoc": "S2-2609401",
                "Type": "CR",
                "TDoc Status": "Agreed",
                "Related WIs": "AmbientIoT_Ph2",
                "Agenda Item": "20.5.2",
                "Agenda item description": "Ambient IoT",
                "Release": "Rel-20",
                "Title": "Agreed CR on Ambient IoT",
            },
        ]
        # User only follows 6G_ARC
        self.followed_tokens = {"6g_arc", "fs_6g_arc"}
        self.wi_db_info = {
            "wi_by_exact_acronym": {
                "fs_6g_arc": {"acronym": "FS_6G_ARC", "code": "1080058", "name": "6G Architecture", "release": "Rel-20"},
                "6g_arc": {"acronym": "6G_ARC", "code": "1080058", "name": "6G Architecture", "release": "Rel-20"},
                "ng_rtc_ph2": {"acronym": "NG_RTC_Ph2", "code": "9999", "name": "NG RTC", "release": "Rel-19"},
                "ambientiot_ph2": {"acronym": "AmbientIoT_Ph2", "code": "8888", "name": "Ambient IoT", "release": "Rel-20"},
            },
            "wi_by_code": {},
            "wi_by_token": {},
            "followed_tokens": self.followed_tokens,
            "starred_tokens": set(),
            "tracked_tokens": self.followed_tokens,
        }

    def test_post_meeting_exporter_includes_unassigned_ls_in_by_default(self):
        thread = PostMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "177"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="followed",
            followed_tokens=self.followed_tokens,
            include_ls_in=True,
            include_general_ls=True,
        )

        tree = thread._filter_and_group_tdocs(self.wi_db_info)

        # Collect all exported TDoc IDs by category and WI
        exported_tdocs = {}
        for rel, wis in tree.items():
            for wi, data in wis.items():
                for cat, doc_list in data["categories"].items():
                    for d in doc_list:
                        exported_tdocs[d["tdoc_id"]] = (rel, wi, cat)

        # 1. S2-2609300 (LS IN, no WI in Related WIs, but NG_RTC_Ph2 in AIDesc)
        # MUST be included under General/LS_IN because it has no official WI and NG_RTC_Ph2 is not followed!
        self.assertIn("S2-2609300", exported_tdocs)
        self.assertEqual(exported_tdocs["S2-2609300"][1], "General")
        self.assertEqual(exported_tdocs["S2-2609300"][2], "LS_IN")

        # 2. S2-2609338 (LS IN, no WI, CR(s) in AIDesc)
        # MUST be included under General/LS_IN!
        self.assertIn("S2-2609338", exported_tdocs)
        self.assertEqual(exported_tdocs["S2-2609338"][1], "General")
        self.assertEqual(exported_tdocs["S2-2609338"][2], "LS_IN")

        # 3. S2-2609305 (AI 4.1 unassigned LS IN)
        # MUST be included under General/LS_IN!
        self.assertIn("S2-2609305", exported_tdocs)
        self.assertEqual(exported_tdocs["S2-2609305"][1], "General")

        # 4. S2-2609303 (LS IN matching followed WI 6G_ARC)
        # MUST be included under FS_6G_ARC!
        self.assertIn("S2-2609303", exported_tdocs)
        self.assertEqual(exported_tdocs["S2-2609303"][1], "FS_6G_ARC")

        # 5. S2-2609399 (LS IN with an ASSIGNED non-followed WI)
        # Should be excluded because it has a specific assigned WI not followed by the user
        self.assertNotIn("S2-2609399", exported_tdocs)

        # 6. S2-2609400 (Agreed CR on followed WI)
        # Included under 6G_ARC
        self.assertIn("S2-2609400", exported_tdocs)
        self.assertEqual(exported_tdocs["S2-2609400"][1], "6G_ARC")

        # 7. S2-2609401 (Agreed CR on non-followed WI)
        # Excluded
        self.assertNotIn("S2-2609401", exported_tdocs)

    def test_post_meeting_exporter_excludes_unassigned_ls_in_when_disabled(self):
        thread = PostMeetingExporterThread(
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "177"},
            docs_ftp_url="",
            revisions_url="",
            scope_mode="followed",
            followed_tokens=self.followed_tokens,
            include_ls_in=True,
            include_general_ls=False,  # Disabled
        )

        tree = thread._filter_and_group_tdocs(self.wi_db_info)
        exported_ids = set()
        for rel, wis in tree.items():
            for wi, data in wis.items():
                for cat, doc_list in data["categories"].items():
                    for d in doc_list:
                        exported_ids.add(d["tdoc_id"])

        # Unassigned LS INs not matching followed WIs must now be excluded
        self.assertNotIn("S2-2609300", exported_ids)
        self.assertNotIn("S2-2609338", exported_ids)
        self.assertNotIn("S2-2609305", exported_ids)

        # But LS IN matching followed WI must still be included
        self.assertIn("S2-2609303", exported_ids)
        self.assertIn("S2-2609400", exported_ids)

    def test_post_meeting_export_dialog_defaults(self):
        dialog = PostMeetingExportDialog(
            parent=None,
            meeting_dir=Path("."),
            all_rows=self.rows,
            visible_rows=self.rows,
            agenda_map={},
            mtg_info={"wg_name": "SA2", "meeting_number": "177"},
            followed_tokens=self.followed_tokens,
            starred_tokens=set(),
        )
        cfg = dialog.get_export_config()
        self.assertTrue(cfg["include_ls_in"])
        self.assertTrue(cfg["include_general_ls"])
        self.assertEqual(cfg["scope_mode"], "followed")

    def test_proxy_model_tracked_wi_filtering_with_unassigned_ls_in(self):
        model = TDocsTableModel(Path("."), self.rows)
        proxy = TDocsFilterProxyModel()
        proxy.setSourceModel(model)

        # Initialize filter sets
        proxy.setTypeFilters({str(r.get("Type", "")).strip() for r in self.rows})
        proxy.setAIFilters({str(r.get("Agenda Item", "")).strip() for r in self.rows})
        proxy.setWIFilters(set().union(*[model.get_related_wis(i) for i in range(len(self.rows))]))
        proxy.setStatusFilters({str(r.get("TDoc Status", "")).strip() for r in self.rows})
        proxy.setMyStatusFilters({str(r.get("My Status", "")).strip() for r in self.rows})
        proxy.setCompanyFilters(set().union(*[r.get("_Sanitized_Companies", ["Other"]) for r in self.rows]))

        proxy.setTrackedWITokens(self.followed_tokens, set())
        proxy.setTrackedWIFilters(True, False)

        # Default: include_unassigned_ls_in is True
        self.assertTrue(proxy.include_unassigned_ls_in)

        visible_ids = set()
        for r in range(proxy.rowCount()):
            src_idx = proxy.mapToSource(proxy.index(r, 0))
            visible_ids.add(model._data[src_idx.row()]["TDoc"])

        # Unassigned LS INs must be visible
        self.assertIn("S2-2609300", visible_ids)
        self.assertIn("S2-2609338", visible_ids)
        self.assertIn("S2-2609305", visible_ids)
        self.assertIn("S2-2609303", visible_ids)
        # Followed CR must be visible
        self.assertIn("S2-2609400", visible_ids)

        # Non-followed assigned items must NOT be visible
        self.assertNotIn("S2-2609399", visible_ids)
        self.assertNotIn("S2-2609401", visible_ids)

        # When toggled off:
        proxy.setIncludeUnassignedLsIn(False)
        visible_ids_off = set()
        for r in range(proxy.rowCount()):
            src_idx = proxy.mapToSource(proxy.index(r, 0))
            visible_ids_off.add(model._data[src_idx.row()]["TDoc"])

        self.assertNotIn("S2-2609300", visible_ids_off)
        self.assertNotIn("S2-2609338", visible_ids_off)
        self.assertNotIn("S2-2609305", visible_ids_off)
        self.assertIn("S2-2609303", visible_ids_off)
        self.assertIn("S2-2609400", visible_ids_off)

    def test_settings_include_unassigned_ls_in_persistence(self):
        settings = MeetingsSettings()
        initial = settings.get_include_unassigned_ls_in()
        self.assertIsInstance(initial, bool)

        # Toggle and verify persistence
        settings.set_include_unassigned_ls_in(False)
        self.assertFalse(settings.get_include_unassigned_ls_in())

        settings.set_include_unassigned_ls_in(True)
        self.assertTrue(settings.get_include_unassigned_ls_in())

    def test_post_meeting_exporter_plain_document_handling(self):
        """Verify that PostMeetingExporter renders LS IN documents in accept mode without change marks."""
        import tempfile, shutil
        from modules.meetings.core.tdoc_markdown_converter import resolve_revision_mode_for_type

        tmp_dir = Path(tempfile.mkdtemp(prefix="test_pm_plain_"))
        try:
            meeting_dir = tmp_dir / "Meeting"
            meeting_dir.mkdir(parents=True, exist_ok=True)

            exporter = PostMeetingExporterThread(
                meeting_dir=meeting_dir,
                all_rows=[],
                visible_rows=[],
                agenda_map={},
                mtg_info={},
                docs_ftp_url="",
                revisions_url="",
                scope_mode="whole",
                output_dir=tmp_dir / "Export",
            )

            # 1. Mode resolution check
            self.assertEqual(resolve_revision_mode_for_type("LS in"), "accept")
            self.assertEqual(resolve_revision_mode_for_type("LS out"), "accept")
            self.assertEqual(resolve_revision_mode_for_type("CR"), "markup")
            self.assertEqual(resolve_revision_mode_for_type("pCR"), "markup")

            # 2. Test docx conversion via _ensure_tdoc_markdown
            import docx
            tdoc_folder = meeting_dir / "S2-2609001"
            tdoc_folder.mkdir(parents=True, exist_ok=True)
            doc = docx.Document()
            doc.add_heading("Incoming Liaison", level=1)
            doc.add_paragraph("This is a clean communication from another WG.")
            docx_path = tdoc_folder / "S2-2609001.docx"
            doc.save(str(docx_path))

            md_path, _ = exporter._ensure_tdoc_markdown(
                "S2-2609001",
                {"TDoc": "S2-2609001", "Type": "LS in", "Title": "Incoming Liaison"}
            )
            self.assertTrue(md_path.exists())
            text = md_path.read_text(encoding="utf-8")
            self.assertIn("mode=accept", text)
            self.assertIn("(revisions: accept)", text)
            self.assertNotIn("Markup legend:", text)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
