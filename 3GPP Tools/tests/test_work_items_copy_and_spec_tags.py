# --- File: tests/test_work_items_copy_and_spec_tags.py ---
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from PyQt5.QtWidgets import QApplication, QMenu

from core.ui.flow_layout import FlowLayout
from modules.specifications.ui.spec_annotations_dialog import (
    BatchTagDialog,
    EditSpecAnnotationsDialog,
)
from modules.work_items.ui.work_items_dialogs import (
    build_wi_copy_data,
    copy_wis_default,
    populate_wi_batch_copy_menu,
    populate_wi_copy_menu,
)


class _DummySpecDB:
    def __init__(self, unique_tags=None, assigned_tags=None):
        self._unique = unique_tags or [
            "AI",
            "CUPS",
            "data framework",
            "data transformation",
            "ISAC",
            "LCM",
            "ML",
            "MnS",
        ]
        self._assigned = assigned_tags or []

    def get_all_unique_tags(self):
        return list(self._unique)

    def get_spec_annotations(self, spec_number: str):
        return {"tags": list(self._assigned), "comments": "Test note", "updated_at": "2026-11-03"}

    def set_spec_annotations(self, spec_number: str, tags: list, comments: str):
        self._assigned = list(tags)
        return True


class TestWorkItemsCopyAndSpecTags(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_build_wi_copy_data_word_and_plain_formats(self):
        row = {
            "code": "1121123",
            "acronym": "5G_A-UPF",
            "name": "Study on 5G Advanced UPF",
            "wg_names": "SA2",
            "release": "Rel-19",
            "latest_wid": "SP-230456r01",
        }
        data = build_wi_copy_data(row)

        expected_wi_url = (
            "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123"
        )
        expected_wid_url = (
            "https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid=SP-230456"
        )

        self.assertEqual(data["wi_url"], expected_wi_url)
        self.assertEqual(data["wid_url"], expected_wid_url)
        self.assertEqual(data["acr_link_plain"], f"5G_A-UPF ({expected_wi_url})")
        self.assertEqual(data["acr_link_html"], f'<a href="{expected_wi_url}">5G_A-UPF</a>')
        self.assertEqual(data["wid_link_plain"], f"SP-230456r01 ({expected_wid_url})")
        self.assertEqual(data["wid_link_html"], f'<a href="{expected_wid_url}">SP-230456r01</a>')
        self.assertIn(expected_wi_url, data["cit_plain"])
        self.assertIn(f'<a href="{expected_wi_url}">', data["cit_html"])
        self.assertIn(expected_wid_url, data["cit_with_wid_html"])

    def test_single_and_batch_wi_copy_menus_and_clipboard(self):
        row1 = {
            "code": "1121123",
            "acronym": "5G_A-UPF",
            "name": "Study on 5G Advanced UPF",
            "wg_names": "SA2",
            "release": "Rel-19",
            "latest_wid": "SP-230456",
        }
        row2 = {
            "code": "980023",
            "acronym": "e_RedCap",
            "name": "Enhanced RedCap",
            "wg_names": "RAN1, RAN2",
            "release": "Rel-18",
            "latest_wid": "",
        }

        single_menu = QMenu()
        populate_wi_copy_menu(single_menu, row1)
        action_texts = [a.text() for a in single_menu.actions() if not a.isSeparator()]
        self.assertTrue(any("Copy WI Acronym with Link" in t for t in action_texts))
        self.assertTrue(any("Copy WI Portal URL Only" in t for t in action_texts))
        self.assertTrue(any("Copy Latest WID with Link" in t for t in action_texts))
        self.assertTrue(any("Copy Latest WID URL Only" in t for t in action_texts))

        batch_menu = QMenu()
        populate_wi_batch_copy_menu(batch_menu, [row1, row2])
        batch_texts = [a.text() for a in batch_menu.actions() if not a.isSeparator()]
        self.assertTrue(any("Copy Acronyms with Links (2)" in t for t in batch_texts))
        self.assertTrue(any("Copy WI Portal URLs List (2)" in t for t in batch_texts))
        self.assertTrue(any("Copy Latest WIDs with Links (1)" in t for t in batch_texts))

        copy_wis_default([row1])
        mime = QApplication.clipboard().mimeData()
        self.assertIn("5G_A-UPF", mime.text())
        self.assertIn("workitemId=1121123", mime.text())
        self.assertIn('<a href="https://portal.3gpp.org/', mime.html())

    def test_spec_annotations_dialog_flow_layout_and_unclipped_tags(self):
        db = _DummySpecDB()
        dlg = EditSpecAnnotationsDialog(
            "23.801-01", "Study on Architecture for 6G System; Stage 2", db
        )
        dlg.show()
        self.app.processEvents()

        fl = dlg.suggestions_layout
        self.assertIsInstance(fl, FlowLayout)
        self.assertEqual(fl.count(), 8)

        # Ensure long tags like '+ data framework' and '+ data transformation' are not clipped
        for i in range(fl.count()):
            btn = fl.itemAt(i).widget()
            text_width = btn.fontMetrics().horizontalAdvance(btn.text())
            self.assertGreater(
                btn.width(),
                text_width,
                f"Suggestion chip '{btn.text()}' was clipped (btn width={btn.width()}, text width={text_width})",
            )

        # Verify adding a suggested tag updates assigned pills and removes from suggestions
        dlg._add_suggested_tag("data framework")
        self.app.processEvents()
        self.assertIn("data framework", dlg.current_tags)
        self.assertEqual(dlg.suggestions_layout.count(), 7)
        dlg.close()

    def test_batch_tag_dialog_uses_flow_layout(self):
        db = _DummySpecDB()
        dlg = BatchTagDialog(["23.501", "23.502"], db)
        dlg.show()
        self.app.processEvents()
        dlg.close()

    def test_work_items_tab_phase2_ribbon_icons_and_toast(self):
        from modules.work_items.ui.ui_tabs import WorkItemsTab

        db_path = ROOT / "src" / "db" / "3gpp_data.db"
        tab = WorkItemsTab(db_path)
        self.assertTrue(hasattr(tab, "ribbon_card"))
        self.assertTrue(hasattr(tab, "table_card"))
        for btn in (tab.add_wi_btn, tab.sync_tracked_btn, tab.sync_btn, tab.clear_filters_btn):
            self.assertFalse(btn.icon().isNull())
        # Triggering sync tracked with empty starred/following must use non-blocking toast helper
        orig_starred, orig_following = set(tab.starred), set(tab.following)
        tab.starred.clear()
        tab.following.clear()
        try:
            tab._sync_tracked_wis()
        finally:
            tab.starred = orig_starred
            tab.following = orig_following
        tab.close()


if __name__ == "__main__":
    unittest.main()
