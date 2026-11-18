import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMenu

from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_menus import build_related_menu, build_row_context_menu
from modules.meetings.ui.tdocs_window import TDocsWindow


class TestTDocsQuickOpenEmailSendAndFilterPreserve(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.tmp_path = Path(tempfile.mkdtemp(prefix="test_tdocs_qo_"))
        self.meeting_dir = self.tmp_path / "SA2_178"
        self.agenda_dir = self.meeting_dir / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)
        self.agenda_file = self.agenda_dir / "TDocs.xlsx"
        self.agenda_file.touch()

        self.mtg_info = {
            "wg_name": "SA2",
            "meeting_number": "178",
            "is_electronic": 0,
            "is_active_sync": False,
        }
        self.sample_data = [
            {
                "TDoc": "S2-2611482",
                "Title": "New SID: Study on satellite access Phase 5",
                "Source": "CATT",
                "Type": "SID new",
                "For": "Approval",
                "Agenda Item": "30.2",
                "TDoc Status": "available",
            },
            {
                "TDoc": "S2-2610535",
                "Title": "Earlier satellite study draft",
                "Source": "Nokia",
                "Type": "Discussion",
                "For": "Discussion",
                "Agenda Item": "30.1",
                "TDoc Status": "revised",
            },
            {
                "TDoc": "S2-2612000",
                "Title": "Architecture alignment CR",
                "Source": "Ericsson",
                "Type": "CR",
                "For": "Approval",
                "Agenda Item": "12.1",
                "TDoc Status": "available",
            },
        ]

    def tearDown(self):
        shutil.rmtree(self.tmp_path, ignore_errors=True)

    def _make_window(self):
        with patch("modules.meetings.ui.tdocs_window.MeetingWIEnricherThread"):
            win = TDocsWindow(self.mtg_info, list(self.sample_data), str(self.agenda_file))
        win.docs_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/"
        win.main_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/"
        return win

    def test_quick_open_sync_when_tdoc_in_current_filter(self):
        win = self._make_window()
        win.type_combo.setCheckedItems(["SID new", "Discussion"])
        win._on_type_changed(win.type_combo.getCheckedItems())
        self.assertEqual(win.proxy.rowCount(), 2)

        thread = MagicMock()
        thread.sync_table_on_success = True
        thread.is_silent_compare = False
        thread.send_email_on_finish = False
        thread.target_filename = "S2-2611482r01"

        with patch.object(win.table, "scrollTo") as mock_scroll:
            win._on_tdoc_action_finished("S2-2611482", True, "Opened successfully.", thread)

        # Current filter preserved, row highlighted and scrolled
        self.assertEqual(set(win.type_combo.getCheckedItems()), {"SID new", "Discussion"})
        self.assertEqual(win.search_input.text(), "")
        self.assertEqual(win.model._highlighted_tdoc, "S2-2611482")
        selected_rows = {idx.row() for idx in win.table.selectionModel().selectedIndexes()}
        self.assertEqual(len(selected_rows), 1)
        sel_row = next(iter(selected_rows))
        self.assertEqual(win.proxy.data(win.proxy.index(sel_row, 1), Qt.UserRole), "S2-2611482")
        mock_scroll.assert_called_once()
        win.close()

    def test_quick_open_sync_clears_and_filters_when_hidden(self):
        win = self._make_window()
        win.company_combo.setCheckedItems(["Ericsson"])
        win._on_company_changed(["Ericsson"])
        win.search_input.setText("Architecture")
        win._apply_search_filter()
        self.assertEqual(win.proxy.rowCount(), 1)
        self.assertEqual(win.proxy.data(win.proxy.index(0, 1), Qt.UserRole), "S2-2612000")

        thread = MagicMock()
        thread.sync_table_on_success = True
        thread.is_silent_compare = False
        thread.send_email_on_finish = False
        thread.target_filename = "S2-2611482"

        win._on_tdoc_action_finished("S2-2611482", True, "Opened successfully.", thread)

        # Filters cleared and search set to reveal S2-2611482
        self.assertTrue(win.company_combo.isAllChecked())
        self.assertEqual(win.search_input.text(), "S2-2611482")
        self.assertEqual(win.proxy.rowCount(), 1)
        self.assertEqual(win.proxy.data(win.proxy.index(0, 1), Qt.UserRole), "S2-2611482")
        self.assertEqual(win.model._highlighted_tdoc, "S2-2611482")
        selected_rows = {idx.row() for idx in win.table.selectionModel().selectedIndexes()}
        self.assertEqual(selected_rows, {0})
        win.close()

    def test_quick_open_failure_does_not_change_filters(self):
        win = self._make_window()
        win.company_combo.setCheckedItems(["Ericsson"])
        win._on_company_changed(["Ericsson"])

        thread = MagicMock()
        thread.sync_table_on_success = True
        thread.target_filename = "S2-2611482"

        with patch("modules.meetings.ui.tdocs_window.QMessageBox.warning"):
            win._on_tdoc_action_finished("S2-2611482", False, "404 Not Found", thread)

        self.assertEqual(win.company_combo.getCheckedItems(), ["Ericsson"])
        self.assertIsNone(win.model._highlighted_tdoc)
        win.close()

    def test_instant_fetch_tags_sync_table_on_success(self):
        win = self._make_window()
        win.instant_fetch_input.setText("S2-2611482r01")
        with patch.object(win, "_trigger_download_thread") as mock_dl:
            win._on_instant_fetch()
        mock_dl.assert_called_once_with(
            "S2-2611482",
            "S2-2611482R01",
            legacy_url=None,
            is_silent_compare=False,
            sync_table_on_success=True,
        )
        win.close()

    def test_find_local_tdoc_files_extracts_zip_and_respects_boundaries(self):
        tdoc_dir = self.meeting_dir / "S2-2611482"
        tdoc_dir.mkdir(parents=True, exist_ok=True)
        # Revision file should NOT be matched when querying base TDoc
        (tdoc_dir / "S2-2611482r01.docx").write_bytes(b"rev")
        self.assertEqual(TDocFileHandler.find_local_tdoc_files(self.meeting_dir, "S2-2611482"), [])

        # Base ZIP present -> extracts base doc cleanly
        zip_path = tdoc_dir / "S2-2611482.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("S2-2611482.docx", b"base-doc-content")

        found = TDocFileHandler.find_local_tdoc_files(self.meeting_dir, "S2-2611482")
        self.assertEqual([p.name for p in found], ["S2-2611482.docx"])
        self.assertTrue(found[0].is_file())

    def test_send_tdoc_via_email_uses_cached_file_and_subject(self):
        win = self._make_window()
        tdoc_dir = self.meeting_dir / "S2-2611482"
        tdoc_dir.mkdir(parents=True, exist_ok=True)
        doc_path = tdoc_dir / "S2-2611482.docx"
        doc_path.write_bytes(b"dummy")

        with patch(
            "modules.meetings.ui.tdocs_window.OutlookClient.create_email_draft",
            return_value=(True, ""),
        ) as mock_draft:
            win._send_tdoc_via_email("S2-2611482")

        mock_draft.assert_called_once()
        subj, attachments = mock_draft.call_args[0]
        self.assertEqual(subj, "[SA2#178, 30.2] S2-2611482 New SID: Study on satellite access Phase 5")
        self.assertEqual([Path(p) for p in attachments], [doc_path])
        win.close()

    def test_send_tdoc_via_email_downloads_first_when_uncached(self):
        win = self._make_window()
        with patch.object(win, "_trigger_download_thread") as mock_dl:
            win._send_tdoc_via_email("S2-2611482r01")

        mock_dl.assert_called_once_with(
            "S2-2611482",
            "S2-2611482r01",
            send_email_on_finish=True,
        )

        # Simulate thread completion
        tdoc_dir = self.meeting_dir / "S2-2611482"
        tdoc_dir.mkdir(parents=True, exist_ok=True)
        rev_doc = tdoc_dir / "S2-2611482r01.docx"
        rev_doc.write_bytes(b"rev")

        thread = MagicMock()
        thread.sync_table_on_success = False
        thread.send_email_on_finish = True
        thread.is_silent_compare = True
        thread.target_filename = "S2-2611482r01"
        thread.extracted_doc_paths = [rev_doc]

        with patch(
            "modules.meetings.ui.tdocs_window.OutlookClient.create_email_draft",
            return_value=(True, ""),
        ) as mock_draft:
            win._on_tdoc_action_finished("S2-2611482", True, "Downloaded", thread)

        mock_draft.assert_called_once()
        subj, attachments = mock_draft.call_args[0]
        self.assertIn("S2-2611482r01", subj)
        self.assertEqual([Path(p) for p in attachments], [rev_doc])
        win.close()

    def test_email_menus_include_send_tdoc_actions(self):
        captured = {}

        def _capture_exec(menu_self, *args, **kwargs):
            captured["menu"] = menu_self
            return None

        sent = []
        with patch.object(QMenu, "exec_", _capture_exec):
            build_row_context_menu(
                parent=None,
                row_data={"TDoc": "S2-2611482", "Title": "Test"},
                revisions=["S2-2611482r01", "S2-2611482r02"],
                family=["S2-2611482", "S2-2611482r01", "S2-2611482r02"],
                docs_ftp_url="",
                revisions_url="",
                meeting_dir=self.meeting_dir,
                download_callback=lambda *a, **kw: None,
                export_llm_callback=lambda *a, **kw: None,
                compose_email_callback=lambda *a, **kw: None,
                open_emails_callback=lambda *a, **kw: None,
                mark_read_callback=lambda: None,
                mark_unread_callback=lambda: None,
                open_details_callback=lambda: None,
                send_tdoc_email_callback=lambda t: sent.append(t),
            )

        emails_act = next(a for a in captured["menu"].actions() if "Emails" in a.text())
        emails_menu = emails_act.menu()
        send_latest_act = next(a for a in emails_menu.actions() if "Send Latest (S2-2611482r02)" in a.text())
        send_latest_act.trigger()
        self.assertEqual(sent, ["S2-2611482r02"])

        with patch.object(QMenu, "exec_", _capture_exec):
            build_related_menu(
                parent=None,
                target_tdoc="S2-2610535",
                valid_tdocs={"S2-2610535"},
                docs_ftp_url="",
                revisions_url="",
                scroll_callback=lambda *a: None,
                download_callback=lambda *a, **kw: None,
                export_llm_callback=lambda *a: None,
                global_search_callback=lambda *a: None,
                compose_email_callback=lambda *a: None,
                pos=None,
                send_tdoc_email_callback=lambda t: sent.append(t),
            )
        send_rel_act = next(a for a in captured["menu"].actions() if "Send S2-2610535 via Email" in a.text())
        send_rel_act.trigger()
        self.assertEqual(sent, ["S2-2611482r02", "S2-2610535"])

    def test_checkable_combobox_update_items_preserves_filter_and_save_user_data_keeps_filters(self):
        cb = CheckableComboBox("Company")
        cb.addItems(["CATT", "Ericsson", "Nokia"])
        cb.setCheckedItems(["CATT"])
        self.assertEqual(cb.getCheckedItems(), ["CATT"])

        cb.updateItems(["CATT", "Ericsson", "Huawei", "Nokia"])
        self.assertEqual(cb.getCheckedItems(), ["CATT"])
        self.assertTrue(cb.isFiltered())

        win = self._make_window()
        win.company_combo.setCheckedItems(["CATT"])
        win._on_company_changed(["CATT"])
        win.search_input.setText("satellite")
        win._apply_search_filter()
        self.assertEqual(win.proxy.rowCount(), 1)

        win._save_user_data("S2-2611482", "🟢 Interesting", "Important satellite SID")

        self.assertEqual(win.company_combo.getCheckedItems(), ["CATT"])
        self.assertTrue(win.company_combo.isFiltered())
        self.assertEqual(win.search_input.text(), "satellite")
        self.assertEqual(win.proxy.rowCount(), 1)
        self.assertEqual(win.proxy.data(win.proxy.index(0, 1), Qt.UserRole), "S2-2611482")
        win.close()


if __name__ == "__main__":
    unittest.main()
