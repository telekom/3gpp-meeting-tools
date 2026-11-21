import sys
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from modules.meetings.ui.tdocs_models import TDocsTableModel


class TestTDocsModelsMergeAgenda(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.meeting_dir = Path(self.temp_dir.name)

        initial_data = [
            {
                'TDoc': 'S2-260123',
                'Title': 'Initial Base Document',
                'Source': 'Nokia',
                'Type': 'CR',
                'For': 'Agreement',
                'Abstract': 'Base Abstract',
                'Secretary Remarks': '',
                'Agenda Item': '5.2',
                'TDoc Status': 'Not Handled',
                'Is revision of': '',
                'Revised to': '',
                'Original LS': '',
                'Reply in': '',
                'Related WIs': 'FS_eVoNR, 5G_Adv',
            }
        ]
        self.model = TDocsTableModel(meeting_dir=self.meeting_dir, data=initial_data)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_merge_agenda_standalone_on_the_fly_tdoc(self):
        """
        Tests injecting a brand-new on-the-fly TDoc without a revision suffix (like S2-2609298).
        Ensures UnboundLocalError for 'fuzzy_get' is resolved, and the document is cleanly added.
        """
        agenda_data = {
            'S2-2609298': {
                'Title': 'Late Submitted Discussion Paper',
                'Source': 'Ericsson',
                'Type': 'Discussion',
                'For': 'Discussion',
                'Comments': 'Late arrival document',
                'Result': 'Not Handled',
                'Agenda Item': '7.1',
                'e-mail_Discussion': ''
            }
        }

        # This should execute without UnboundLocalError
        self.model.merge_agenda_data(agenda_data)

        # Verify row was added
        tdoc_map = {r.get('TDoc'): r for r in self.model._data}
        self.assertIn('S2-2609298', tdoc_map)
        row = tdoc_map['S2-2609298']
        self.assertEqual(row['Title'], 'Late Submitted Discussion Paper')
        self.assertEqual(row['Source'], 'Ericsson')
        self.assertEqual(row['Type'], 'Discussion')
        self.assertEqual(row['Related WIs'], '')
        self.assertEqual(row['Agenda Item'], '7.1')
        self.assertEqual(row['TDoc Status'], 'Not Handled')
        self.assertIn('S2-2609298', self.model.valid_tdocs)

    def test_merge_agenda_revision_inherits_related_wis_and_metadata(self):
        """
        Tests injecting a revision of an existing base TDoc (e.g. S2-260123r01).
        Verifies that Related WIs, Agenda Item, and doc type are correctly inherited from base_row.
        """
        agenda_data = {
            'S2-260123r01': {
                'Title': '',
                'Source': '',
                'Type': '',
                'For': '',
                'Comments': 'Revision of S2-260123',
                'Result': 'Agreed',
                'Agenda Item': '',
                'e-mail_Discussion': ''
            }
        }

        self.model.merge_agenda_data(agenda_data)

        tdoc_map = {r.get('TDoc'): r for r in self.model._data}
        self.assertIn('S2-260123r01', tdoc_map)
        rev_row = tdoc_map['S2-260123r01']
        base_row = tdoc_map['S2-260123']

        # Inherited from base_row
        self.assertEqual(rev_row['Related WIs'], 'FS_eVoNR, 5G_Adv')
        self.assertEqual(rev_row['Agenda Item'], '5.2')
        self.assertEqual(rev_row['Title'], 'Initial Base Document')
        self.assertEqual(rev_row['Source'], 'Nokia')
        self.assertEqual(rev_row['Is revision of'], 'S2-260123')
        self.assertEqual(base_row['Revised to'], 'S2-260123r01')

    def test_merge_agenda_revision_with_missing_base_tdoc(self):
        """
        Tests injecting a revision whose base TDoc is not in model._data.
        Verifies that it handles absent base_row gracefully without exceptions.
        """
        agenda_data = {
            'S2-269999r01': {
                'Title': 'Orphan Revision',
                'Source': 'Qualcomm',
                'Type': '',
                'For': '',
                'Comments': '',
                'Result': 'Not Handled',
                'Agenda Item': '9.9',
                'e-mail_Discussion': ''
            }
        }

        self.model.merge_agenda_data(agenda_data)

        tdoc_map = {r.get('TDoc'): r for r in self.model._data}
        self.assertIn('S2-269999r01', tdoc_map)
        row = tdoc_map['S2-269999r01']
        self.assertEqual(row['Related WIs'], '')
        self.assertEqual(row['Is revision of'], 'S2-269999')

    def test_parse_chair_result_normalizes_sa2_and_ran_statuses(self):
        """Verifies parse_chair_result extracts clean canonical statuses, remarks, and revision targets."""
        from modules.meetings.core.tdocs.chair_notes_agreements_parser import parse_chair_result

        cases = [
            ("", ("", "", [])),
            ("APPROVED", ("Approved", "", [])),
            ("NOTED", ("Noted", "", [])),
            ("POSTPONED", ("Postponed", "", [])),
            ("WITHDRAWN", ("Withdrawn", "", [])),
            ("OPEN", ("Open", "", [])),
            ("Replied to", ("Replied to", "", [])),
            ("Revised to S2-2612689", ("Revised", "", ["S2-2612689"])),
            ("revised to S2-2613055", ("Revised", "", ["S2-2613055"])),
            ("Rev in R3-264569 Agreed", ("Agreed", "Rev in R3-264569 Agreed", ["R3-264569"])),
            ("Merged into S2-2612833", ("Merged", "Merged into S2-2612833", [])),
            (
                "OPEN\nDiscuss if agree with EN, and word the EN",
                ("Open", "Discuss if agree with EN, and word the EN", []),
            ),
            (
                "OPEN  Once updated we will add these slides to the tdoc",
                ("Open", "Once updated we will add these slides to the tdoc", []),
            ),
            (
                "Revision assigned after offline discussion, not opened yet.",
                ("Not Handled", "Revision assigned after offline discussion, not opened yet.", []),
            ),
            ("Block Approve?", ("Approved", "Block Approve?", [])),
            ("Block Approved", ("Approved", "", [])),
            ("Agreed?", ("", "Agreed?", [])),
            ("not noted", ("Not Noted", "", [])),
            ("🟢 Agreement", ("🟢 Agreement", "", [])),
        ]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(parse_chair_result(raw), expected)

    def test_merge_agenda_normalizes_chair_statuses_and_links_revisions(self):
        """Verifies merge_agenda_data keeps TDoc Status clean and preserves remarks/links."""
        base_rows = [
            {
                'TDoc': 'S2-2612591',
                'Title': 'LCS direct interface',
                'Source': 'Ericsson',
                'Type': 'P-CR',
                'For': 'Approval',
                'Secretary Remarks': '',
                'Agenda Item': '20.6.14',
                'TDoc Status': 'Available',
                'Is revision of': '',
                'Revised to': '',
                'Related WIs': 'FS_6G_ARC',
            },
            {
                'TDoc': 'S2-2612689',
                'Title': 'LCS direct interface',
                'Source': 'Ericsson',
                'Type': 'P-CR',
                'For': 'Approval',
                'Secretary Remarks': '',
                'Agenda Item': '20.6.14',
                'TDoc Status': 'Available',
                'Is revision of': '',
                'Revised to': '',
                'Related WIs': 'FS_6G_ARC',
            },
            {
                'TDoc': 'S2-2611681',
                'Title': 'Agentic AI',
                'Source': 'NCSC',
                'Type': 'DISCUSSION',
                'For': 'Endorsement',
                'Secretary Remarks': '',
                'Agenda Item': '20.6.0',
                'TDoc Status': 'Available',
                'Is revision of': '',
                'Revised to': '',
                'Related WIs': 'FS_6G_ARC',
            },
        ]
        model = TDocsTableModel(meeting_dir=self.meeting_dir, data=base_rows)

        agenda_data = {
            'S2-2612591': {
                'Comments': 'Discuss offline',
                'Result': 'Revised to S2-2612689',
                'Agenda Item': '20.6.14',
                'e-mail_Discussion': '',
            },
            'S2-2611681': {
                'Comments': 'Handle',
                'Result': 'Block Approve?',
                'Agenda Item': '20.6.0',
                'e-mail_Discussion': '',
            },
            'S2-2613114': {
                'Title': 'Measurement based CN PDB',
                'Source': 'Huawei, HiSilicon',
                'Type': 'CR',
                'For': 'Approval',
                'Comments': 'Revision of S2-2613039',
                'Result': 'OPEN\nDiscuss if agree with EN, and word the EN',
                'Agenda Item': '20.50',
                'e-mail_Discussion': '',
            },
            'S2-2612832': {
                'Title': 'DL Data Transfer Assistance',
                'Source': 'China Mobile',
                'Type': 'CR',
                'For': 'Approval',
                'Comments': 'Revision of S2-2612797',
                'Result': 'Merged into S2-2612833',
                'Agenda Item': '20.5.2',
                'e-mail_Discussion': '',
            },
            'S2-2612931': {
                'Title': 'Consolidated architecture proposal',
                'Source': 'China Mobile',
                'Type': 'P-CR',
                'For': 'Approval',
                'Comments': 'Revision of S2-2612168',
                'Result': 'Revision assigned after offline discussion, not opened yet.',
                'Agenda Item': '20.6.0',
                'e-mail_Discussion': '',
            },
        }
        model.merge_agenda_data(agenda_data)
        tdoc_map = {r.get('TDoc'): r for r in model._data}

        # Stale Available updated to Revised, Revised to linked bidirectionally
        self.assertEqual(tdoc_map['S2-2612591']['TDoc Status'], 'Revised')
        self.assertEqual(tdoc_map['S2-2612591']['Revised to'], 'S2-2612689')
        self.assertEqual(tdoc_map['S2-2612689']['Is revision of'], 'S2-2612591')

        # Block Approve? promotes stale Available -> Approved while keeping original note in Secretary Remarks
        self.assertEqual(tdoc_map['S2-2611681']['TDoc Status'], 'Approved')
        self.assertIn('Block Approve?', tdoc_map['S2-2611681']['Secretary Remarks'])

        # Multi-line OPEN status split cleanly
        self.assertEqual(tdoc_map['S2-2613114']['TDoc Status'], 'Open')
        self.assertIn('Discuss if agree with EN, and word the EN', tdoc_map['S2-2613114']['Secretary Remarks'])

        # Merged into -> Merged + remark link preserved
        self.assertEqual(tdoc_map['S2-2612832']['TDoc Status'], 'Merged')
        self.assertIn('Merged into S2-2612833', tdoc_map['S2-2612832']['Secretary Remarks'])

        # Not opened yet -> Not Handled + full text moved to Secretary Remarks
        self.assertEqual(tdoc_map['S2-2612931']['TDoc Status'], 'Not Handled')
        self.assertIn('Revision assigned after offline discussion, not opened yet.', tdoc_map['S2-2612931']['Secretary Remarks'])

    def test_sub_item_combo_includes_no_sub_item_for_mixed_rows(self):
        """Verifies rows without Sub-Item remain visible and statuses are not greyed out."""
        from PyQt5.QtCore import QSettings
        import tempfile as tf
        from modules.meetings.ui.tdocs_window import TDocsWindow

        rows = [
            {
                'TDoc': 'S2-260100',
                'Title': 'With Sub-Item',
                'Source': 'Nokia',
                'Type': 'CR',
                'For': 'Approval',
                'Agenda Item': '20.6.1',
                'Sub-Item': 'NAS Routing',
                'TDoc Status': 'Agreed',
                'Related WIs': 'FS_6G_ARC',
            },
            {
                'TDoc': 'S2-260101',
                'Title': 'Without Sub-Item',
                'Source': 'Ericsson',
                'Type': 'CR',
                'For': 'Approval',
                'Agenda Item': '20.6.14',
                'Sub-Item': '',
                'TDoc Status': 'Merged',
                'Related WIs': 'FS_6G_ARC',
            },
        ]
        with tf.TemporaryDirectory() as tmp_dir:
            agenda_path = Path(tmp_dir) / 'Agenda' / 'test.xlsx'
            agenda_path.parent.mkdir(parents=True, exist_ok=True)
            qsettings = QSettings(str(Path(tmp_dir) / 'filters.ini'), QSettings.IniFormat)
            win = TDocsWindow({'wg_name': 'SA2', 'meeting_number': '178'}, rows, str(agenda_path), filter_settings=qsettings)
            self.assertEqual(win.proxy.rowCount(), 2)
            # Both 'Agreed' and 'Merged' should be available (UserRole+1 == True)
            status_model = win.status_combo.model()
            avail = {
                status_model.item(i).data(256): status_model.item(i).data(257)  # Qt.UserRole=256, Qt.UserRole+1=257
                for i in range(1, status_model.rowCount())
            }
            self.assertTrue(avail.get('Agreed'))
            self.assertTrue(avail.get('Merged'))
            win.close()


if __name__ == '__main__':
    unittest.main()
