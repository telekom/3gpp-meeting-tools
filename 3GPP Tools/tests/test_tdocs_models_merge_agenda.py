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


if __name__ == '__main__':
    unittest.main()
