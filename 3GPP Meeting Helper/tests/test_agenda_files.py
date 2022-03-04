import unittest
import server.tdoc as server

class Test_test_agenda_files(unittest.TestCase):
    def test_129Bis(self):
        test_agenda = 'updated S2-1811604-SA2-129bis-Agenda-v12.docx'
        agenda_match = server.agenda_regex.match(test_agenda)
        agenda_docx_match = server.agenda_docx_regex.match(test_agenda)
        self.assertIsNotNone(agenda_match)
        self.assertEqual(agenda_match.groupdict()['version'], '12')
        self.assertIsNotNone(agenda_docx_match)
        self.assertEqual(agenda_docx_match.groupdict()['version'], '12')

    def test_non_agenda(self):
        test_agenda = 'DRAFT S2-1812771 (was 12339 KI #9 solution 13 update).DOC'
        agenda_match = server.agenda_regex.match(test_agenda)
        agenda_docx_match = server.agenda_docx_regex.match(test_agenda)
        self.assertIsNone(agenda_match)
        self.assertIsNone(agenda_docx_match)

    def test_129_zip(self):
        test_agenda = 'draft S2-18nnnnn-SA2-129-Agenda-v1h.zip'
        agenda_match = server.agenda_regex.match(test_agenda)
        self.assertIsNotNone(agenda_match)
        self.assertEqual(agenda_match.groupdict()['version'], '1')

    def test_129_doc(self):
        test_agenda = 'draft S2-18nnnnn-SA2-129-Agenda-v1h.doc'
        agenda_match = server.agenda_regex.match(test_agenda)
        self.assertIsNotNone(agenda_match)
        self.assertEqual(agenda_match.groupdict()['version'], '1')

    def test_agenda_file_list_with_drafts_and_non_drafts(self):
        agenda_list = [ 
            'draft_S2-1910847_SA2-136-Agenda-v3.docx', 
            'draft_S2-1910847_SA2-136-Agenda-v2.docx', 
            'draft_S2-1910847_SA2-136-Agenda-v10.docx',
            'draft_S2-1910847_SA2-136-Agenda-v1.docx',
            'S2-1911962_SA2-136-Agenda_v1.docx',
            'S2-1911962_SA2-136-Agenda_v2.docx',
            'S2-1911975_SA2-136-Agenda_v3.docx',
            'Some other file.docx']
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNotNone(latest_agenda)
        self.assertEqual(latest_agenda, 'S2-1911975_SA2-136-Agenda_v3.docx')

    def test_agenda_file_list_with_drafts(self):
        agenda_list = [ 
            'draft_S2-1910847_SA2-136-Agenda-v3.docx', 
            'draft_S2-1910847_SA2-136-Agenda-v2.docx', 
            'draft_S2-1910847_SA2-136-Agenda-v10.docx',
            'draft_S2-1910847_SA2-136-Agenda-v1.docx',
            'Some other file.docx']
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNotNone(latest_agenda)
        self.assertEqual(latest_agenda, 'draft_S2-1910847_SA2-136-Agenda-v10.docx')

    def test_agenda_file_list_with_non_drafts(self):
        agenda_list = [ 
            'S2-1911962_SA2-136-Agenda_v1.docx',
            'S2-1911962_SA2-136-Agenda_v2.docx',
            'S2-1911975_SA2-136-Agenda_v3.docx',
            'Some other file.docx']
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNotNone(latest_agenda)
        self.assertEqual(latest_agenda, 'S2-1911975_SA2-136-Agenda_v3.docx')

    def test_agenda_file_list_with_non_agendas(self):
        agenda_list = [ 
            'Some other file.docx']
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNone(latest_agenda)

    def test_agenda_file_list_with_empty(self):
        agenda_list = []
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNone(latest_agenda)

    def test_agenda_file_list_with_none(self):
        agenda_list = None
        latest_agenda = server.get_latest_agenda_file(agenda_list)
        self.assertIsNone(latest_agenda)

    def test_file_names(self):
        file_name = 'S2-2000001_SA2-136AH-Agenda_v1-cl.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.01)

        file_name = 'S2-2000001_SA2-136AH-Agenda_v1-rm.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.01)

        file_name = 'S2-2000001_SA2-136AH-Agenda_v2.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.02)

        file_name = 'S2-2000001_SA2-136AH-Agenda_v2-cl.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.02)

        file_name = 'S2-2000001_SA2-136AH-Agenda_v1cl.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.01)

        file_name = 'S2-2000001_SA2-136AH-Agenda_v1rm.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 2000001.01)

        file_name = 'Agenda_v1rm.docx'
        version = server.get_agenda_file_version_number(file_name)
        self.assertEqual(version, 0.01)

if __name__ == '__main__':
    unittest.main()
