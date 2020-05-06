import unittest
import os.path
import parsing.word as word_parser

class Test_test_agenda_parsing(unittest.TestCase):
    def test_agenda_draft_135_v4(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'agenda', 'draft_S2-190xxxx_SA2-135-Agenda-v4.docx')
        agenda_names = word_parser.import_agenda(file_name)
        self.assertEqual(agenda_names['1'], 'Opening of the meeting   9:00 on Monday')

    def test_agenda_134_v15(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'agenda', 'updated_S2-1906993_SA2-134-Agenda-v15.docx')
        agenda_names = word_parser.import_agenda(file_name)
        self.assertEqual(agenda_names['6.8'], 'Access Traffic Steering, Switch and Splitting support in the 5G system architecture (ATSSS)')

if __name__ == '__main__':
    unittest.main()
