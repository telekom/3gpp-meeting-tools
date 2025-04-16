import unittest

from parsing.html import chairnotes


class GenericTdocParsing(unittest.TestCase):
    def test_chairnotes_1(self):
        file_list = [
            "ChairNotes_Andy_02-20-1930_DST.doc",
            "ChairNotes_Andy_02-21-1030.doc",
            "ChairNotes_Dario_02-21-0524.doc",
            "ChairNotes_Dario_02-21-0835.doc",
            "ChairNotes_Dario_02-21-0956.doc",
            "ChairNotes_Wanqiang_02-20-1739.doc",
            "ChairNotes_Wanqiang_02-21-1000.doc",
            "ChairNotes_Wanqiang_02-21-1232.doc",
            "Combined_ChairNotes_02-21-1330.doc",
            "Combined_ChairNotes_02-21-1600 (end of meeting).doc",
            "ChairNotes_Wanqiang_02-22-1232.doc",
        ]

        latest_files = chairnotes.get_latest_chairnotes_files(file_list)
        self.assertEqual(len(latest_files), 2)

        latest_file_combined = latest_files[0]
        self.assertEqual(latest_file_combined.file, 'Combined_ChairNotes_02-21-1600 (end of meeting).doc')
        self.assertEqual(latest_file_combined.is_combined, True)
        self.assertEqual(latest_file_combined.authors, ['Dario', 'Andy'])

        latest_file_wanqiang = latest_files[1]
        self.assertEqual(latest_file_wanqiang.file, 'ChairNotes_Wanqiang_02-22-1232.doc')
        self.assertEqual(latest_file_wanqiang.is_combined, False)
        self.assertEqual(latest_file_wanqiang.authors, ['Wanqiang'])

    def test_chairnotes_2(self):
        file_list = [
            "ChairNotes_Andy_02-20-1930_DST.doc",
            "ChairNotes_Andy_02-21-1030.doc",
            "ChairNotes_Dario_02-21-0524.doc",
            "ChairNotes_Dario_02-21-0835.doc",
            "ChairNotes_Dario_02-21-0956.doc",
            "ChairNotes_Wanqiang_02-20-1739.doc",
            "ChairNotes_Wanqiang_02-21-1000.doc",
            "ChairNotes_Wanqiang_02-21-1232.doc",
            "Combined_ChairNotes_02-21-1330.doc",
            "Combined_ChairNotes_02-21-1600 (end of meeting).doc"
        ]

        latest_files = chairnotes.get_latest_chairnotes_files(file_list)
        self.assertEqual(len(latest_files), 1)
        latest_file = latest_files[0]
        self.assertEqual(latest_file.file, 'Combined_ChairNotes_02-21-1600 (end of meeting).doc')
        self.assertEqual(latest_file.is_combined, True)
