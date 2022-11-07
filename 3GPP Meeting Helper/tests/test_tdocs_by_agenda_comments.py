import unittest
import parsing.html.common as html_parser

class Test_test_tdocs_by_agenda_comments(unittest.TestCase):
    def test_tdoc_comment_empty(self):
        comments = ''
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, '')
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')

    def test_tdoc_comment_none(self):
        comments = None
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, '')
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')

    def test_tdoc_comment_s21901105(self):
        comments = 'Revision of S2-1900064, merging S2-1900142, S2-1900585, S2-1900281, S2-1900147 and part of S2-1900587. Revised in parallel session to S2-1901260.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, 'S2-1900142, S2-1900585, S2-1900281, S2-1900147, S2-1900587')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, 'S2-1900064')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901260')

    def test_tdoc_comment_s219000064(self):
        comments = 'Revised in parallel session, merging S2-1900142, S2-1900585, S2-1900281, S2-1900147 and part of S2-1900587, to S2-1901105.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, 'S2-1900142, S2-1900585, S2-1900281, S2-1900147, S2-1900587')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901105')

    def test_tdoc_comment_s21900142(self):
        comments = 'Merged into S2-1901105.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, 'S2-1901105')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, '')

    def test_tdoc_comment_s21900587(self):
        comments = 'Merged into S2-1901105 and S2-1901106.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, 'S2-1901105, S2-1901106')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, '')

    def test_tdoc_comment_s21901262(self):
        comments = 'Revision of S2-1901108. This CR was agreed.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, 'S2-1901108')
        self.assertEqual(parsed_comments.revised_to, '')

    def test_tdoc_comment_s21900272(self):
        comments = '	Revised in parallel session, merging parts of S2-1900611, to S2-1901220.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, 'S2-1900611')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901220')

    def test_tdoc_comment_s21901220(self):
        comments = 'Revision of S2-1900272, merging parts of S2-1900611. This was postponed.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, 'S2-1900611')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, 'S2-1900272')
        self.assertEqual(parsed_comments.revised_to, '')

    def test_tdoc_comment_s21900611(self):
        comments = 'Partially merged into S2-1901220 and revised in S2-1901221. Revised in parallel session to S2-1901221.	'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, 'S2-1901220')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901221')

    def test_tdoc_comment_s2190441(self):
        comments = 'Revised in parallel session, merging S2-1900502 and S2-1900563 to S2-1901222.'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, 'S2-1900502, S2-1900563')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901222')

    def test_tdoc_comment_s2190497(self):
        comments = 'Revision of S2-1811853 from S2#129BIS. Revised in parallel session to S2-1901140. NSSAI IWK'
        parsed_comments = html_parser.parse_tdoc_comments(comments)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901140')

    def test_tdoc_comment_s2190497_2(self):
        comments = 'Revision of S2-1811853 from S2#129BIS. Revised in parallel session to S2-1901140. NSSAI IWK'
        parsed_comments = html_parser.parse_tdoc_comments(comments, ignore_from_previous_meetings=False)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, 'S2-1811853')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901140')

    def test_tdoc_comment_s2190497_3(self):
        comments = 'Revision of S2-1811853 from S2#129BIS. Revised in parallel session to S2-1901140. NSSAI IWK'
        parsed_comments = html_parser.parse_tdoc_comments(comments, ignore_from_previous_meetings=True)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, '')
        self.assertEqual(parsed_comments.revised_to, 'S2-1901140')

    def test_tdoc_comment_s21908318(self):
        comments = 'Revision of S2-1907759. Revised in parallel session, merging related CRs, to S2-1908320.'
        parsed_comments = html_parser.parse_tdoc_comments(comments, ignore_from_previous_meetings=True)
        self.assertEqual(parsed_comments.merge_of, '')
        self.assertEqual(parsed_comments.merged_to, '')
        self.assertEqual(parsed_comments.revision_of, 'S2-1907759')
        self.assertEqual(parsed_comments.revised_to, 'S2-1908320')
        

if __name__ == '__main__':
    unittest.main()
