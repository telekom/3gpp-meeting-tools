import unittest
import os

import parsing
from parsing.html.common import TdocsByAgendaData
import parsing.html.tdocs_by_agenda_v3


class TestTdocsByAgenda(unittest.TestCase):
    def test_129Bis(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '129Bis.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '129BIS', 'Expected 129BIS')

        self.assertEqual(len(meeting.tdocs), 1529, 'Expected TDoc entries')

        self.assertEqual(meeting.tdocs.at['S2-1812368', 'Revised to'], 'S2-1813194', 'Expected Revised do S2-1813194')
        self.assertEqual(meeting.tdocs.at['S2-1813194', 'Revision of'], 'S2-1812368', 'Expected Revision of S2-1812368')

        self.assertEqual(meeting.tdocs.at['S2-1812440', 'Merged to'], 'S2-1813085', 'Expected Merged to S2-1813085')
        self.assertEqual(meeting.tdocs.at['S2-1813085', 'Merge of'], 'S2-1812440', 'Expected Merge of S2-1812440')

        # Meetings 'from S2#129' or similar (i.e. postponed from a past meeting) are not taken as past references.
        # This is so as to make reporting easier and to make the .htm file self.contained
        self.assertEqual(meeting.tdocs.at['S2-1813085', 'Original TDocs'], 'S2-1812299, S2-1812440',
                         'Expected Original TDocs S2-18122990,S2-1812440')
        self.assertEqual(meeting.tdocs.at['S2-1813085', 'Final TDocs'], 'S2-1813308', 'Expected Final TDocs S2-1813308')

        # Check some contribution mappings
        self.assertTrue(meeting.tdocs.at['S2-1811737', 'Contributed by Deutsche Telekom'], 'DT contribution')
        self.assertTrue(meeting.tdocs.at['S2-1811737', 'Contributed by TIM'], 'Telecom Italia contribution')
        self.assertTrue(meeting.tdocs.at['S2-1811737', 'Contributed by Intel'], 'Intel contribution')
        self.assertFalse(meeting.tdocs.at['S2-1811737', 'Contributed by Qualcomm'], 'Qualcomm contribution')
        self.assertFalse(meeting.tdocs.at['S2-1811737', 'Contributed by Nokia'], 'Nokia contribution')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 14, 'Length of the "Others" contributors')

    def test_129(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '129.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '129', 'Expected 129')

        self.assertEqual(len(meeting.tdocs), 1508, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 17, 'Length of the "Others" contributors')

    def test_128Bis(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128Bis.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '128BIS', 'Expected 128BIS')

        self.assertEqual(len(meeting.tdocs), 1401, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 12, 'Length of the "Others" contributors')

    def test_128(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '128', 'Expected 128')

        self.assertEqual(len(meeting.tdocs), 1301, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 12, 'Length of the "Others" contributors')

    def test_inbox(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', 'inbox.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '129BIS', 'Expected 129BIS')

        self.assertEqual(len(meeting.tdocs), 1640, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 14, 'Length of the "Others" contributors')

    def test_130_1(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2019.01.22 TdocsByAgenda.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130')

        self.assertEqual(len(meeting.tdocs), 853, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 16, 'Length of the "Others" contributors')

    def test_130_2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2019.01.24 TdocsByAgenda.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130')

        self.assertEqual(len(meeting.tdocs), 1026, 'Expected TDoc entries')

        # Check the length of the "Others" mapping
        self.assertEqual(len(meeting.others_cosigners), 16, 'Length of the "Others" contributors')

    def test_130_3(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '2019.01.31 130.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130S')

        self.assertEqual(len(meeting.tdocs), 1306, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-1901378', :]

        revised_to = test_row['Revised to']
        revision_of = test_row['Revision of']
        merged_to = test_row['Merged to']
        merge_of = test_row['Merge of']
        self.assertEqual(revised_to, '')
        self.assertEqual(revision_of, 'S2-1901260')
        self.assertEqual(merged_to, '')
        self.assertEqual(merge_of, '')

        # Meetings 'from S2#129' or similar (i.e. postponed from a past meeting) are not taken as past references.
        # This is so as to make reporting easier and to make the .htm file self.contained
        original_tdocs = test_row['Original TDocs']
        final_tdocs = test_row['Final TDocs']
        self.assertEqual(original_tdocs, 'S2-1900142, S2-1900147, S2-1900281, S2-1900585, S2-1900587')
        self.assertEqual(final_tdocs, 'S2-1901378')

        # Check some contribution mappings
        self.assertTrue(test_row['Contributed by Deutsche Telekom'], 'DT contribution')
        self.assertFalse(test_row['Contributed by TIM'], 'Not a Telecom Italia contribution')
        self.assertFalse(test_row['Contributed by Intel'], 'Not a Intel contribution')
        self.assertTrue(test_row['Contributed by ZTE'], 'ZTE contribution')
        self.assertTrue(test_row['Contributed by Sprint'], 'Nokia Sprint')

    def test_130_4(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '2019.01.31 130.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130S')

        self.assertEqual(len(meeting.tdocs), 1306, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-1900064', :]

        revised_to = test_row['Revised to']
        revision_of = test_row['Revision of']
        merged_to = test_row['Merged to']
        merge_of = test_row['Merge of']
        self.assertEqual(revised_to, 'S2-1901105')
        self.assertEqual(revision_of, '')
        self.assertEqual(merged_to, '')
        self.assertEqual(merge_of, 'S2-1900142, S2-1900585, S2-1900281, S2-1900147, S2-1900587')

        # Meetings 'from S2#129' or similar (i.e. postponed from a past meeting) are not taken as past references.
        # This is so as to make reporting easier and to make the .htm file self.contained
        original_tdocs = test_row['Original TDocs']
        final_tdocs = test_row['Final TDocs']
        self.assertEqual(original_tdocs, 'S2-1900142, S2-1900147, S2-1900281, S2-1900585, S2-1900587')
        self.assertEqual(final_tdocs, 'S2-1901378')

        # Check some contribution mappings
        self.assertTrue(test_row['Contributed by Deutsche Telekom'], 'DT contribution')
        self.assertFalse(test_row['Contributed by TIM'], 'Not a Telecom Italia contribution')
        self.assertFalse(test_row['Contributed by Intel'], 'Not a Intel contribution')
        self.assertFalse(test_row['Contributed by ZTE'], 'ZTE contribution')
        self.assertFalse(test_row['Contributed by Sprint'], 'Nokia Sprint')

    def test_date_130(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '2019.01.31 130.htm')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2019)
        self.assertEqual(datetime.month, 1)
        self.assertEqual(datetime.day, 25)
        self.assertEqual(datetime.hour, 17)
        self.assertEqual(datetime.minute, 25)

    def test_date_129Bis(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '129Bis.htm')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2018)
        self.assertEqual(datetime.month, 11)
        self.assertEqual(datetime.day, 30)
        self.assertEqual(datetime.hour, 7)
        self.assertEqual(datetime.minute, 14)

    def test_date_129(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '129.htm')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2018)
        self.assertEqual(datetime.month, 10)
        self.assertEqual(datetime.day, 27)
        self.assertEqual(datetime.hour, 10)
        self.assertEqual(datetime.minute, 18)

    def test_date_128Bis(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128Bis.htm')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2018)
        self.assertEqual(datetime.month, 8)
        self.assertEqual(datetime.day, 31)
        self.assertEqual(datetime.hour, 13)
        self.assertEqual(datetime.minute, 29)

    def test_date_128(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128.htm')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2018)
        self.assertEqual(datetime.month, 7)
        self.assertEqual(datetime.day, 13)
        self.assertEqual(datetime.hour, 14)
        self.assertEqual(datetime.minute, 38)

    def test_date_comparison_128_128Bis(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128Bis.htm')
        datetime_128Bis = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)

        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '128.htm')
        datetime_128 = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)

        self.assertGreater(datetime_128Bis, datetime_128)

    def test_corrupt_dtdocs_by_agenda(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2019.07.02 TDocsByAgenda_wrong.html')
        datetime = TdocsByAgendaData.get_tdoc_by_agenda_date(file_name)
        self.assertEqual(datetime.year, 2019)
        self.assertEqual(datetime.month, 6)
        self.assertEqual(datetime.day, 28)
        self.assertEqual(datetime.hour, 16)
        self.assertEqual(datetime.minute, 36)

        meeting = TdocsByAgendaData(file_name)
        self.assertEqual(meeting.meeting_number, 'Unknown', 'Expected Unknown')
        self.assertEqual(len(meeting.tdocs), 1, 'Expected TDoc entries')

    def test_exported_dtdocs_by_agenda_v1(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2019.01.24 TdocsByAgenda.htm')
        meeting = TdocsByAgendaData(file_name, v=1)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130')
        self.assertEqual(len(meeting.tdocs), 1026, 'Expected TDoc entries')

    def test_exported_dtdocs_by_agenda_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2019.01.24 TdocsByAgenda.htm')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '130', 'Expected 130')
        self.assertEqual(len(meeting.tdocs), 1026, 'Expected TDoc entries')

    def test_exported_dtdocs_by_agenda_134_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '134.htm')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '134', 'Expected 134')
        self.assertEqual(len(meeting.tdocs), 1802, 'Expected TDoc entries')
        test_row = meeting.tdocs.loc['S2-1908578', :]
        self.assertEqual(test_row['Revision of'], 'S2-1908544', 'Expected S2-1908544')

    def test_136(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '136.html')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '136', 'Expected 136')
        self.assertEqual(len(meeting.tdocs), 1201, 'Expected TDoc entries')

    def test_136_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '136_v2.html')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '136', 'Expected 136')
        self.assertEqual(len(meeting.tdocs), 1762, 'Expected TDoc entries')

    def test_136_v3(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '136_v3.html')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '136', 'Expected 136')
        self.assertEqual(len(meeting.tdocs), 1815, 'Expected TDoc entries')

    def test_136_missing_ais(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '136_missing_AIs.html')
        meeting = TdocsByAgendaData(file_name, v=2)

        self.assertEqual(meeting.meeting_number, '136', 'Expected 136')
        self.assertEqual(len(meeting.tdocs), 1857, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-1910969', :]
        self.assertEqual(test_row['Title'],
                         'LS from SA WG3LI: LS on Enhancing Location Information Reporting with Dual Connectivity')
        self.assertEqual(test_row['Result'], 'Noted')
        self.assertEqual(test_row['Comments'], 'Noted')
        self.assertEqual(test_row['AI'], '', 'Expected empty string')

    def test_137e(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2020.02.24 TdocsByAgenda SA2-137E.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '137', 'Expected 137')
        self.assertEqual(len(meeting.tdocs), 544, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-2001981', :]
        self.assertEqual(test_row['AI'], '5.4', 'Expected AI 5.4')

    def test_137e_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2020.02.29 TdocsByAgenda SA2-137E.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '137', 'Expected 137')
        self.assertEqual(len(meeting.tdocs), 713, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-2001981', :]
        self.assertEqual(test_row['AI'], '5.4', 'Expected AI 5.4')

    def test_138e(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', '138E_final.html')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '138E', 'Expected 138E')
        self.assertEqual(len(meeting.tdocs), 824, 'Expected TDoc entries')

        test_row = meeting.tdocs.loc['S2-2002769', :]
        self.assertEqual(test_row['AI'], '7.10.3', 'Expected AI 7.10.3')
        self.assertEqual(test_row['Title'], "23.502 CR2190 (Rel-16, 'F'): UE radio capability for 5GS and IWK")
        self.assertEqual(test_row['Result'], 'Revised')

        test_row = meeting.tdocs.loc['S2-2002694', :]
        self.assertEqual(test_row['AI'], '6.3', 'Expected AI 6.3')
        self.assertEqual(test_row['Title'],
                         "23.501 CR2226 (Rel-16, 'F'): Multiple N6 interfaces per Network Instance for Ethernet traffic")
        self.assertEqual(test_row['Comments'], 'Noted')
        self.assertEqual(test_row['Result'], 'Noted')

    def test_155(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2023.02.14 TdocsByAgenda SA2-155.htm')
        meeting = TdocsByAgendaData(file_name)

        self.assertEqual(meeting.meeting_number, '155', 'Expected 155')
        self.assertEqual(len(meeting.tdocs), 991, 'Expected TDoc entries')

    def test_159_format_155_file(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2023.02.14 TdocsByAgenda SA2-155.htm')
        html_content = parsing.html.common.TdocsByAgendaData.get_tdoc_by_agenda_html(file_name, return_raw_html=True)
        is_159_format = parsing.html.tdocs_by_agenda_v3.assert_if_tdocs_by_agenda_post_sa2_159(html_content)
        self.assertFalse(is_159_format)

    def test_159_format_159_file_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2023.10.03 TdocsByAgenda SA2-159_v2.htm')
        html_content = parsing.html.common.TdocsByAgendaData.get_tdoc_by_agenda_html(file_name, return_raw_html=True)
        is_159_format = parsing.html.tdocs_by_agenda_v3.assert_if_tdocs_by_agenda_post_sa2_159(html_content)
        self.assertTrue(is_159_format)

    def test_159_file_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2023.10.03 TdocsByAgenda SA2-159_v2.htm')
        html_content = parsing.html.common.TdocsByAgendaData.get_tdoc_by_agenda_html(file_name, return_raw_html=True)
        sa2_159_content = parsing.html.tdocs_by_agenda_v3.parse_tdocs_by_agenda_v3(html_content)
        test_row = sa2_159_content.loc['S2-2310462', :]
        # print(test_row)
        self.assertEqual(test_row['AI'], '8.9', 'Expected AI 8.9')
        self.assertEqual(test_row['Title'],
                         "23.247 CR0307 (Rel-18, 'F'): Update UE pre-configuration")
        self.assertEqual(test_row['Comments'], '')
        self.assertEqual(test_row['Result'], '')

    def test_159_meeting_number_159_file_v2(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda',
                                 '2023.10.03 TdocsByAgenda SA2-159_v2.htm')
        meeting = TdocsByAgendaData(file_name)
        self.assertEqual(meeting.meeting_number, '159', 'Expected 159')
        self.assertEqual(len(meeting.tdocs), 1190, 'Expected 1190 TDoc entries')

if __name__ == '__main__':
    unittest.main()
