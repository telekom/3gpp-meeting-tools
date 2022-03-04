import unittest
import pandas as pd
import os.path

import application.word
import parsing.html as html_parser
import parsing.word as word_parser

class Test_test_report(unittest.TestCase):
    @unittest.skip("Skipping tests that open Word")
    def test_empty_df(self):
        columns = [ 'AI', 'Type', 'Doc For', 'Title', 'Source', 'Rel', 'Work Item', 'Comments', 'Result', 'Revision of', 'Revised to', 'Merge of', 'Merged to', '#', 'TS', 'CR', 'Original TDocs' ]
        df = pd.DataFrame(columns=columns)
        df.index.name = 'TD#'

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc(df, doc, 'TSGS2_134_Sapporo')

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_example_df(self):
        columns = [ 'AI', 'Type', 'Doc For', 'Title', 'Source', 'Rel', 'Work Item', 'Comments', 'Result', 'Revision of', 'Revised to', 'Merge of', 'Merged to', '#', 'TS', 'CR', 'Original TDocs' ]
        
        data = [
            ('2', 'AGENDA', 'Approval', 'Draft Agenda for SA WG2#133', 'SA WG2 Chairman', '', '', 'Approved', 'Approved', '', '', '', '', '0', '', '', 'S2-1904846'),
            ('3', 'REPORT', 'Approval', 'Draft Report of SA WG2 meeting #132', 'SA WG2 Secretary', '', '', 'Approved', 'Approved', '', '', '', '', '1', '', '', 'S2-1904847'),
            ('4.1', 'LS In', 'Action', 'LS from RAN WG2: LS to SA2 and SA5 on IAB impact to CN', 'RAN WG2 (R2-1905475)', 'Rel-16', 'NR_IAB-Core', 'This will be taken into account in handling of the related WI. Noted', 'Noted', '', '', '', '', '2', '', '', 'S2-1904889'),
            ('4.1', 'LS In', 'Information', "LS from ITU-T SG13: LS on information about consent of ITU-T Recommendation Y.3172 'Architectural framework for machine learning in future networks including IMT-2020'", 'ITU-T SG13 (SG13-LS95.)', '', '', 'Noted', 'Noted', '', '', '', '', '3', '', '', 'S2-1904946'),
            ('4.1', 'LS In', 'Action', 'LS from SA WG4: LS on TS 26.501 5G Media Streaming (5GMS); General description and architecture', 'SA WG4 (S4h190837)', 'Rel-16', '5GMSA', 'Postponed', 'Postponed', '', '', '', '', '4', '', '', 'S2-1905810'),
            ('4.1', 'LS In', 'Information', 'LS from NGMN:  NGMN 5G End-to-End Architecture Framework', 'NGMN (Liaison_NGMN_P1_to_3GPP.)', '', '', 'Noted', 'Noted', '', '', '', '', '5', '', '', 'S2-1905823'),
            ('5.1', 'LS In', 'Information', 'LS from SA WG4: Reply LS (to RAN WG2) for inclusion of Receive Only Mode MBMS service parameters in USD', 'SA WG4 (S4-190563)', 'Rel-15', 'TEI15', 'Noted in parallel session.', 'Noted', '', '', '', '', '6', '', '', 'S2-1904901'),
            ('5.1', 'LS In', 'Action', 'LS from SA WG3LI: Response LS on reporting all Cell IDs in 5G', 'SA WG3LI (S3i190265)', 'Rel-15', 'LI15', 'Response drafted in S2-1906031. Final response in S2-1906170', 'Replied to', '', '', '', '', '7', '', '', 'S2-1904945'),
            ('5.1', 'LS OUT', 'Approval', '[DRAFT] Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Created in parallel session. Response to S2-1904945. Revised in parallel session to S2-1906166.', 'Revised', '', 'S2-1906166', '', '', '8', '', '', 'S2-1906031'),
            ('5.1', 'LS OUT', 'Approval', '[DRAFT] Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Revision of S2-1906031. Revised in parallel session to S2-1906170.', 'Revised', 'S2-1906031', 'S2-1906170', '', '', '9', '', '', 'S2-1906031'),
            ('5.1', 'LS OUT', 'Approval', 'Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Revision of S2-1906166. Agreed in parallel session. This was Block approved.', 'Approved', 'S2-1906166', '', '', '', '10', '', '', 'S2-1906031'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revised in parallel session to S2-1906028.', 'Revised', '', 'S2-1906028', '', '', '11', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513R1 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revision of S2-1905584. Revised in parallel session to S2-1906165.', 'Revised', 'S2-1905584', 'S2-1906165', '', '', '12', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513R2 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revision of S2-1906028. Agreed in parallel session. This was Block approved.', 'Agreed', 'S2-1906028', '', '', '', '13', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3514 (Rel-16', 'A'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-16', 'EDCE5', 'LATE DOC: Rx 17/05 16:00. For e-mail approval. Approved.', '"Agreed OPEN"', '', '', '', '', '14', '23.401', '3514', 'S2-1905613'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349 (Rel-15', ''F'): Location reporting of secondary cell", "'Nokia, 'Nokia Shanghai Bell", 'Rel-15', "LI15, TEI15", 'Revised in parallel session to S2-1906029.', 'Revised', '', 'S2-1906029', '', '', '15', '23.501', '1349', 'S2-1905243'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349R1 (Rel-15', ''F'): Location reporting of secondary cell", "Nokia, 'Nokia Shanghai Bell", 'Rel-15', "LI15, TEI15", 'Revision of S2-1905243. Revised in parallel session to S2-1906163.', 'Revised', 'S2-1905243', 'S2-1906163', '', 'S2-1906165', '16', '23.501', '1349', 'S2-1905243'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349R2 (Rel-15', ''F'): Location Reporting of secondary cell", "Nokia, 'Nokia Shanghai Bell", 'Rel-15', "TEI15, LI15", 'Revision of S2-1906029. Agreed in parallel session. This was Block approved.', 'Agreed', 'S2-1906029', '', '', '', '17', '23.501', '1349', 'S2-1905243')
        ]
        index = [ 'S2-1904846', 'S2-1904847', 'S2-1904889', 'S2-1904946', 'S2-1905810', 'S2-1905823', 'S2-1904901', 'S2-1904945', 'S2-1906031', 'S2-1906166', 'S2-1906170', 'S2-1905584', 'S2-1906028', 'S2-1906165', 'S2-1905613', 'S2-1905243', 'S2-1906029', 'S2-1906163' ]
        df = pd.DataFrame(columns=columns, data=data, index=index)
        df.index.name = 'TD#'

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc(df, doc, 'TSGS2_134_Sapporo')

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_example_df_by_wi(self):
        columns = [ 'AI', 'Type', 'Doc For', 'Title', 'Source', 'Rel', 'Work Item', 'Comments', 'Result', 'Revision of', 'Revised to', 'Merge of', 'Merged to', '#', 'TS', 'CR', 'Original TDocs' ]
        
        data = [
            ('2', 'AGENDA', 'Approval', 'Draft Agenda for SA WG2#133', 'SA WG2 Chairman', '', '', 'Approved', 'Approved', '', '', '', '', '0', '', '', 'S2-1904846'),
            ('3', 'REPORT', 'Approval', 'Draft Report of SA WG2 meeting #132', 'SA WG2 Secretary', '', '', 'Approved', 'Approved', '', '', '', '', '1', '', '', 'S2-1904847'),
            ('4.1', 'LS In', 'Action', 'LS from RAN WG2: LS to SA2 and SA5 on IAB impact to CN', 'RAN WG2 (R2-1905475)', 'Rel-16', 'NR_IAB-Core, Test WI', 'This will be taken into account in handling of the related WI. Noted', 'Noted', '', '', '', '', '2', '', '', 'S2-1904889'),
            ('4.1', 'LS In', 'Action', 'LS from RAN WG2: LS to SA2 and SA5 on IAB impact to CN', 'RAN WG2 (R2-1905475)', 'Rel-16', 'NR_IAB-Core', 'This will be taken into account in handling of the related WI. Noted', 'Noted', '', '', '', '', '2', '', '', 'S2-1904889'),
            ('4.1', 'LS In', 'Information', "LS from ITU-T SG13: LS on information about consent of ITU-T Recommendation Y.3172 'Architectural framework for machine learning in future networks including IMT-2020'", 'ITU-T SG13 (SG13-LS95.)', '', '', 'Noted', 'Noted', '', '', '', '', '3', '', '', 'S2-1904946'),
            ('4.1', 'LS In', 'Action', 'LS from SA WG4: LS on TS 26.501 5G Media Streaming (5GMS); General description and architecture', 'SA WG4 (S4h190837)', 'Rel-16', '5GMSA', 'Postponed', 'Postponed', '', '', '', '', '4', '', '', 'S2-1905810'),
            ('4.1', 'LS In', 'Information', 'LS from NGMN:  NGMN 5G End-to-End Architecture Framework', 'NGMN (Liaison_NGMN_P1_to_3GPP.)', '', '', 'Noted', 'Noted', '', '', '', '', '5', '', '', 'S2-1905823'),
            ('5.1', 'LS In', 'Information', 'LS from SA WG4: Reply LS (to RAN WG2) for inclusion of Receive Only Mode MBMS service parameters in USD', 'SA WG4 (S4-190563)', 'Rel-15', 'TEI15', 'Noted in parallel session.', 'Noted', '', '', '', '', '6', '', '', 'S2-1904901'),
            ('5.1', 'LS In', 'Action', 'LS from SA WG3LI: Response LS on reporting all Cell IDs in 5G', 'SA WG3LI (S3i190265)', 'Rel-15', 'LI15', 'Response drafted in S2-1906031. Final response in S2-1906170', 'Replied to', '', '', '', '', '7', '', '', 'S2-1904945'),
            ('5.1', 'LS OUT', 'Approval', '[DRAFT] Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Created in parallel session. Response to S2-1904945. Revised in parallel session to S2-1906166.', 'Revised', '', 'S2-1906166', '', '', '8', '', '', 'S2-1906031'),
            ('5.1', 'LS OUT', 'Approval', '[DRAFT] Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Revision of S2-1906031. Revised in parallel session to S2-1906170.', 'Revised', 'S2-1906031', 'S2-1906170', '', '', '9', '', '', 'S2-1906031'),
            ('5.1', 'LS OUT', 'Approval', 'Response LS on reporting all Cell IDs in 5G', 'SA WG2', '-', 'LI15', 'Revision of S2-1906166. Agreed in parallel session. This was Block approved.', 'Approved', 'S2-1906166', '', '', '', '10', '', '', 'S2-1906031'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revised in parallel session to S2-1906028.', 'Revised', '', 'S2-1906028', '', '', '11', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513R1 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revision of S2-1905584. Revised in parallel session to S2-1906165.', 'Revised', 'S2-1905584', 'S2-1906165', '', '', '12', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3513R2 (Rel-15', ''F'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-15', 'EDCE5', 'Revision of S2-1906028. Agreed in parallel session. This was Block approved.', 'Agreed', 'S2-1906028', '', '', '', '13', '23.401', '3513', 'S2-1905584'),
            ('5.1', 'CR', 'Approval', "23.401 CR3514 (Rel-16', 'A'): Secondary Cell ID Reporting - completion and signalling efficiency", 'Vodafone', 'Rel-16', 'EDCE5', 'LATE DOC: Rx 17/05 16:00. For e-mail approval. Approved.', '"Agreed OPEN"', '', '', '', '', '14', '23.401', '3514', 'S2-1905613'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349 (Rel-15', ''F'): Location reporting of secondary cell", "'Nokia, 'Nokia Shanghai Bell", 'Rel-15', "LI15, TEI15", 'Revised in parallel session to S2-1906029.', 'Revised', '', 'S2-1906029', '', '', '15', '23.501', '1349', 'S2-1905243'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349R1 (Rel-15', ''F'): Location reporting of secondary cell", "Nokia, 'Nokia Shanghai Bell", 'Rel-15', "LI15, TEI15", 'Revision of S2-1905243. Revised in parallel session to S2-1906163.', 'Revised', 'S2-1905243', 'S2-1906163', '', 'S2-1906165', '16', '23.501', '1349', 'S2-1905243'),
            ('5.1', 'CR', 'Approval', "23.501 CR1349R2 (Rel-15', ''F'): Location Reporting of secondary cell", "Nokia, 'Nokia Shanghai Bell", 'Rel-15', "TEI15, LI15", 'Revision of S2-1906029. Agreed in parallel session. This was Block approved.', 'Agreed', 'S2-1906029', '', '', '', '17', '23.501', '1349', 'S2-1905243')
        ]
        index = [ 'S2-1904846', 'S2-1904847', 'S2-1904889-2', 'S2-1904889', 'S2-1904946', 'S2-1905810', 'S2-1905823', 'S2-1904901', 'S2-1904945', 'S2-1906031', 'S2-1906166', 'S2-1906170', 'S2-1905584', 'S2-1906028', 'S2-1906165', 'S2-1905613', 'S2-1905243', 'S2-1906029', 'S2-1906163' ]
        df = pd.DataFrame(columns=columns, data=data, index=index)
        df.index.name = 'TD#'

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(df, doc, 'TSGS2_134_Sapporo')

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_test_df_by_wi_sa134_broken_formatting(self):
        html = '134-2.html'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', html)
        tdocs_by_agenda = html_parser.tdocs_by_agenda(file_name)
        df = tdocs_by_agenda.tdocs

        # Broken between 6.14 and 6.15.1
        # ais_to_skip = ['1', '2', '2.1', '3', '4', '4.1', '5.1', '5.2', '5.3', '5.4', '6.1', '6.2', '6.3', '6.4', '6.5', '6.5.1', '6.5.2', '6.5.3', '6.5.4', '6.5.5', '6.5.6', '6.5.7', '6.5.8', '6.5.9', '6.5.10', '6.5.11', '6.6', '6.6.1', '6.6.2', '6.7', '6.7.1', '6.7.2', '6.8', '6.8.1', '6.8.2', '6.9', '6.9.1', '6.9.2', '6.11', '6.12', '6.13.1', '6.13.2', '6.19', '6.19.1', '6.19.2', '6.20', '6.20.1', '6.20.2', '6.24', '6.28', '6.28.1', '6.28.2', '6.29', '6.30', '6.5.2', '6.5.3' ]
        ais_to_output = ['6.14', '6.15', '6.15.1']

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(df, doc, 'TSGS2_134_Sapporo', ais_to_output=ais_to_output)

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_test_df_by_wi_sa134(self):
        html = '134-2.html'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', html)
        tdocs_by_agenda = html_parser.tdocs_by_agenda(file_name)
        df = tdocs_by_agenda.tdocs

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(df, doc, 'TSGS2_134_Sapporo', source='DT')

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_test_df_by_wi_sa134_template(self):
        html = '134-2.html'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', html)
        template = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'reports', 'Report_3GPP_DT_template.docx')
        tdocs_by_agenda = html_parser.tdocs_by_agenda(file_name)
        df = tdocs_by_agenda.tdocs

        doc = application.word.open_word_document(filename=template)
        word_parser.insert_doc_data_to_doc_by_wi(df, doc, 'TSGS2_134_Sapporo', source='DT', save_to_folder=os.path.dirname(os.path.realpath(__file__)))

    def test_source_diff_same(self):
        source1 = 'Ericsson, Nokia'
        source2 = 'Ericsson, Nokia'
        diff = word_parser.diff_sources(source1, source2)

        self.assertEqual(diff, '')

    def test_source_diff_add(self):
        source1 = 'Ericsson, Nokia'
        source2 = 'Ericsson, Nokia, China Mobile, AT&T'
        diff = word_parser.diff_sources(source1, source2)

        self.assertEqual(diff, '+ China Mobile, AT&T')

    def test_source_diff_change(self):
        source1 = 'Ericsson, China Mobile, Nokia'
        source2 = 'Ericsson, Nokia, China Mobile, AT&T'
        diff = word_parser.diff_sources(source1, source2)

        self.assertEqual(diff, 'Ericsson, Nokia, China Mobile, AT&T')

    @unittest.skip("Long test (~5min). Useful only to manually-generate the report")
    def test_test_df_by_wi_sa134_DT_dtocs(self):
        html = '134-2.html'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs_by_agenda', html)
        tdocs_by_agenda = html_parser.tdocs_by_agenda(file_name)
        df = tdocs_by_agenda.tdocs

        ais_to_output = ['6.14', '6.15', '6.15.1']

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(df, doc, 'TSGS2_134_Sapporo', ais_to_output=ais_to_output, source='DT')
if __name__ == '__main__':
    unittest.main()
