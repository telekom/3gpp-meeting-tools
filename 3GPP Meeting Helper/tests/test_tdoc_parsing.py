import unittest
import os.path
import parsing.word as word_parser

class Test_test_tdoc_parsing(unittest.TestCase):
    def test_S2_1811605(self):
        tdoc = 'S2-1811605_S2_129_Draft_Rep_v006rm.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Draft Report of SA WG2 meetings #1298BIS')
        self.assertEqual(parsed_tdoc.source, 'Secretary of SA WG2')

    def test_S2_1811620(self):
        tdoc = 'S2-1811620_S2-1811226_R2-1816041.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Reply LS on Dynamic PLR Allocation in eVoLP')
        self.assertEqual(parsed_tdoc.source, 'RAN2')

    def test_S2_1812080(self):
        tdoc = 'S2-1812080_WID_eSBA_v1.1.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'New WID: Enhancements to the Service-Based 5G System Architecture')
        self.assertEqual(parsed_tdoc.source, 'China Mobile')

    def test_S2_1812226(self):
        tdoc = 'S2-1812226_eSBA update of solution 15_v1.2.docx'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'eSBA: Update of solution 15')
        self.assertEqual(parsed_tdoc.source, 'Huawei, Hisilicon')

    def test_S2_1812368(self):
        tdoc = 'S2-1812368 - 23.401 Rel-15 Correction of SPRC.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Correction of Serving PLMN Rate Control')
        self.assertEqual(parsed_tdoc.source, 'Huawei, HiSilicon')

    def test_S2_1812372(self):
        tdoc = 'S2-1812372 - Evaluation for solutions 1.1 and 1.2.docx'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Evaluation of solutions #1.1 and #1.2')
        self.assertEqual(parsed_tdoc.source, 'Huawei, HiSilicon')

    def test_S2_1812758(self):
        tdoc = 'S2-1812758 was S2-1811878_S5-186491 Reply LS on Data Analytics in SA WG2 v1.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'LS on slice related Data Analytics')
        self.assertEqual(parsed_tdoc.source, '3GPP SA2')

    def test_S2_1813311(self):
        tdoc = 'S2-1813311_was2736_SG_23401_r15_v1.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Corrections to Service Gap Control')
        self.assertEqual(parsed_tdoc.source, 'Ericsson, Verizon')

    def test_S2_1900429(self):
        tdoc = 'S2-1900429_VLAN_N3IWFSupport.docx'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'TS 23.501: NPN support for PLMN services via N3IWF')
        self.assertEqual(parsed_tdoc.source, 'Qualcomm Incorporated')

    def test_S2_1900585(self):
        tdoc = 'S2-1900585_eSBA 501 Adding SFSF.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_tdoc = word_parser.parse_document(file_name)
        self.assertEqual(parsed_tdoc.title, 'Adding SFSF function for eSBA')
        self.assertEqual(parsed_tdoc.source, 'ZTE')

if __name__ == '__main__':
    unittest.main()
