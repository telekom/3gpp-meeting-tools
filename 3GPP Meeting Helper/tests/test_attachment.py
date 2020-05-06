import unittest
import parsing.outlook

class Test_test_attachment(unittest.TestCase):
    def test_attachment_1(self):
        text = '''Attachment:
"S2-1902881 was S2-1902754 was S2-1901674  23.501 cr secondary NSSAI_V1+mario+QC_gv_rev1.doc"
(136k) can be downloaded at:
http://list.etsi.org/scripts/wa.exe?F2=000004E4&L=3GPP_TSG_SA_WG2
        '''
        data = parsing.outlook.get_attachment_data(text)
        self.assertEqual(data.filename, 'S2-1902881 was S2-1902754 was S2-1901674  23.501 cr secondary NSSAI_V1+mario+QC_gv_rev1.doc')
        self.assertEqual(data.url, 'http://list.etsi.org/scripts/wa.exe?F2=000004E4&L=3GPP_TSG_SA_WG2') 

    def test_attachment_None(self):
        text = None

        data = parsing.outlook.get_attachment_data(text)
        self.assertIsNone(data)

    def test_attachment_Empty(self):
        text = ''

        data = parsing.outlook.get_attachment_data(text)
        self.assertIsNone(data)

if __name__ == '__main__':
    unittest.main()
