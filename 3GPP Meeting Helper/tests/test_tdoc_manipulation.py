import unittest
import server.tdoc as server
import tdoc.utils as tdoc

class Test_test(unittest.TestCase):
    def test_remote_tdoc_name(self):
        name = server.get_remote_filename('TSGS2_129BIS_West_Palm_Beach', 'S2-18134567')
        self.assertEqual(name, "http://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_129BIS_West_Palm_Beach/Docs/S2-18134567.zip")

    def test_remote_tdoc_name_with_Revision(self):
        name = server.get_remote_filename('TSGS2_129BIS_West_Palm_Beach', 'S2-18134567r09')
        self.assertEqual(name, "http://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_129BIS_West_Palm_Beach/Inbox/Revisions/S2-18134567r09.zip")

    def test_tdoc_id_none(self):
        tdoc_id = None
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_empty(self):
        tdoc_id = ''
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_short1(self):
        tdoc_id = 'S2'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_short2(self):
        tdoc_id = 'S2-'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_short3(self):
        tdoc_id = 'S2-1'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_short4(self):
        tdoc_id = 'S2-18'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_tdoc_id_short5(self):
        tdoc_id = 'S2-181'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertEqual(year, 2018)
        self.assertEqual(tdoc_number, 1)

    def test_tdoc_id_1800123(self):
        tdoc_id = 'S2-1800123'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertEqual(year, 2018)
        self.assertEqual(tdoc_number, 123)

    def test_tdoc_id_1800123_with_revision(self):
        tdoc_id = 'S2-1800123r12'
        year,tdoc_number,revision = tdoc.get_tdoc_year(tdoc_id, include_revision=True)
        self.assertEqual(year, 2018)
        self.assertEqual(tdoc_number, 123)
        self.assertEqual(revision, 'r12')

    def test_tdoc_id_1800123_with_wrong_revision(self):
        tdoc_id = 'S2-1800123r2'
        year,tdoc_number,revision = tdoc.get_tdoc_year(tdoc_id, include_revision=True)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)
        self.assertIsNone(revision)

    def test_tdoc_id_zip(self):
        tdoc_id = 'S2-1800123.zip'
        year,tdoc_number = tdoc.get_tdoc_year(tdoc_id)
        self.assertIsNone(year)
        self.assertIsNone(tdoc_number)

    def test_no_tdoc1(self):
        tdoc_id = None
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_no_tdoc2(self):
        tdoc_id = ''
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_no_tdoc3(self):
        tdoc_id = 'blah'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_no_tdoc4(self):
        tdoc_id = 'S2-12'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_is_tdoc1(self):
        tdoc_id = 'S2-121'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertTrue(is_tdoc)

    def test_is_tdoc2(self):
        tdoc_id = 'S2-1213'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertTrue(is_tdoc)

    def test_is_tdoc2_with_rev(self):
        tdoc_id = 'S2-1213r12'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertTrue(is_tdoc)

    def test_is_tdoc2_with_wrong_ref(self):
        tdoc_id = 'S2-1213r2'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_is_tdoc3(self):
        tdoc_id = 'S2-1213.zip'
        is_tdoc = tdoc.is_tdoc(tdoc_id)
        self.assertFalse(is_tdoc)

    def test_ts_folder(self):
        folder = server.get_ts_folder(36, 13)
        self.assertEqual(folder, 'http://www.3gpp.org/ftp/Specs/latest/Rel-13/36_series/')

if __name__ == '__main__':
    unittest.main()
