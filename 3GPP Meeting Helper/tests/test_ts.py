import unittest
import server
import tdoc

class Test_test_ts(unittest.TestCase):
    def test_ts_id_none(self):
        tdoc_id = None
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_empty(self):
        tdoc_id = ''
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_1(self):
        tdoc_id = 'S2'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_2(self):
        tdoc_id = '23'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_3(self):
        tdoc_id = '23.'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_4(self):
        tdoc_id = '23.4'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_5(self):
        tdoc_id = '23.45'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_6(self):
        tdoc_id = '23.345'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertTrue(is_ts)

    def test_ts_id_7(self):
        tdoc_id = '23.4565'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_8(self):
        tdoc_id = '23.456.zip'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_9(self):
        tdoc_id = '23.345-f66'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_id_10(self):
        tdoc_id = '23.345-f66.zip'
        is_ts = tdoc.is_ts(tdoc_id)
        self.assertFalse(is_ts)

    def test_ts_parse_id_none(self):
        tdoc_id = None
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_empty(self):
        tdoc_id = ''
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_1(self):
        tdoc_id = 'S2'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_2(self):
        tdoc_id = '23'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_3(self):
        tdoc_id = '23.'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_4(self):
        tdoc_id = '23.4'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_5(self):
        tdoc_id = '23.45'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_6(self):
        tdoc_id = '23.345'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, '')

    def test_ts_parse_id_6_2(self):
        tdoc_id = '23345'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, '')

    def test_ts_parse_id_7(self):
        tdoc_id = '23.4565'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_7_2(self):
        tdoc_id = '234565'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNone(parsed_ts)

    def test_ts_parse_id_8(self):
        tdoc_id = '23.456.zip'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 456)
        self.assertEqual(parsed_ts.version, '')

    def test_ts_parse_id_8_2(self):
        tdoc_id = '23456.zip'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 456)
        self.assertEqual(parsed_ts.version, '')

    def test_ts_parse_id_9(self):
        tdoc_id = '23.345-f66'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, 'f66')
        self.assertEqual(parsed_ts.match, '23.345-f66')

    def test_ts_parse_id_9_2(self):
        tdoc_id = '23345-f66'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, 'f66')
        self.assertEqual(parsed_ts.match, '23345-f66')

    def test_ts_parse_id_9_3(self):
        tdoc_id = '23345-f66.zip'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, 'f66')
        self.assertEqual(parsed_ts.match, '23345-f66.zip')

    def test_ts_parse_id_10(self):
        tdoc_id = '23.345-f66.zip'
        parsed_ts = tdoc.parse_ts_number(tdoc_id)
        self.assertIsNotNone(parsed_ts)
        self.assertEqual(parsed_ts.series, 23)
        self.assertEqual(parsed_ts.number, 345)
        self.assertEqual(parsed_ts.version, 'f66')
        self.assertEqual(parsed_ts.match, '23.345-f66.zip')

if __name__ == '__main__':
    unittest.main()
