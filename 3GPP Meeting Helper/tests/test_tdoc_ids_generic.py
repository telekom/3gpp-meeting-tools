import unittest

from tdoc.utils import is_generic_tdoc


class GenericTdocParsing(unittest.TestCase):
    def test_S2_1811605(self):
        tdoc = 'S2-1811605'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNotNone(parsed_tdoc)

    def test_S21811605_not_a_tdoc(self):
        tdoc = 'S21811605'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNone(parsed_tdoc)

    def test_S2_not_a_tdoc(self):
        tdoc = 'S2'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNone(parsed_tdoc)

    def test_s3i230505(self):
        tdoc = 's3i230505'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNotNone(parsed_tdoc)

    def test_s3i_not_a_tdoc(self):
        tdoc = 's3i'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNone(parsed_tdoc)

    def test_S3_240014(self):
        tdoc = 'S3-240014'
        parsed_tdoc = is_generic_tdoc(tdoc)
        self.assertIsNotNone(parsed_tdoc)


if __name__ == '__main__':
    unittest.main()
