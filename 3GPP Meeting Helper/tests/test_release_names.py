import unittest

from server.specs import version_to_file_version, file_version_to_version


class MyTestCase(unittest.TestCase):
    def test_version_to_file_version1(self):
        file_version = version_to_file_version('18.2.0')
        self.assertEqual(file_version, 'i20')

    def test_version_to_file_version2(self):
        file_version = version_to_file_version('18.2.1')
        self.assertEqual(file_version, 'i21')

    def test_version_to_file_version3(self):
        file_version = version_to_file_version('17.2.1')
        self.assertEqual(file_version, 'h21')

    def test_file_version_to_version1(self):
        file_version = file_version_to_version('i20')
        self.assertEqual(file_version, '18.2.0')

    def test_file_version_to_version2(self):
        file_version = file_version_to_version('i21')
        self.assertEqual(file_version, '18.2.1')

    def test_file_version_to_version3(self):
        file_version = file_version_to_version('h21')
        self.assertEqual(file_version, '17.2.1')


if __name__ == '__main__':
    unittest.main()
