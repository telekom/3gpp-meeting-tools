import unittest
import parsing.excel
import os.path

class Test_test_comments(unittest.TestCase):
    def test_read_comments(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments', 'comments.xlsx')
        df = parsing.excel.read_comments_file(file_name)

        found = df['S2-1903175']
        expected = '[Josep]: dvreznesrq23efrhr t'
        self.assertEqual(found, expected)
        found = df['S2-1903176']
        expected = '[Josep]: ewrtw t\n[Dieter]: loren ipsum'
        self.assertEqual(found, expected)
        found = df['S2-1903177']
        expected = '[Josep]: More yellow things\n[Dieter]: latin stuff\n\nMerged with Nokias (4590)'
        self.assertEqual(found, expected)
    def test_get_comments_files(self):
        dir_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments')
        comment_files = parsing.excel.get_comments_files_in_dir(dir_name)

        self.assertEqual(len(comment_files), 2)
        self.assertEqual(comment_files[0], '2019.04.09 Some Comments.xlsx')
        self.assertEqual(comment_files[1], 'comments.xlsx')
    def test_get_comments_from_dir_merge(self):
        dir_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments')
        df = parsing.excel.get_comments_from_dir(dir_name, merge_comments=True)

        self.assertEqual(len(df), 41)

        found = df['S2-1903175']
        expected = '[Josep]: dvreznesrq23efrhr t'
        self.assertEqual(found, expected)
        found = df['S2-1903176']
        expected = '[Josep]: ewrtw t\n[Dieter]: loren ipsum'
        self.assertEqual(found, expected)
        found = df['S2-1903177']
        expected = '[Josep]: Funnily enough, I would be OK with the removal of these set mentions. But the last addition should read "NF/NF service"\n[Dieter]: welches Set ist gemeint?\n\nMerged with Nokias (4590)'

        found = df['S2-190485']
        expected = 'Just a test session comment'
        self.assertEqual(found, expected)
        found = df['S2-190486']
        expected = '[Maria]: Only a comment from Maria'
        self.assertEqual(found, expected)

        found = df['S2-1903852']
        expected = 'Test\n\n[Josep]: dfeefeefefeeeee\n[Dieter]: dfaf'
        self.assertEqual(found, expected)
    def test_get_comments_from_dir_no_merge(self):
        dir_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments')
        df = parsing.excel.get_comments_from_dir(dir_name, merge_comments=False)

        self.assertEqual(len(df), 41)

        found = df['S2-1903175']
        expected = '[Josep]: dvreznesrq23efrhr t'
        self.assertEqual(found, expected)
        found = df['S2-1903176']
        expected = '[Josep]: ewrtw t\n[Dieter]: loren ipsum'
        self.assertEqual(found, expected)
        found = df['S2-1903177']
        expected = '[Josep]: Funnily enough, I would be OK with the removal of these set mentions. But the last addition should read "NF/NF service"\n[Dieter]: welches Set ist gemeint?\n\nMerged with Nokias (4590)'

        found = df['S2-190485']
        expected = 'Just a test session comment'
        self.assertEqual(found, expected)
        found = df['S2-190486']
        expected = '[Maria]: Only a comment from Maria'
        self.assertEqual(found, expected)

        found = df['S2-1903852']
        expected = '[Josep]: dfeefeefefeeeee\n[Dieter]: dfaf'
        self.assertEqual(found, expected)
    def test_get_comments_from_dir_default_merge(self):
        dir_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments')
        df = parsing.excel.get_comments_from_dir(dir_name)

        self.assertEqual(len(df), 41)

        found = df['S2-1903175']
        expected = '[Josep]: dvreznesrq23efrhr t'
        self.assertEqual(found, expected)
        found = df['S2-1903176']
        expected = '[Josep]: ewrtw t\n[Dieter]: loren ipsum'
        self.assertEqual(found, expected)

        found = df['S2-190485']
        expected = 'Just a test session comment'
        self.assertEqual(found, expected)
        found = df['S2-190486']
        expected = '[Maria]: Only a comment from Maria'
        self.assertEqual(found, expected)

        found = df['S2-1903852']
        expected = '[Josep]: dfeefeefefeeeee\n[Dieter]: dfaf'
        self.assertEqual(found, expected)

    def test_read_comments_format(self):
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments', 'comments.xlsx')
        comments = parsing.excel.read_comments_format(file_name)

        found = comments['S2-1903175']
        self.assertEqual(found[0][0], 'Josep')
        self.assertEqual(found[0][1], 'dvreznesrq23efrhr t')
        self.assertEqual(found[0][2], 'FFFFC7CE')
        self.assertEqual(found[0][3], 'FF9C0006')
        found = comments['S2-1903176']

    def test_get_comments_from_dir_default_merge_format(self):
        dir_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'comments')
        comments = parsing.excel.get_comments_from_dir_format(dir_name)

        self.assertEqual(len(comments), 41)
        comments_to_check = comments['S2-1903175']
        self.assertEqual(len(comments_to_check), 1)
        comments_to_check = comments['S2-1902958']
        self.assertEqual(len(comments_to_check), 8)

    def test_get_reddest_color(self):
        test_list = {'S2-190XXXX': [('1',None,'FF01FFFF', 'FF02FFFF'), ('2',None,'FF02FFFF', 'FF01FFFF'), ('3',None,'FFB2FFFF', 'FF00FFFF')]}
        (fg_color, text_color) = parsing.excel.get_colors_from_comments(test_list)

        self.assertEqual(fg_color['S2-190XXXX'], 'FFB2FFFF')
        self.assertEqual(text_color['S2-190XXXX'], 'FF02FFFF')
if __name__ == '__main__':
    unittest.main()
