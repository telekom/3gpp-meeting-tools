import unittest
import os.path
import parsing.word.pywin32 as word_parser


class Test_test_tdoc_cr_parsing(unittest.TestCase):
    def test_S2_1811605(self):
        tdoc = 'S2-1812368 - 23.401 Rel-15 Correction of SPRC.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_cr = word_parser.parse_cr(file_name)

        self.assertEqual(parsed_cr.Spec, '23.401')
        self.assertEqual(parsed_cr.Cr, '3485')
        self.assertEqual(parsed_cr.CurrentVersion, '15.5.0')
        self.assertEqual(parsed_cr.ProposedChangeAffectsUiic, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsMe, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsRan, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsCn, True)

        self.assertEqual(parsed_cr.Title, "Correction of Serving PLMN Rate Control")
        self.assertEqual(parsed_cr.SourceToWg, "Huawei, HiSilicon")
        self.assertEqual(parsed_cr.Category, "F")
        self.assertIn("Control is enforced per UE", parsed_cr.ReasonForChange)
        self.assertIn("Control in the MME is enforced per UE", parsed_cr.SummaryOfChange)
        self.assertIn("starting time between the MME and the UE/PDN GW/SCEF, and as a result leading unexpected PDU dropping/delaying at the MME", parsed_cr.ConsequencesIfNotApproved)
        self.assertEqual(parsed_cr.ClausesAffected, "4.7.7.2, 5.7.2")