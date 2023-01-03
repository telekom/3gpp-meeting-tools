import unittest
import os.path
import parsing.word.pywin32 as word_parser


class Test_test_tdoc_cr_parsing(unittest.TestCase):
    def test_S2_1811605(self):
        tdoc = 'S2-1812368 - 23.401 Rel-15 Correction of SPRC.doc'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_cr = word_parser.parse_cr(file_name, use_cache=False)

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

    def test_S2_2210570(self):
        tdoc = 'S2-2210570_R18_AIMLsys_Support of QoS request for a list of UEs and reusing URLLC QoS monitoring for AIML-based services_23.502_v3.docx'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_cr = word_parser.parse_cr(file_name, use_cache=False)

        self.assertEqual(parsed_cr.Spec, '23.502')
        self.assertEqual(parsed_cr.Cr, '3627')
        self.assertEqual(parsed_cr.CurrentVersion, '17.6.0')
        self.assertEqual(parsed_cr.ProposedChangeAffectsUiic, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsMe, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsRan, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsCn, True)

        self.assertEqual(parsed_cr.Title, "Support of QoS request for a list of UEs and reusing URLLC QoS monitoring for AIML-based services")
        self.assertEqual(parsed_cr.SourceToWg, "Samsung")
        self.assertEqual(parsed_cr.Category, "B")
        self.assertIn("In order to request QoS for the AIML communication with each of the members of the group", parsed_cr.ReasonForChange)
        self.assertIn("Update AF Session setup with required QoS procedure to support QoS request for a list of UEs", parsed_cr.SummaryOfChange)
        self.assertIn("URLLC QoS monitoring mechnisam and QoS request for a list of UEs (e.g., FL member UEs) is not supported for AI/ML-based services", parsed_cr.ConsequencesIfNotApproved)
        self.assertEqual(parsed_cr.ClausesAffected, "4.15.6.6, 5.2.5.3.2, 5.2.6.9.1, 5.2.6.9.2, 5.2.6.9.3, 5.2.8.3.1, 5.2.13.2.4, 5.2.26.2.1")

    def test_S2_2209002(self):
        # This CR has merged Tables 2 and 3
        tdoc = 'S2-2209002-23501-Correction RedCap.docx'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_cr = word_parser.parse_cr(file_name, use_cache=False)

        self.assertEqual(parsed_cr.Spec, '23.501')
        self.assertEqual(parsed_cr.Cr, 'XXX')
        self.assertEqual(parsed_cr.CurrentVersion, '17.6.0')
        self.assertEqual(parsed_cr.ProposedChangeAffectsUiic, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsMe, True)
        self.assertEqual(parsed_cr.ProposedChangeAffectsRan, True)
        self.assertEqual(parsed_cr.ProposedChangeAffectsCn, False)

        self.assertEqual(parsed_cr.Title, "Correction of reference for RedCap indication from UE")
        self.assertEqual(parsed_cr.SourceToWg, "Qualcomm Inc.")
        self.assertEqual(parsed_cr.Category, "F")
        self.assertIn("In the context of Rel-17 NR RedCap, the following text was included in clause 5.41:", parsed_cr.ReasonForChange)
        self.assertIn("Replace TS 38.331 with the correct reference TS 38.300 for RAN ", parsed_cr.SummaryOfChange)
        self.assertIn("Stage 2 specification incorrectly implies UE indication of NR RedCap is in RRC", parsed_cr.ConsequencesIfNotApproved)
        self.assertEqual(parsed_cr.ClausesAffected, "5.41")

    def test_S2_2209667(self):
        # This CR has other template errors
        tdoc = 'S2-2209667-9163r02-23502-Suecr_v1.DOCX'
        file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tdocs', tdoc)
        parsed_cr = word_parser.parse_cr(file_name, use_cache=False)

        self.assertEqual(parsed_cr.Spec, '23.502')
        self.assertEqual(parsed_cr.Cr, '3598')
        self.assertEqual(parsed_cr.CurrentVersion, '17.6.0')
        self.assertEqual(parsed_cr.ProposedChangeAffectsUiic, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsMe, True)
        self.assertEqual(parsed_cr.ProposedChangeAffectsRan, False)
        self.assertEqual(parsed_cr.ProposedChangeAffectsCn, True)

        self.assertEqual(parsed_cr.Title, "Support of unavailability period")
        self.assertEqual(parsed_cr.SourceToWg, "Samsung, Qualcomm, LG Electronics, Apple, Nokia, Nokia Shanghai Bell, KDDI")
        self.assertEqual(parsed_cr.Category, "B")
        self.assertIn("SA2 has completed the FS_SUECR study and based on the conclusions of 23.700-61 TR, this CR proposes the corresponding changes", parsed_cr.ReasonForChange)
        self.assertIn("Introduce support of unavailability period feature", parsed_cr.SummaryOfChange)
        self.assertIn("Unavailability period feature is not available", parsed_cr.ConsequencesIfNotApproved)
        self.assertEqual(parsed_cr.ClausesAffected, "4.2.2.2.1, 4.2.2.2.2, 4.2.2.3.2, 4.15.3.1, 4.15.3.2.3b")