import unittest
from pathlib import Path

from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.core.parsing.xnap_asn1_parser import XnAPAsn1Parser
from modules.nas.core.parsing.ran3_asn1_parser import RAN3Asn1Parser
from modules.nas.core.parsing.rrc_asn1_parser import RRCAsn1Parser


class TestXnAPParser(unittest.TestCase):
    """Unit tests for XnAP (TS 38.423) parsing and non-regression guarantees."""

    @classmethod
    def setUpClass(cls):
        cls.xnap_docx = Path("C:/Users/colomikunoj/3GPP_Delegate_Helper/specs/38.423/38423-j30.docx")
        cls.ngap_docx = Path("C:/Users/colomikunoj/3GPP_Delegate_Helper/specs/38.413/38413-j30.docx")
        cls.rrc_docx = Path("C:/Users/colomikunoj/3GPP_Delegate_Helper/specs/38.331/38331-fu0.docx")

    def test_xnap_dispatcher_routing(self):
        """Verifies that ProtocolDocxDispatcher routes TS 38.423 to XnAPAsn1Parser."""
        if not self.xnap_docx.exists():
            self.skipTest("TS 38.423 test docx not found")
        disp = ProtocolDocxDispatcher([self.xnap_docx])
        spec_num = disp.extract_spec_number()
        self.assertEqual(spec_num, "38.423")

    def test_xnap_message_and_ie_counts(self):
        """Verifies that XnAP extracts exactly the 103 Elementary Procedure messages and does not misclassify IEs."""
        if not self.xnap_docx.exists():
            self.skipTest("TS 38.423 test docx not found")
        parser = XnAPAsn1Parser([self.xnap_docx], spec_number="38.423")
        messages, ie_defs = parser.parse()

        # Exactly 103 messages: 102 elementary procedure messages + 1 PrivateMessage
        self.assertEqual(len(messages), 103)
        self.assertGreater(len(ie_defs), 1000)

        msg_names = {m["message_name"] for m in messages}
        ie_names = {ie["ie_name"] for ie in ie_defs}

        # Key elementary procedure messages must be present
        self.assertIn("HandoverRequest", msg_names)
        self.assertIn("HandoverRequestAcknowledge", msg_names)
        self.assertIn("SNStatusTransfer", msg_names)
        self.assertIn("UEContextRelease", msg_names)
        self.assertIn("XnSetupRequest", msg_names)
        self.assertIn("XnSetupResponse", msg_names)
        self.assertIn("CPCCancel", msg_names)

        # Clause 9.2 IEs must NOT be misclassified as messages
        self.assertNotIn("AvailableBitrateReportThresholdList", msg_names)
        self.assertNotIn("IntendedTDD-DL-ULConfiguration-NR", msg_names)

        # But they MUST exist in ie_defs (All IEs / Types)
        self.assertIn("AvailableBitrateReportThresholdList", ie_names)
        self.assertIn("IntendedTDD-DL-ULConfiguration-NR", ie_names)

    def test_cpc_cancel_clause_and_direction(self):
        """Verifies CPCCancel is resolved to Clause 9.1.2.31 with correct direction and description."""
        if not self.xnap_docx.exists():
            self.skipTest("TS 38.423 test docx not found")
        parser = XnAPAsn1Parser([self.xnap_docx], spec_number="38.423")
        messages, _ = parser.parse()

        cpc = next((m for m in messages if m["message_name"] == "CPCCancel"), None)
        self.assertIsNotNone(cpc, "CPCCancel message not found")
        self.assertEqual(cpc["clause"], "9.1.2.31")
        self.assertIn("M-NG-RAN node", cpc["direction"])
        self.assertIn("S-NG-RAN node", cpc["direction"])
        self.assertTrue(len(cpc["description"]) > 20)

    def test_message_type_synthesis_clause(self):
        """Verifies MessageType (Clause 9.2.3.1) is prepended to XnAP messages."""
        if not self.xnap_docx.exists():
            self.skipTest("TS 38.423 test docx not found")
        parser = XnAPAsn1Parser([self.xnap_docx], spec_number="38.423")
        messages, _ = parser.parse()

        ho_req = next((m for m in messages if m["message_name"] == "HandoverRequest"), None)
        self.assertIsNotNone(ho_req)
        self.assertTrue(len(ho_req["ies"]) > 0)
        first_ie = ho_req["ies"][0]
        self.assertEqual(first_ie["information_element"], "MessageType")
        self.assertIn("9.2.3.1", first_ie["type_reference"])

    def test_direction_arrow_parsing(self):
        """Verifies that Wingdings arrow symbols (w:sym F0AE) are translated to right arrows."""
        if not self.xnap_docx.exists():
            self.skipTest("TS 38.423 test docx not found")
        parser = XnAPAsn1Parser([self.xnap_docx], spec_number="38.423")
        messages, _ = parser.parse()

        ho_req = next((m for m in messages if m["message_name"] == "HandoverRequest"), None)
        self.assertIsNotNone(ho_req)
        self.assertIn("→", ho_req["direction"])
        self.assertIn("source NG-RAN node", ho_req["direction"])
        self.assertIn("target NG-RAN node", ho_req["direction"])

    def test_non_regression_ngap(self):
        """Guarantees zero regression on NGAP (TS 38.413) message and IE extraction."""
        if not self.ngap_docx.exists():
            self.skipTest("TS 38.413 test docx not found")
        disp = ProtocolDocxDispatcher([self.ngap_docx])
        messages, ie_defs = disp.parse()
        self.assertEqual(len(messages), 191)
        self.assertEqual(len(ie_defs), 1314)

    def test_non_regression_rrc(self):
        """Guarantees zero regression on NR RRC (TS 38.331) message and IE extraction."""
        if not self.rrc_docx.exists():
            self.skipTest("TS 38.331 test docx not found")
        disp = ProtocolDocxDispatcher([self.rrc_docx])
        messages, ie_defs = disp.parse()
        self.assertEqual(len(messages), 46)
        self.assertEqual(len(ie_defs), 1621)


if __name__ == "__main__":
    unittest.main()
