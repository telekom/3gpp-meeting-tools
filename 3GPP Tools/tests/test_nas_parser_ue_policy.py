import unittest
try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.nas_parser import NASDocxParser, TS24526_POLICY_MESSAGES
from modules.nas.core.parsing.protocol_parser_constants import RE_SPEC_NUMBER, W_NS
from modules.nas.core.parsing.protocol_parser_utils import _convert_table_to_html, _extract_tc_text


def _make_tc(text: str = "", grid_span: int = 1, v_merge: str = "") -> ET.Element:
    tc = ET.Element(f"{W_NS}tc")
    tc_pr = ET.SubElement(tc, f"{W_NS}tcPr")
    if grid_span > 1:
        gs = ET.SubElement(tc_pr, f"{W_NS}gridSpan")
        gs.set(f"{W_NS}val", str(grid_span))
    if v_merge:
        vm = ET.SubElement(tc_pr, f"{W_NS}vMerge")
        if v_merge != "continue":
            vm.set(f"{W_NS}val", v_merge)
    p = ET.SubElement(tc, f"{W_NS}p")
    r = ET.SubElement(p, f"{W_NS}r")
    t = ET.SubElement(r, f"{W_NS}t")
    t.text = text
    return tc


class TestNASParserUEPolicy(unittest.TestCase):
    def test_spec_number_regex_recognizes_24526(self):
        m = RE_SPEC_NUMBER.search("24526-j40")
        self.assertIsNotNone(m)
        self.assertEqual(f"{m.group(1)}.{m.group(2)}", "24.526")

    def test_collapse_word_half_octet_spare_row(self):
        """
        Tests Figure D.6.2.7 layout where Word splits '0 | 0 | 0 | 0' and 'Spare' across two rows
        while 'UE policy part type' and 'octet q+2' span vertically via w:vMerge.
        """
        tbl = ET.Element(f"{W_NS}tbl")
        # Row 0: bit header 8..1 + empty
        r0 = ET.SubElement(tbl, f"{W_NS}tr")
        for b in ["8", "7", "6", "5", "4", "3", "2", "1", ""]:
            r0.append(_make_tc(b))
        # Row 1: 0 0 0 0 (cols 0..3, vMerge=none), UE policy part type (cols 4..7, vMerge=restart), octet q+2 (col 8, vMerge=restart)
        r1 = ET.SubElement(tbl, f"{W_NS}tr")
        for _ in range(4):
            r1.append(_make_tc("0"))
        r1.append(_make_tc("UE policy part type", grid_span=4, v_merge="restart"))
        r1.append(_make_tc("octet q+2", grid_span=1, v_merge="restart"))
        # Row 2: Spare (cols 0..3, vMerge=none), empty vMerge=continue (cols 4..7), empty vMerge=continue (col 8)
        r2 = ET.SubElement(tbl, f"{W_NS}tr")
        r2.append(_make_tc("Spare", grid_span=4))
        r2.append(_make_tc("", grid_span=4, v_merge="continue"))
        r2.append(_make_tc("", grid_span=1, v_merge="continue"))

        html_out = _convert_table_to_html(tbl, is_figure_diagram=True)
        self.assertIn("0 Spare", html_out)
        self.assertIn("UE policy part type", html_out)
        self.assertIn("octet q+2", html_out)
        # Should be collapsed into 2 rows total (bit header row + 1 combined octet row), not 3 rows
        self.assertEqual(html_out.count("<tr>"), 2)

    def test_collapse_word_half_octet_mcc_nibble_row(self):
        """
        Tests Figure D.6.2.3 layout where Word puts empty top placeholder + 'MCC digit 1' on row 1
        and 'MCC digit 2' + vMerge=continue on row 2.
        """
        tbl = ET.Element(f"{W_NS}tbl")
        r0 = ET.SubElement(tbl, f"{W_NS}tr")
        for b in ["8", "7", "6", "5", "4", "3", "2", "1", ""]:
            r0.append(_make_tc(b))
        r1 = ET.SubElement(tbl, f"{W_NS}tr")
        r1.append(_make_tc("", grid_span=4))
        r1.append(_make_tc("MCC digit 1", grid_span=4, v_merge="restart"))
        r1.append(_make_tc("octet d+2", grid_span=1, v_merge="restart"))
        r2 = ET.SubElement(tbl, f"{W_NS}tr")
        r2.append(_make_tc("MCC digit 2", grid_span=4))
        r2.append(_make_tc("", grid_span=4, v_merge="continue"))
        r2.append(_make_tc("", grid_span=1, v_merge="continue"))

        html_out = _convert_table_to_html(tbl, is_figure_diagram=True)
        self.assertIn("MCC digit 2", html_out)
        self.assertIn("MCC digit 1", html_out)
        self.assertIn("octet d+2", html_out)
        self.assertEqual(html_out.count("<tr>"), 2)

    def test_preserve_line_breaks_in_description_table(self):
        """
        Tests that multi-line bit encodings in TS 24.526 Table 5.2.1 separated by <w:br/>
        or multiple <w:p> elements are rendered with <br/> in HTML.
        """
        tbl = ET.Element(f"{W_NS}tbl")
        r0 = ET.SubElement(tbl, f"{W_NS}tr")
        tc = ET.SubElement(r0, f"{W_NS}tc")
        p1 = ET.SubElement(tc, f"{W_NS}p")
        r1 = ET.SubElement(p1, f"{W_NS}r")
        t1 = ET.SubElement(r1, f"{W_NS}t")
        t1.text = "00000001 Match-all type"
        ET.SubElement(r1, f"{W_NS}br")
        t2 = ET.SubElement(r1, f"{W_NS}t")
        t2.text = "00001000 OS Id + OS App Id type"

        self.assertEqual(
            _extract_tc_text(tc, preserve_line_breaks=True),
            "00000001 Match-all type\n00001000 OS Id + OS App Id type",
        )
        html_out = _convert_table_to_html(tbl, is_figure_diagram=False)
        self.assertIn("00000001 Match-all type<br/>00001000 OS Id + OS App Id type", html_out)

    def test_parse_clause_8_table_rejects_ie_coding_tables(self):
        """
        Ensures Table D.6.3.1 ('Value part of the UE policy section management result information element')
        is not misclassified as a Clause 8 / D.5 message content table.
        """
        parser = NASDocxParser([])
        tbl = ET.Element(f"{W_NS}tbl")
        r0 = ET.SubElement(tbl, f"{W_NS}tr")
        r0.append(_make_tc("Value part of the UE policy section management result information element (octet 4 to k)"))
        r1 = ET.SubElement(tbl, f"{W_NS}tr")
        for col in ["8", "7", "6", "5", "4", "3", "2", "1", "", "All other values are reserved"]:
            r1.append(_make_tc(col))

        ies = parser._parse_clause_8_table(tbl)
        self.assertEqual(ies, [])

    def test_unroll_ue_policy_section_management_list(self):
        """
        Verifies that UE policy section management list (D.6.2) is unrolled into URSP, ANDSP, V2XP, ProSeP, A2XP, RSLPP.
        """
        parser = NASDocxParser([], spec_number="24.501")
        messages = [
            {
                "clause": "D.5.1.1.1",
                "message_name": "MANAGE UE POLICY COMMAND",
                "table_caption": "Table D.5.1.1.1: MANAGE UE POLICY COMMAND message content",
                "ies": [
                    {
                        "iei": "",
                        "information_element": "UE policy section management list",
                        "field_path": "UE policy section management list",
                        "depth": 0,
                        "type_reference": "UE policy section management list D.6.2",
                        "presence": "M",
                        "format": "LV-E",
                        "length": "8-65535",
                    }
                ],
            }
        ]
        ie_defs = [{"clause": "D.6.2", "ie_name": "UE policy section management list", "raw_description": "<p>Test</p>"}]
        parser._unroll_container_sub_ies(messages, ie_defs)

        ies = messages[0]["ies"]
        self.assertEqual(len(ies), 7)
        self.assertEqual(ies[1]["information_element"], "URSP (UE route selection policy)")
        self.assertEqual(ies[1]["type_reference"], "TS 24.526 5.2")
        self.assertEqual(ies[1]["depth"], 1)
        self.assertEqual(ies[2]["information_element"], "ANDSP (Access network discovery and selection policy)")
        self.assertEqual(ies[2]["type_reference"], "TS 24.526 5.3.1")

    def test_ts24526_policy_messages_structure(self):
        self.assertEqual(len(TS24526_POLICY_MESSAGES), 2)
        ursp_msg = TS24526_POLICY_MESSAGES[0]
        self.assertEqual(ursp_msg["clause"], "5.2")
        ie_names = [ie["information_element"] for ie in ursp_msg["ies"]]
        self.assertIn("Precedence value of URSP rule", ie_names)
        self.assertIn("Traffic descriptor", ie_names)
        self.assertIn("Route selection descriptor list", ie_names)
        self.assertIn("S-NSSAI type", ie_names)
        self.assertIn("DNN type", ie_names)


if __name__ == "__main__":
    unittest.main()
