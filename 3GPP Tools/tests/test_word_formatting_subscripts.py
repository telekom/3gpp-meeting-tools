"""Tests for subscript, superscript, and font formatting preservation in Word clause extraction."""

from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

from modules.meetings.core.exporters import tdoc2md
from modules.nas.core.parsing.protocol_parser_utils import _extract_p_text


class TestWordFormattingSubscripts(unittest.TestCase):
    """Verify that subscripts (e.g. K_NF), superscripts, bold, italic, and math are preserved."""

    def setUp(self):
        self.conv = tdoc2md.DocxConverter(Path("dummy.docx"), "dummy", Path("."))
        self.conv.W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    def test_subscript_preservation_for_keys(self):
        """Verify K_NF and K_SEAF subscripts are preserved in Markdown and clean plain text."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Key </w:t></w:r>
          <w:r><w:t>K</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>NF</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:t>K</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>SEAF</w:t></w:r>
        </w:p>"""
        elem = ET.fromstring(xml)

        # 1. tdoc2md Markdown rendering
        segs = []
        self.conv._walk(elem, "NORMAL", segs)
        para = self.conv._render_segments(segs)
        self.assertEqual(para.text, "Key K<sub>NF</sub> and K<sub>SEAF</sub>")
        self.assertEqual(para.plain, "Key K_NF and K_SEAF")

        # 2. protocol_parser_utils extraction (Spec Search DB / diffing)
        plain_extracted = _extract_p_text(elem)
        self.assertEqual(plain_extracted, "Key K_NF and K_SEAF")

    def test_numeric_subscripts_and_superscripts(self):
        """Verify numeric subscripts (K1) and superscripts (2^16) are preserved."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Index </w:t></w:r>
          <w:r><w:t>K</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>1</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:t>2</w:t></w:r>
          <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>16</w:t></w:r>
        </w:p>"""
        elem = ET.fromstring(xml)

        segs = []
        self.conv._walk(elem, "NORMAL", segs)
        para = self.conv._render_segments(segs)
        self.assertEqual(para.text, "Index K<sub>1</sub> and 2<sup>16</sup>")
        self.assertEqual(para.plain, "Index K_1 and 2^16")

        # In protocol_parser_utils, numeric digits use Unicode glyphs
        plain_extracted = _extract_p_text(elem)
        self.assertIn("K\u2081", plain_extracted)  # K₁
        self.assertIn("2\xb9\u2076", plain_extracted)  # 2¹⁶

    def test_office_math_omml_equations(self):
        """Verify Office Math (<m:oMath>) with subscripts and fractions is extracted properly."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                   xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t>Derive </w:t></w:r>
          <m:oMath>
            <m:sSub>
              <m:e><m:r><m:t>K</m:t></m:r></m:e>
              <m:sub><m:r><m:t>AMF</m:t></m:r></m:sub>
            </m:sSub>
          </m:oMath>
          <w:r><w:t> from </w:t></w:r>
          <m:oMath>
            <m:sSub>
              <m:e><m:r><m:t>K</m:t></m:r></m:e>
              <m:sub><m:r><m:t>SEAF</m:t></m:r></m:sub>
            </m:sSub>
          </m:oMath>
        </w:p>"""
        elem = ET.fromstring(xml)

        segs = []
        self.conv._walk(elem, "NORMAL", segs)
        para = self.conv._render_segments(segs)
        self.assertEqual(para.text, "Derive K<sub>AMF</sub> from K<sub>SEAF</sub>")
        self.assertEqual(para.plain, "Derive K_AMF from K_SEAF")

        plain_extracted = _extract_p_text(elem)
        self.assertEqual(plain_extracted, "Derive K_AMF from K_SEAF")

    def test_bold_italic_and_code_formatting(self):
        """Verify bold, italic, strikethrough, underline, and monospace code font are preserved."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:rPr><w:b/></w:rPr><w:t>shall</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:rPr><w:i/></w:rPr><w:t>defined</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:rPr><w:strike/></w:rPr><w:t>old text</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underlined</w:t></w:r>
          <w:r><w:t> and </w:t></w:r>
          <w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>ngKSI</w:t></w:r>
        </w:p>"""
        elem = ET.fromstring(xml)

        segs = []
        self.conv._walk(elem, "NORMAL", segs)
        para = self.conv._render_segments(segs)
        expected_md = "**shall** and *defined* and ~~old text~~ and <u>underlined</u> and `ngKSI`"
        self.assertEqual(para.text, expected_md)
        self.assertEqual(para.plain, "shall and defined and old text and underlined and ngKSI")

    def test_revisions_with_subscript(self):
        """Verify track changes revisions (<w:ins>/<w:del>) preserve subscripts inside revision labels."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Use key </w:t></w:r>
          <w:ins w:id="1" w:author="Editor" w:date="2026-11-02T12:00:00Z">
            <w:r><w:t>K</w:t></w:r>
            <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>NF</w:t></w:r>
          </w:ins>
        </w:p>"""
        elem = ET.fromstring(xml)

        # Markup mode
        self.conv.revision_mode = "markup"
        segs = []
        self.conv._walk(elem, "NORMAL", segs)
        para = self.conv._render_segments(segs)
        self.assertIn("[INSERTED: K<sub>NF</sub>]", para.text)
        self.assertIn("K_NF", para.plain)

        # Accept mode
        self.conv.revision_mode = "accept"
        segs_acc = []
        self.conv._walk(elem, "NORMAL", segs_acc)
        para_acc = self.conv._render_segments(segs_acc)
        self.assertEqual(para_acc.text, "Use key K<sub>NF</sub>")

    def test_heading_subscript_preservation_without_redundant_bold(self):
        """Verify headings preserve subscripts without wrapping the entire heading in redundant **."""
        xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
          <w:r><w:rPr><w:b/></w:rPr><w:t>5.2 Key Derivation for </w:t></w:r>
          <w:r><w:rPr><w:b/></w:rPr><w:t>K</w:t></w:r>
          <w:r><w:rPr><w:b/><w:vertAlign w:val="subscript"/></w:rPr><w:t>NF</w:t></w:r>
        </w:p>"""
        elem = ET.fromstring(xml)

        res = self.conv._render_para_block(elem, top=False, in_cell=False)
        self.assertEqual(res, "## 5.2 Key Derivation for K<sub>NF</sub>")

    def test_backwards_compatible_legacy_3tuple_segs(self):
        """Verify _render_segments handles legacy 3-tuples (kind, state, val) without fmt."""
        legacy_segs = [
            ("text", "NORMAL", "Hello "),
            ("text", "NORMAL", "World"),
            ("text", "INS", " New"),
        ]
        self.conv.revision_mode = "markup"
        para = self.conv._render_segments(legacy_segs)
        self.assertEqual(para.text, "Hello World[INSERTED:  New]")
        self.assertEqual(para.plain, "Hello World New")


if __name__ == "__main__":
    unittest.main()
