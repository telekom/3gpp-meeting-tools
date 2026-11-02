# --- File: tests/test_pcr_formatting.py ---
"""
Unit tests for authentic 3GPP pseudo-CR (pCR) base document generation.
Verifies exact typography, hanging indent (1985 twips), underline borders,
italic abstract, and boxed blue 14pt change separators matching S3-263735
and the official 3GPP template (Pseudo-CR-Form-v0-1.docx).
"""

import io
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from lxml import etree

from modules.word_tools.core.cr_pcr_generator import (
    ChangeBlock,
    ContributionMetadata,
    build_pcr_base_document,
    W_NS,
)
from modules.word_tools.core.styles_3gpp import Style3GPP


class TestPcrFormatting(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="test_pcr_fmt_")
        self.target_path = Path(self.temp_dir) / "test_pcr_base.docx"

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _generate_doc(self, meta: ContributionMetadata) -> etree._Element:
        block = ChangeBlock(
            index=0,
            source_path="dummy.docx",
            ranges=[(0, 10)],
            headings=[],
            label="First change",
        )
        build_pcr_base_document(meta, [block], self.target_path)
        self.assertTrue(self.target_path.exists())

        with zipfile.ZipFile(self.target_path, "r") as z:
            doc_xml = z.read("word/document.xml")
            return etree.fromstring(doc_xml)

    def test_header_line_1_and_line_2_formatting(self):
        """Header line 1 has CRCoverPage, 12pt bold Arial, non-italic TDoc; line 2 has underline border."""
        meta = ContributionMetadata(
            meeting_wg="SA3",
            meeting_num="131",
            meeting_dates="17-20 Nov. 2026",
            meeting_location="Calgary, Canada",
            tdoc_id="S3-263735",
            source_company="Deutsche Telekom",
            title="Update to Solution #8.1.1",
            doc_for="Approval",
            agenda_item="5.3.1",
            work_item="FS_6G_SEC",
        )
        root = self._generate_doc(meta)
        ns = {"w": W_NS}
        body = root.find("w:body", ns)
        paras = body.findall("w:p", ns)

        # Paragraph 0: Line 1 (WG and TDoc)
        p0 = paras[0]
        p0_style = p0.xpath(".//w:pStyle/@w:val", namespaces=ns)
        self.assertEqual(p0_style, [Style3GPP.CR_COVER_PAGE])
        tabs = p0.xpath(".//w:tabs/w:tab", namespaces=ns)
        self.assertTrue(any(t.get(f"{{{W_NS}}}val") == "right" and t.get(f"{{{W_NS}}}pos") == "9639" for t in tabs))

        # Check TDoc text is non-italic
        runs = p0.xpath(".//w:r", namespaces=ns)
        tdoc_run = [r for r in runs if "S3-263735" in "".join(r.xpath(".//w:t/text()", namespaces=ns))]
        self.assertEqual(len(tdoc_run), 1)
        self.assertEqual(len(tdoc_run[0].xpath(".//w:i", namespaces=ns)), 0, "TDoc number must NOT be italic")
        self.assertEqual(len(tdoc_run[0].xpath(".//w:b", namespaces=ns)), 1, "TDoc number must be bold")
        sz_vals = tdoc_run[0].xpath(".//w:sz/@w:val", namespaces=ns)
        self.assertEqual(sz_vals, ["24"], "TDoc number must be 12pt (sz=24)")

        # Paragraph 1: Line 2 (Date and Location)
        p1 = paras[1]
        p1_style = p1.xpath(".//w:pStyle/@w:val", namespaces=ns)
        self.assertEqual(p1_style, [Style3GPP.CR_COVER_PAGE])
        bottom_bdr = p1.xpath(".//w:pBdr/w:bottom", namespaces=ns)
        self.assertEqual(len(bottom_bdr), 1, "Date/Location line must have a bottom border dividing rule")
        self.assertEqual(bottom_bdr[0].get(f"{{{W_NS}}}val"), "single")

    def test_metadata_hanging_indent_and_columnar_alignment(self):
        """All metadata paragraphs must have 1985 twips hanging indent (3.5 cm) and Arial bold formatting."""
        meta = ContributionMetadata(
            meeting_wg="SA3",
            meeting_num="130",
            meeting_dates="October 12 – 16, 2026",
            meeting_location="Prague",
            tdoc_id="S3-263735",
            source_company="Deutsche Telekom",
            title="A#3 - KI#3.1 - Sol#New - Multiple NAS security termination points",
            doc_for="Approval",
            agenda_item="5.3.1.3.2",
            work_item="FS_6G_SEC",
        )
        root = self._generate_doc(meta)
        ns = {"w": W_NS}
        body = root.find("w:body", ns)
        paras = body.findall("w:p", ns)

        # Paragraphs 2 to 6 are Source, Title, Document for, Agenda item, Work Item
        for idx in range(2, 7):
            p = paras[idx]
            ind = p.xpath(".//w:ind", namespaces=ns)
            self.assertEqual(len(ind), 1, f"Paragraph {idx} must have w:ind")
            self.assertEqual(ind[0].get(f"{{{W_NS}}}left"), "1985", f"Paragraph {idx} ind left must be 1985 twips")
            self.assertEqual(ind[0].get(f"{{{W_NS}}}hanging"), "1985", f"Paragraph {idx} ind hanging must be 1985 twips")
            sp = p.xpath(".//w:spacing/@w:after", namespaces=ns)
            self.assertEqual(sp, ["120"], f"Paragraph {idx} spacing after must be 120 twips (6pt)")

    def test_abstract_italic_arial_formatting(self):
        """Abstract must be styled in Arial Italic matching S3-263735."""
        meta = ContributionMetadata(
            meeting_wg="SA3",
            meeting_num="130",
            meeting_dates="October 12 – 16, 2026",
            meeting_location="Prague",
            tdoc_id="S3-263735",
            source_company="Deutsche Telekom",
            title="Title",
            doc_for="Approval",
            agenda_item="5.3",
            work_item="FS_6G_SEC",
            abstract="This contribution provides further details for multiple NAS security termination points.",
        )
        root = self._generate_doc(meta)
        ns = {"w": W_NS}
        body = root.find("w:body", ns)
        paras = body.findall("w:p", ns)

        # Abstract is paragraph 8 (p0=Line1, p1=Line2, p2..p6=Meta, p7=Spacer, p8=Abstract)
        p_abs = paras[8]
        full_text = "".join(p_abs.xpath(".//w:t/text()", namespaces=ns))
        self.assertIn("Abstract of the contribution:", full_text)
        self.assertIn("This contribution provides further details", full_text)

        # Check all runs in abstract paragraph have <w:i/>
        runs = p_abs.xpath(".//w:r", namespaces=ns)
        for r in runs:
            i_tag = r.xpath(".//w:i", namespaces=ns)
            self.assertEqual(len(i_tag), 1, "Abstract run must be italic")
            rf = r.xpath(".//w:rFonts/@w:ascii", namespaces=ns)
            self.assertEqual(rf, ["Arial"], "Abstract run font must be Arial")

    def test_boxed_blue_change_separators(self):
        """Change separators must have 4-sided rectangular border, centered, 14pt blue Arial."""
        meta = ContributionMetadata(
            meeting_wg="SA3",
            meeting_num="130",
            title="Title",
            doc_for="Approval",
        )
        root = self._generate_doc(meta)
        ns = {"w": W_NS}
        body = root.find("w:body", ns)
        paras = body.findall("w:p", ns)

        # Find separator paragraphs
        sep_paras = []
        for p in paras:
            text = "".join(p.xpath(".//w:t/text()", namespaces=ns))
            if "* * * First Change * * * *" in text or "* * * End of Changes * * * *" in text:
                sep_paras.append((text, p))

        self.assertEqual(len(sep_paras), 2)
        for text, p in sep_paras:
            bdr = p.xpath(".//w:pBdr", namespaces=ns)
            self.assertEqual(len(bdr), 1, f"Separator '{text}' must have w:pBdr")
            for side in ["top", "left", "bottom", "right"]:
                side_el = bdr[0].find(f"{{{W_NS}}}{side}")
                self.assertIsNotNone(side_el, f"Separator '{text}' missing border side '{side}'")
                self.assertEqual(side_el.get(f"{{{W_NS}}}val"), "single")
                self.assertEqual(side_el.get(f"{{{W_NS}}}sz"), "4")

            jc = p.xpath(".//w:jc/@w:val", namespaces=ns)
            self.assertEqual(jc, ["center"], f"Separator '{text}' must be centered")

            runs = p.xpath(".//w:r", namespaces=ns)
            self.assertTrue(len(runs) > 0)
            color = runs[0].xpath(".//w:color/@w:val", namespaces=ns)
            self.assertEqual(color, ["0000FF"], f"Separator '{text}' color must be blue (0000FF)")
            sz = runs[0].xpath(".//w:sz/@w:val", namespaces=ns)
            self.assertEqual(sz, ["28"], f"Separator '{text}' font size must be 14pt (sz=28)")

    def test_xml_escaping_in_metadata(self):
        """Special XML characters (&, <, >) in user inputs must not produce XML syntax errors."""
        meta = ContributionMetadata(
            meeting_wg="SA3 & SA2",
            meeting_num="130",
            meeting_dates="Oct 12 – 16, 2026",
            meeting_location="Prague <Center>",
            tdoc_id="S3-263735",
            source_company="Deutsche Telekom & Partners",
            title="Security Architecture & Protocols <Phase 1>",
            doc_for="Approval",
            agenda_item="5.3 & 5.4",
            work_item="FS_6G_SEC",
            abstract="Abstract with & and <brackets>.",
            discussion_text="Discussion with <bold> & text.",
            proposal_intro="Proposal with <change> & text.",
        )
        # If not escaped, etree.fromstring in _generate_doc will raise XMLSyntaxError
        root = self._generate_doc(meta)
        self.assertIsNotNone(root)


if __name__ == "__main__":
    unittest.main()
