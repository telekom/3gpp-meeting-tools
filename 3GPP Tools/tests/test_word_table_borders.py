import io
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from lxml import etree

from modules.word_tools.core.styles_3gpp import (
    W_NS,
    TABLE_GRID_XML,
    promote_3gpp_styles_in_xml,
    merge_and_promote_3gpp_styles,
    ensure_3gpp_package_styles,
)


class TestWordTableBorders(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="test_tbl_borders_")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_promote_styles_injects_table_grid(self):
        """Verify that promote_3gpp_styles_in_xml ensures TableGrid exists with borders."""
        raw_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:styleId="Normal">
                <w:name w:val="Normal"/>
            </w:style>
        </w:styles>"""

        promoted = promote_3gpp_styles_in_xml(raw_xml)
        root = etree.fromstring(promoted)
        ns = {"w": W_NS}
        tg = root.xpath(".//w:style[@w:styleId='TableGrid']", namespaces=ns)
        self.assertEqual(len(tg), 1)
        tbl_pr = tg[0].find("w:tblPr", ns)
        self.assertIsNotNone(tbl_pr)
        tbl_borders = tbl_pr.find("w:tblBorders", ns)
        self.assertIsNotNone(tbl_borders)
        self.assertIsNotNone(tbl_borders.find("w:top", ns))
        self.assertIsNotNone(tbl_borders.find("w:insideH", ns))
        self.assertIsNotNone(tbl_borders.find("w:insideV", ns))

    def test_merge_and_promote_preserves_existing_table_styles(self):
        """Verify that merge_and_promote_3gpp_styles preserves custom table styles and TableGrid."""
        target_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="table" w:styleId="CustomTableStyle">
                <w:name w:val="Custom Table Style"/>
            </w:style>
            <w:style w:type="table" w:styleId="TableGrid">
                <w:name w:val="Table Grid"/>
                <w:tblPr>
                    <w:tblBorders>
                        <w:top w:val="single" w:sz="6"/>
                    </w:tblBorders>
                </w:tblPr>
            </w:style>
        </w:styles>"""

        template_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:styleId="TH">
                <w:name w:val="TH"/>
            </w:style>
        </w:styles>"""

        merged = merge_and_promote_3gpp_styles(target_xml, template_xml)
        root = etree.fromstring(merged)
        ns = {"w": W_NS}

        # Custom table style preserved
        custom = root.xpath(".//w:style[@w:styleId='CustomTableStyle']", namespaces=ns)
        self.assertEqual(len(custom), 1)

        # Existing TableGrid preserved
        tg = root.xpath(".//w:style[@w:styleId='TableGrid']", namespaces=ns)
        self.assertEqual(len(tg), 1)

        # 3GPP TH style added and promoted
        th = root.xpath(".//w:style[@w:styleId='TH']", namespaces=ns)
        self.assertEqual(len(th), 1)
        self.assertIsNotNone(th[0].find("w:qFormat", ns))

    def test_ensure_3gpp_package_styles_non_destructive(self):
        """Verify ensure_3gpp_package_styles on a .docx package preserves TableGrid and themes."""
        docx_path = Path(self.temp_dir) / "test_doc.docx"

        # Build a minimal valid DOCX package
        styles_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="table" w:styleId="TableGrid">
                <w:name w:val="Table Grid"/>
                <w:tblPr>
                    <w:tblBorders>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    </w:tblBorders>
                </w:tblPr>
            </w:style>
        </w:styles>"""

        theme_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CustomTheme"/>"""

        with zipfile.ZipFile(docx_path, "w") as z:
            z.writestr("word/styles.xml", styles_xml)
            z.writestr("word/theme/theme1.xml", theme_xml)
            z.writestr("word/document.xml", b"<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>")

        ensure_3gpp_package_styles(docx_path)

        with zipfile.ZipFile(docx_path, "r") as z:
            out_styles = z.read("word/styles.xml")
            out_theme = z.read("word/theme/theme1.xml")

            root = etree.fromstring(out_styles)
            ns = {"w": W_NS}
            tg = root.xpath(".//w:style[@w:styleId='TableGrid']", namespaces=ns)
            self.assertEqual(len(tg), 1, "TableGrid must NOT be destroyed by ensure_3gpp_package_styles")

            # Theme must be preserved
            self.assertEqual(out_theme, theme_xml, "Existing theme must NOT be overwritten")


if __name__ == "__main__":
    unittest.main()
