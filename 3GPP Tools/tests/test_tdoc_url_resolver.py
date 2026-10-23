import os
import shutil
import tempfile
import unittest
from pathlib import Path

from core.utils.paths import to_extended_path
from modules.meetings.core.tdoc_url_resolver import (
    normalize_3gpp_url,
    is_tdoc_revision,
    resolve_tdoc_download_url,
)
from modules.meetings.core.tdoc_markdown_converter import (
    write_index,
    ExportedDoc,
    CONVERTER_VERSION,
)


class TestTDocUrlResolver(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="3gpp_test_url_"))

    def tearDown(self):
        shutil.rmtree(to_extended_path(self.tmp_dir), ignore_errors=True)

    def test_normalize_3gpp_url(self):
        # Empty
        self.assertEqual(normalize_3gpp_url(""), "")
        self.assertEqual(normalize_3gpp_url(None), "")

        # FTP protocol conversion
        ftp_url = "ftp://ftp.3gpp.org/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs"
        self.assertEqual(
            normalize_3gpp_url(ftp_url),
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs",
        )

        # HTTP conversion to HTTPS
        http_url = "http://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs"
        self.assertEqual(
            normalize_3gpp_url(http_url),
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs",
        )

        # Relative path conversion
        rel_url = "tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs"
        self.assertEqual(
            normalize_3gpp_url(rel_url),
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs",
        )

        # Trailing slash removal
        trailing = "https://www.3gpp.org/ftp/tsg_sa/Docs/"
        self.assertEqual(
            normalize_3gpp_url(trailing),
            "https://www.3gpp.org/ftp/tsg_sa/Docs",
        )

    def test_is_tdoc_revision(self):
        self.assertFalse(is_tdoc_revision("S2-2407123"))
        self.assertFalse(is_tdoc_revision("R2-2300001"))
        self.assertFalse(is_tdoc_revision(""))

        self.assertTrue(is_tdoc_revision("S2-2407123r01"))
        self.assertTrue(is_tdoc_revision("S2-2407123r1"))
        self.assertTrue(is_tdoc_revision("S2-2407123rev2"))
        self.assertTrue(is_tdoc_revision("S2-2407123_r02"))
        self.assertTrue(is_tdoc_revision("S2-2407123-rev3b"))

    def test_resolve_tdoc_download_url(self):
        docs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs"
        revs_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/INBOX/Revisions"

        # 1. Base TDoc routes to Docs
        u1 = resolve_tdoc_download_url("S2-2407123", docs_url=docs_url, revisions_url=revs_url)
        self.assertEqual(
            u1,
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs/S2-2407123.zip",
        )

        # 2. Revision TDoc routes to INBOX/Revisions
        u2 = resolve_tdoc_download_url("S2-2407123r01", docs_url=docs_url, revisions_url=revs_url)
        self.assertEqual(
            u2,
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/INBOX/Revisions/S2-2407123r01.zip",
        )

        # 3. Revision TDoc with only docs_url derives INBOX/Revisions from meeting base
        u3 = resolve_tdoc_download_url("S2-2407123r01", docs_url=docs_url)
        self.assertEqual(
            u3,
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/INBOX/Revisions/S2-2407123r01.zip",
        )

        # 4. Meta dict resolution
        meta = {
            "docs_folder_url": "tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs",
        }
        u4 = resolve_tdoc_download_url("S2-2407123", meta=meta)
        self.assertEqual(
            u4,
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs/S2-2407123.zip",
        )

        # 5. Explicit pre-existing URL in meta
        meta_explicit = {
            "3gpp_url": "https://www.3gpp.org/ftp/custom/S2-2407123.zip"
        }
        u5 = resolve_tdoc_download_url("S2-2407123", meta=meta_explicit)
        self.assertEqual(u5, "https://www.3gpp.org/ftp/custom/S2-2407123.zip")

    def test_write_index_manifest_contains_3gpp_links(self):
        docs = [
            ExportedDoc(
                tdoc_id="S2-2401000",
                name="S2-2401000",
                md_path=self.tmp_dir / "S2-2401000" / "S2-2401000.md",
                text="",
                meta={"docs_folder_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs"},
                doc_type="CR",
                agenda_item="8.2",
                title="Architecture Enhancements",
                status="Agreed",
                images=1,
                warnings=[],
            )
        ]
        index_path = write_index(self.tmp_dir, docs)
        self.assertTrue(index_path.exists())
        content = index_path.read_text(encoding="utf-8")

        self.assertIn("| 3GPP Server |", content)
        self.assertIn("https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_164_Maastricht/Docs/S2-2401000.zip", content)
        self.assertIn("[📥 Download ZIP]", content)


if __name__ == "__main__":
    unittest.main()
