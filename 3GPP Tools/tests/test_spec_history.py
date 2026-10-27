# --- File: tests/test_spec_history.py ---
"""
Unit tests for Specification Change History, p-CR tracking, and TDoc opener engines.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add src and root to path
src_dir = Path(__file__).resolve().parent.parent / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))
root_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.core.spec_history_extractor import (
    RE_TDOC_TOKEN,
    detect_header_columns,
    normalize_version_token,
)
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener


class TestSpecHistoryExtractor(unittest.TestCase):
    """Tests for TDoc regex matching, p-CR detection, and column header normalization."""

    def test_tdoc_regex_extraction(self):
        text = (
            "Documented approved p-CRs at S2#175: S2-2605098, S2-2605238, S2-2605240. "
            "Also revision S2-2606265r01 and typo R3-261549R3-261553 were integrated."
        )
        matches = RE_TDOC_TOKEN.findall(text)
        expected = ["S2-2605098", "S2-2605238", "S2-2605240", "S2-2606265r01", "R3-261549", "R3-261553"]
        self.assertEqual(matches, expected)

    def test_normalize_version_token(self):
        self.assertEqual(normalize_version_token("v18.2.0"), "18.2.0")
        self.assertEqual(normalize_version_token("0.8.0"), "0.8.0")
        self.assertEqual(normalize_version_token("-"), "")

    def test_detect_header_columns(self):
        headers = ["Date", "Meeting", "TDoc", "CR", "Rev", "Cat", "Subject/Comment", "New version"]
        mapping = detect_header_columns(headers)
        self.assertIsNotNone(mapping)
        self.assertEqual(mapping.get("date"), 0)
        self.assertEqual(mapping.get("meeting"), 1)
        self.assertEqual(mapping.get("tdoc"), 2)
        self.assertEqual(mapping.get("cr"), 3)
        self.assertEqual(mapping.get("version"), 7)


class TestSpecHistoryDatabase(unittest.TestCase):
    """Tests for SpecHistoryDatabase storage, cross-referencing, and wipe functionality."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_spec_history.db"
        self.db = SpecHistoryDatabase(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_upsert_and_retrieve_history(self):
        records = [
            {
                "spec_number": "23.501",
                "version": "18.2.0",
                "date": "2024-03",
                "meeting": "SP-103",
                "source_tdoc": "SP-240123",
                "cr_number": "3104",
                "revision": "1",
                "category": "F",
                "subject": "Correction to AMF selection",
                "pcr_tdocs": [],
                "row_order": 1,
            },
            {
                "spec_number": "23.501",
                "version": "18.3.0",
                "date": "2024-06",
                "meeting": "SP-104",
                "source_tdoc": "SP-240582",
                "cr_number": "3210",
                "revision": "-",
                "category": "B",
                "subject": "Documented approved p-CRs: S2-2401234, S2-2401235",
                "pcr_tdocs": ["S2-2401234", "S2-2401235"],
                "row_order": 2,
            },
        ]

        inserted = self.db.upsert_history_records(records)
        self.assertEqual(inserted, 2)

        # Retrieve for version
        entries = self.db.get_history_for_spec("23.501", version="18.2.0")
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["cr_number"], "3104")
        self.assertEqual(entries[0]["source_tdoc"], "SP-240123")

        # Check p-CR extraction on second entry
        entries_v183 = self.db.get_history_for_spec("23.501", version="18.3.0")
        self.assertEqual(len(entries_v183), 1)
        tdocs_in_v183 = [t["tdoc"] for t in self.db.get_tdocs_for_spec_version("23.501", "18.3.0")]
        self.assertIn("S2-2401234", tdocs_in_v183)
        self.assertIn("S2-2401235", tdocs_in_v183)

        # Reverse lookup by TDoc
        tdoc_refs = self.db.get_tdoc_references("S2-2401234")
        self.assertEqual(len(tdoc_refs), 1)
        self.assertEqual(tdoc_refs[0]["spec_number"], "23.501")
        self.assertEqual(tdoc_refs[0]["version"], "18.3.0")

        # Summary check
        summary = self.db.get_version_summary(["23.501"])
        self.assertEqual(summary.get(("23.501", "18.2.0")), 1)
        self.assertEqual(summary.get(("23.501", "18.3.0")), 1)

        # Statistics
        stats = self.db.get_statistics()
        self.assertEqual(stats["total_records"], 2)
        self.assertEqual(stats["unique_specs"], 1)

        # Wipe check
        wiped = self.db.wipe_database()
        self.assertTrue(wiped)
        stats_after = self.db.get_statistics()
        self.assertEqual(stats_after["total_records"], 0)


class TestSpecTDocOpener(unittest.TestCase):
    """Tests for TDoc citation formatting and jump requests."""

    def test_citation_formatting(self):
        plain, html = SpecTDocOpener.format_tdoc_citation(
            tdoc_id="S2-2605098",
            spec_number="23.501",
            version="19.1.0",
            meeting="SA WG2#175",
            subject="TR skeleton for FS_6G_ARC",
        )
        self.assertIn("S2-2605098", plain)
        self.assertIn("TS 23.501", plain)
        self.assertIn("v19.1.0", plain)
        self.assertIn("SA WG2#175", plain)
        self.assertIn("https://portal.3gpp.org", html)


if __name__ == "__main__":
    unittest.main()
