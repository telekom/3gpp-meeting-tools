# --- File: tests/test_spec_origin_cr_ranking.py ---
"""
Unit and Integration Tests for 3GPP Origin CR Ranking & Affected Clause Attribution.

Verifies:
  1. Headless OpenXML CR cover sheet clause and category extraction.
  2. Database schema migration for 'affected_clauses' in spec_change_history and tdoc_metadata_cache.
  3. Multi-signal ranking matrix identifying foundational origin CRs (e.g. CR 2679 for TS 23.501 NSAC).
  4. Non-exclusive affected clauses scoring (additive boost without penalizing missing metadata).
  5. SpecVersionHistoryDialog pre-setting search filter, clause filter, and relevance sorting.
"""

import io
import os
import sqlite3
import sys
import unittest
import zipfile
from pathlib import Path

# Ensure 'src' is on sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT))

from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.core.spec_history_extractor import (
    extract_cr_cover_sheet_metadata,
    SpecHistoryExtractor,
)
from modules.specifications.core.spec_clause_attribution import (
    SpecClauseAttributionEngine,
    SpecFeatureOriginRankingThread,
)


class TestCoverSheetExtractor(unittest.TestCase):
    """Tests headless OpenXML extraction of CR cover sheet metadata."""

    def setUp(self):
        self.temp_dir = PROJECT_ROOT / "output" / "test_scratch"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.test_docx = self.temp_dir / "test_cr_cover.docx"

        # Create a synthetic 3GPP CR .docx with cover sheet XML
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Network Slice Admission Control Function (NSACF) definition</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Spec Number:</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>23.501</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>CR #:</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>2679</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Category:</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Clauses affected:</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>4.2.3, 5.15.11, 5.15.11.1, 6.3.24</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
          </w:body>
        </w:document>
        """
        with zipfile.ZipFile(self.test_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", doc_xml)
            zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')

    def tearDown(self):
        if self.test_docx.exists():
            try:
                self.test_docx.unlink()
            except Exception:
                pass

    def test_extract_cr_cover_sheet_metadata(self):
        meta = extract_cr_cover_sheet_metadata(self.test_docx)
        self.assertEqual(meta.get("category"), "B")
        self.assertEqual(meta.get("cr_number"), "2679")
        self.assertEqual(meta.get("spec_number"), "23.501")
        self.assertIn("5.15.11", meta.get("affected_clauses", []))
        self.assertIn("4.2.3", meta.get("affected_clauses", []))
        self.assertIn("6.3.24", meta.get("affected_clauses", []))
        self.assertIn("4.2.3, 5.15.11", meta.get("affected_clauses_text", ""))

    def test_extract_cr_from_zip_package(self):
        # Create a container zip (like an FTP download package) that wraps the docx
        wrapper_zip = self.temp_dir / "SP-210083.zip"
        with zipfile.ZipFile(wrapper_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(self.test_docx, arcname="SP-210083.docx")

        meta = extract_cr_cover_sheet_metadata(wrapper_zip)
        self.assertEqual(meta.get("category"), "B")
        self.assertEqual(meta.get("cr_number"), "2679")
        self.assertIn("5.15.11", meta.get("affected_clauses", []))

        if wrapper_zip.exists():
            wrapper_zip.unlink()


class TestSpecHistoryDbMigration(unittest.TestCase):
    """Tests database migration and affected_clauses storage in 3gpp_spec_history.db."""

    def setUp(self):
        self.temp_db_path = PROJECT_ROOT / "output" / "test_scratch" / "test_hist_mig.db"
        self.temp_db_path.parent.mkdir(parents=True, exist_ok=True)
        if self.temp_db_path.exists():
            self.temp_db_path.unlink()

        self.db = SpecHistoryDatabase(self.temp_db_path)

    def tearDown(self):
        if self.temp_db_path.exists():
            try:
                self.temp_db_path.unlink()
            except Exception:
                pass

    def test_affected_clauses_schema_columns(self):
        with self.db._get_connection() as conn:
            hist_cols = [c[1] for c in conn.execute("PRAGMA table_info(spec_change_history)").fetchall()]
            tdoc_cols = [c[1] for c in conn.execute("PRAGMA table_info(tdoc_metadata_cache)").fetchall()]
            self.assertIn("affected_clauses", hist_cols)
            self.assertIn("affected_clauses", tdoc_cols)

    def test_upsert_and_retrieve_affected_clauses(self):
        records = [{
            "spec_number": "23.501",
            "version": "17.0.0",
            "row_order": 1,
            "date": "2021-03",
            "meeting": "SP#91E",
            "source_tdoc": "SP-210083",
            "cr_number": "2679",
            "revision": "1",
            "category": "B",
            "subject": "TS23.501 KI#1 Network Slice Admission Control Function (NSACF) definition",
            "affected_clauses": "4.2.3, 5.15.11, 6.3.24",
        }]
        count = self.db.upsert_history_records(records)
        self.assertEqual(count, 1)

        fetched = self.db.get_spec_history("23.501", version="17.0.0")
        self.assertEqual(len(fetched), 1)
        self.assertEqual(fetched[0]["affected_clauses"], "4.2.3, 5.15.11, 6.3.24")
        self.assertEqual(fetched[0]["cr_number"], "2679")

        # Update affected clauses helper
        hist_id = fetched[0]["id"]
        ok = self.db.update_history_affected_clauses(hist_id, "5.15.11, 5.15.11.1")
        self.assertTrue(ok)
        refetched = self.db.get_spec_history("23.501", version="17.0.0")
        self.assertEqual(refetched[0]["affected_clauses"], "5.15.11, 5.15.11.1")


class TestOriginCrRankingEngine(unittest.TestCase):
    """Tests multi-signal ranking on TS 23.501 NSAC and synthetic test beds."""

    def setUp(self):
        self.real_db_path = PROJECT_ROOT / "src" / "db" / "3gpp_spec_history.db"
        if not self.real_db_path.exists():
            self.real_db_path = PROJECT_ROOT / "db" / "3gpp_spec_history.db"
        self.db = SpecHistoryDatabase(self.real_db_path)
        self.engine = SpecClauseAttributionEngine(history_db=self.db)

    def test_rank_origin_cr_nsac_real_database(self):
        """Verifies empirical ranking on TS 23.501 NSAC: CR 2679 must be ranked #1 as Foundational Origin CR."""
        if not self.db.has_spec_history("23.501"):
            self.skipTest("TS 23.501 history not in test environment database.")

        ranked = self.engine.rank_origin_crs_for_query(
            spec_number="23.501",
            query="NSAC",
            affected_clauses_hint=["5.15.11", "5.15.11.1"]
        )
        self.assertTrue(len(ranked) > 0)

        # Top candidate must be CR 2679
        top_cr = ranked[0]
        self.assertEqual(top_cr["cr_number"], "2679")
        self.assertEqual(top_cr["category"], "B")
        self.assertEqual(top_cr["version"], "17.0.0")
        self.assertEqual(top_cr["source_tdoc"], "SP-210083")
        self.assertEqual(top_cr["role"], "🎯 Foundational Origin CR")
        self.assertGreaterEqual(top_cr["score"], 70)

        # Verify evidence badges
        evidence_str = " ".join(top_cr["evidence_badges"])
        self.assertIn("Category B", evidence_str)
        self.assertIn("Definition", evidence_str)

    def test_non_exclusive_affected_clauses_synthetic(self):
        """Verifies that missing affected clauses metadata never penalizes or excludes valid CRs."""
        temp_db_path = PROJECT_ROOT / "output" / "test_scratch" / "test_non_exclusive.db"
        temp_db_path.parent.mkdir(parents=True, exist_ok=True)
        if temp_db_path.exists():
            temp_db_path.unlink()

        test_db = SpecHistoryDatabase(temp_db_path)
        records = [
            # CR 1: Has query match + Cat B, but NO affected clauses (typical real-world cover sheet omission)
            {
                "spec_number": "29.999",
                "version": "1.0.0",
                "row_order": 1,
                "date": "2022-01",
                "meeting": "CT#95",
                "source_tdoc": "CP-220001",
                "cr_number": "100",
                "revision": "0",
                "category": "B",
                "subject": "Introduction of Advanced Feature X (AFX) definition",
                "affected_clauses": "",  # Empty!
            },
            # CR 2: Has affected clauses match, but is Cat F (essential correction) later
            {
                "spec_number": "29.999",
                "version": "1.2.0",
                "row_order": 2,
                "date": "2022-06",
                "meeting": "CT#96",
                "source_tdoc": "CP-220100",
                "cr_number": "105",
                "revision": "1",
                "category": "F",
                "subject": "Correction of AFX parameter handling",
                "affected_clauses": "5.2.1",
            }
        ]
        test_db.upsert_history_records(records)
        engine = SpecClauseAttributionEngine(history_db=test_db)

        ranked = engine.rank_origin_crs_for_query(
            spec_number="29.999",
            query="AFX",
            affected_clauses_hint=["5.2.1"]
        )
        self.assertEqual(len(ranked), 2)
        # CR 100 (Cat B, Introduction) must be ranked #1 despite empty affected_clauses
        self.assertEqual(ranked[0]["cr_number"], "100")
        self.assertEqual(ranked[0]["role"], "🎯 Foundational Origin CR")

        # CR 105 must be identified as correction
        self.assertEqual(ranked[1]["cr_number"], "105")
        self.assertEqual(ranked[1]["role"], "🔧 Maintenance / Correction")

        if temp_db_path.exists():
            temp_db_path.unlink()


class TestHistoryDialogIntegration(unittest.TestCase):
    """Tests SpecVersionHistoryDialog pre-setting and relevance ranking."""

    @classmethod
    def setUpClass(cls):
        from PyQt5.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_dialog_preset_filters(self):
        from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog

        dlg = SpecVersionHistoryDialog(
            spec_number="23.501",
            initial_search="NSAC",
            initial_clause="5.15.11",
            sort_by_origin_relevance=True,
        )

        self.assertEqual(dlg.search_input.text(), "NSAC")
        self.assertEqual(dlg.clause_input.text(), "5.15.11")
        self.assertTrue(dlg.chk_clause_filter.isChecked())
        self.assertTrue(dlg.chk_origin_ranking.isChecked())
        # All Versions must be selected for origin ranking
        self.assertEqual(dlg.version_combo.currentIndex(), 0)

        # Verify table has 9 columns with Relevance when origin ranking is active
        self.assertEqual(dlg.table.columnCount(), 9)
        self.assertEqual(dlg.table.horizontalHeaderItem(0).text(), "Relevance")

        # Top item should be CR 2679
        if dlg.table.rowCount() > 0:
            cr_item = dlg.table.item(0, 5)  # CR # column (shifted by 1)
            self.assertEqual(cr_item.text(), "2679")

        dlg.close()


if __name__ == "__main__":
    unittest.main()
