# --- File: tests/test_spec_clause_attribution.py ---
import io
import sys
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication

from modules.specifications.core.spec_clause_attribution import (
    SpecClauseAttributionEngine,
    SpecClauseAttributionThread,
    SpecClauseDeepVerifyThread,
)


class TestSpecClauseAttribution(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def test_clean_tokens(self):
        text = "This solution addresses Key Issue #4 for Ambient IoT Protocol Architecture in TS 23.501."
        tokens = SpecClauseAttributionEngine.clean_tokens(text)
        self.assertIn("ambient", tokens)
        self.assertIn("iot", tokens)
        self.assertIn("protocol", tokens)
        self.assertIn("architecture", tokens)
        # Verify common stopwords are removed
        self.assertNotIn("this", tokens)
        self.assertNotIn("for", tokens)
        self.assertNotIn("in", tokens)

    def test_extract_key_issue_references(self):
        text = (
            "Clause 6.20: This solution resolves Key Issue #4 and addresses KI #5. "
            "Furthermore, Key Issue 4 is mentioned again along with KI12."
        )
        ki_refs = SpecClauseAttributionEngine.extract_key_issue_references(text)
        self.assertIn("KI #4", ki_refs)
        self.assertIn("KI #5", ki_refs)
        self.assertIn("KI #12", ki_refs)
        # Deduplicated count
        self.assertEqual(len(ki_refs), 3)

    def test_extract_salient_diff_keywords(self):
        diff = (
            "--- TS 23.501 v18.2.0\n"
            "+++ TS 23.501 v18.3.0\n"
            "@@ -10,3 +10,5 @@\n"
            "-The old architecture shall use standard interfaces.\n"
            "+The new architecture shall support Ambient-IoT and Npcf_SMPolicyControl.\n"
            "+Furthermore, the UE communicates via AF-data-path without FFS."
        )
        keywords = SpecClauseAttributionEngine.extract_salient_diff_keywords(diff)
        self.assertIn("Ambient-IoT", keywords)
        self.assertIn("Npcf_SMPolicyControl", keywords)
        self.assertIn("AF-data-path", keywords)
        # Removed text should not be in salient added keywords
        self.assertNotIn("standard", keywords)

    def test_mode_a_normative_ts_plenary_cr_ranking(self):
        mock_history_db = MagicMock()
        mock_history_db.get_history_for_spec.return_value = [
            {
                "cr_number": "3104",
                "revision": "1",
                "category": "F",
                "meeting": "SA#104",
                "source_tdoc": "SP-240582",
                "subject": "Correction to Service-based interfaces for Nnsacf discovery",
            },
            {
                "cr_number": "3105",
                "revision": "-",
                "category": "C",
                "meeting": "SA#104",
                "source_tdoc": "SP-240582",
                "subject": "Support for Satellite Backhaul Routing in 5GS",
            },
        ]

        engine = SpecClauseAttributionEngine(history_db=mock_history_db, search_db=None)

        results = engine.attribute_clause(
            spec_number="23.501",
            clause_number="4.2.6",
            target_version="18.3.0",
            base_version="18.2.0",
            clause_title="Service-based interfaces",
            diff_text="+Nnsacf service interface is added for discovery.",
        )

        self.assertGreaterEqual(len(results), 2)
        top = results[0]
        self.assertEqual(top["cr_number"], "3104")
        self.assertEqual(top["tdoc"], "SP-240582")
        self.assertIn("Title Overlap", " ".join(top["evidence_badges"]))
        self.assertGreater(top["score"], results[1]["score"])

    def test_mode_b_study_tr_placeholder_new_solution_ranking(self):
        """
        Tests TR Mode B where candidate p-CRs used placeholder '6.x' / 'Solution #X',
        did not know final assigned 'Solution 20', but matched via Key Issue Bridge
        and stripped title semantics.
        """
        mock_history_db = MagicMock()
        mock_history_db.get_history_for_spec.return_value = [
            {
                "cr_number": "-",
                "meeting": "SA2#175-AHE",
                "source_tdoc": "-",
                "subject": "Documented approved p-CRs at S2#175-AHE: S2-2606363, S2-2606365, S2-2606368.",
                "tdocs": [
                    {"raw_tdoc": "S2-2606363"},
                    {"raw_tdoc": "S2-2606365"},
                    {"raw_tdoc": "S2-2606368"},
                ],
            }
        ]

        mock_report_gen = MagicMock()
        mock_report_gen.resolve_tdocs_metadata.return_value = [
            {
                "tdoc": "S2-2606363",
                "title": "pCR 23.801-01: New Solution on Protocol Architecture for Ambient IoT (KI#4)",
                "source": "Qualcomm, Nokia",
                "meeting": "SA2#175-AHE",
                "agenda_item": "20.6.4",
                "ai_desc": "Key Issue #4",
                "abstract": "This contribution proposes a new protocol architecture solution resolving Key Issue #4.",
            },
            {
                "tdoc": "S2-2606365",
                "title": "pCR 23.801-01: Solution on Energy Harvesting Evaluation",
                "source": "Ericsson",
                "meeting": "SA2#175-AHE",
                "agenda_item": "20.6.1",
                "ai_desc": "Key Issue #1",
                "abstract": "Evaluation of energy harvesting.",
            },
            {
                "tdoc": "S2-2606368",
                "title": "pCR 23.801-01: Update to References",
                "source": "Huawei",
                "meeting": "SA2#175-AHE",
                "agenda_item": "2",
                "ai_desc": "References",
                "abstract": "Reference updates.",
            },
        ]

        engine = SpecClauseAttributionEngine(
            history_db=mock_history_db,
            search_db=None,
            report_gen=mock_report_gen,
        )

        results = engine.attribute_clause(
            spec_number="23.801-01",
            clause_number="6.20",
            target_version="0.8.0",
            base_version="0.7.0",
            is_new_clause=True,
            clause_title="Solution #20: Protocol Architecture for Ambient IoT",
            diff_text="+This solution addresses Key Issue #4 for Ambient IoT.",
            clause_content="This solution addresses Key Issue #4 for Ambient IoT.",
        )

        self.assertEqual(len(results), 3)
        top = results[0]
        self.assertEqual(top["tdoc"], "S2-2606363")
        self.assertEqual(top["source"], "Qualcomm, Nokia")
        self.assertGreater(top["score"], 60)
        # Check evidence badges include Key Issue and Title Overlap
        badges_str = " ".join(top["evidence_badges"])
        self.assertIn("KI #4", badges_str)
        self.assertIn("Title Overlap", badges_str)

    def test_verify_candidate_in_docx_with_track_changes(self):
        """
        Tests Stage 2 content fingerprinting with OpenXML Track Changes awareness:
        ensures <w:del> subtrees are excluded, <w:ins> text is included, and placeholder
        numbers ('6.x.1' vs '6.20.1') are normalized.
        """
        # Create an in-memory .docx file with OpenXML structure
        docx_buffer = io.BytesIO()
        document_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>'
            '<w:p><w:r><w:t>Clause 6.x.1 General Architecture for Ambient IoT</w:t></w:r></w:p>'
            '<w:p>'
            '  <w:del><w:r><w:delText>Old deleted text that should be ignored.</w:delText></w:r></w:del>'
            '  <w:ins><w:r><w:t>The Ambient IoT Device communicates directly with the Base Station using backscatter signaling.</w:t></w:r></w:ins>'
            '</w:p>'
            '</w:body>'
            '</w:document>'
        )

        with zipfile.ZipFile(docx_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("word/document.xml", document_xml)

        # Write to temporary file for path inspection
        temp_docx = Path("tests/scratch_test_sample.docx")
        temp_docx.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(temp_docx, "wb") as f:
                f.write(docx_buffer.getvalue())

            diff = (
                "--- TR 23.801-01 v0.7.0\n"
                "+++ TR 23.801-01 v0.8.0\n"
                "@@ -0,0 +1,4 @@\n"
                "+Clause 6.20.1 General Architecture for Ambient IoT\n"
                "+The Ambient IoT Device communicates directly with the Base Station using backscatter signaling.\n"
            )

            matched, pct, snippet = SpecClauseAttributionEngine.verify_candidate_in_docx(
                tdoc="S2-2606363",
                diff_text=diff,
                doc_path=temp_docx,
            )

            self.assertTrue(matched)
            self.assertGreaterEqual(pct, 90)
            self.assertIn("backscatter signaling", snippet)

        finally:
            if temp_docx.exists():
                temp_docx.unlink()

    def test_attribution_thread_execution(self):
        mock_engine = MagicMock()
        mock_engine.attribute_clause.return_value = [{"tdoc": "S2-2606363", "score": 90}]

        thread = SpecClauseAttributionThread(
            engine=mock_engine,
            spec_number="23.801-01",
            clause_number="6.20",
            target_version="0.8.0",
            base_version="0.7.0",
            clause_title="Protocol Architecture for Ambient IoT",
            diff_text="+diff line",
        )

        results_holder = []
        thread.finished.connect(lambda res: results_holder.extend(res))
        thread.run()

        self.assertEqual(len(results_holder), 1)
        self.assertEqual(results_holder[0]["tdoc"], "S2-2606363")

    def test_deep_verify_thread_execution(self):
        mock_engine = MagicMock()
        mock_engine.verify_candidate_in_docx.return_value = (True, 95, "matched snippet")

        candidates = [
            {"tdoc": "S2-2606363", "score": 75, "confidence": "High", "evidence_badges": []}
        ]

        with patch("modules.specifications.core.spec_tdoc_opener.SpecTDocOpener.resolve_local_tdoc_file") as mock_res:
            mock_res.return_value = Path("tests/dummy.docx")
            with patch.object(Path, "exists", return_value=True):
                thread = SpecClauseDeepVerifyThread(
                    engine=mock_engine,
                    candidates=candidates,
                    diff_text="+diff line",
                )

                verified_holder = []
                thread.finished.connect(lambda cands: verified_holder.extend(cands))
                thread.run()

                self.assertEqual(len(verified_holder), 1)
                self.assertTrue(verified_holder[0]["is_verified"])
                self.assertEqual(verified_holder[0]["confidence"], "Verified")
                self.assertEqual(verified_holder[0]["score"], 95)


if __name__ == "__main__":
    unittest.main()
