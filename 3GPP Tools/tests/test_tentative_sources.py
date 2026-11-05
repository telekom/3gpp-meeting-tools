# --- File: tests/test_tentative_sources.py ---
import sys
import unittest
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel
from modules.meetings.ui.contributions.contribution_models import ContributionResultsTableModel


class TestTentativeSources(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_question_mark_and_parenthesized_question_mark_excluded_by_default(self):
        src = "CATT, Ericsson?, Deutsche Telekom?, T-Mobile USA?, Samsung (?), Nokia, ZTE (?)"
        confirmed = CompanySanitizer.get_matching_contributors(src)
        self.assertEqual(set(confirmed), {"CATT", "Nokia"})

        all_contribs = CompanySanitizer.get_matching_contributors(src, include_tentative=True)
        self.assertEqual(
            set(all_contribs),
            {"CATT", "Nokia", "Ericsson", "Deutsche Telekom", "T-Mobile USA", "Samsung", "ZTE"},
        )

        conf_tuple, tent_tuple = CompanySanitizer.parse_contributors_detailed(src)
        self.assertEqual(set(conf_tuple), {"CATT", "Nokia"})
        self.assertEqual(
            set(tent_tuple),
            {"Ericsson", "Deutsche Telekom", "T-Mobile USA", "Samsung", "ZTE"},
        )

    def test_square_brackets_excluded_same_as_question_mark(self):
        src = "China Mobile, [Huawei, HiSilicon, Deutsche Telekom], ZTE"
        confirmed = CompanySanitizer.get_matching_contributors(src)
        self.assertEqual(set(confirmed), {"China Mobile", "ZTE"})

        conf_tuple, tent_tuple = CompanySanitizer.parse_contributors_detailed(src)
        self.assertEqual(set(conf_tuple), {"China Mobile", "ZTE"})
        self.assertEqual(set(tent_tuple), {"Huawei", "Deutsche Telekom"})

    def test_role_brackets_not_treated_as_tentative(self):
        src = "China Unicom [pen holder], CATT [Rapporteur], [Nokia, Ericsson]"
        self.assertTrue(CompanySanitizer.has_tentative_sources(src))
        conf_tuple, tent_tuple = CompanySanitizer.parse_contributors_detailed(src)
        self.assertEqual(set(conf_tuple), {"China Unicom", "CATT"})
        self.assertEqual(set(tent_tuple), {"Nokia", "Ericsson"})

        src_only_role = "China Unicom [pen holder], CATT [Rapporteur]"
        self.assertFalse(CompanySanitizer.has_tentative_sources(src_only_role))
        self.assertEqual(set(CompanySanitizer.get_matching_contributors(src_only_role)), {"China Unicom", "CATT"})

    def test_missing_comma_after_question_mark_and_corporate_suffix(self):
        src1 = "CATT, Huawei? Samsung? China Mobile Com. Corporation? TNO?"
        conf1, tent1 = CompanySanitizer.parse_contributors_detailed(src1)
        self.assertEqual(set(conf1), {"CATT"})
        self.assertEqual(set(tent1), {"Huawei", "Samsung", "China Mobile", "TNO"})

        src2 = "CATT, Samsung?, InterDigital Inc.?, Rakuten Mobile, Inc?, NVIDIA?"
        conf2, tent2 = CompanySanitizer.parse_contributors_detailed(src2)
        self.assertEqual(set(conf2), {"CATT"})
        self.assertEqual(set(tent2), {"Samsung", "Interdigital", "Rakuten", "NVIDIA"})

    def test_tokenize_reconstructs_exact_original_string(self):
        samples = [
            "AT&T, Ericsson, Nokia, Apple?, Qualcomm, Interdigital",
            "CATT, [InterDigital Inc., NEC, Nokia, OPPO, ZTE, Deutsche Telekom], ETRI",
            "Deutsche Telekom (pen holder), [Apple, AT&T, Nokia]?",
            "China Mobile, Lenovo, Ericsson,[vivo, Huawei, HiSilicon, Qualcomm Incorporated",
            "Charter Communications, CableLabs, [Tencent], Nokia, [Samsung], [Deutsche Telecom]",
        ]
        for s in samples:
            segs = CompanySanitizer.tokenize_source_segments(s)
            reconstructed = "".join(text for text, _ in segs)
            self.assertEqual(reconstructed, s)

    def test_html_formatting_and_tooltip_explanation(self):
        src = "CATT, Nokia, [Deutsche Telekom, Ericsson], Samsung?"
        html_out = CompanySanitizer.format_source_html(src)
        self.assertIn("CATT, Nokia, ", html_out)
        self.assertIn('<span style="color: #94A3B8; font-style: italic;">[Deutsche Telekom, Ericsson], Samsung?</span>', html_out)

        tooltip = CompanySanitizer.build_source_tooltip_html(src)
        self.assertIn("Confirmed Contributors (2):", tooltip)
        self.assertIn("Unconfirmed / Carry-over (3):", tooltip)
        self.assertIn("not confirmed", tooltip)
        self.assertIn("excluded from contribution counts and statistics", tooltip)

    def test_primary_contributor_respects_tentative_flag(self):
        src_tent_first = "China Mobile?, Deutsche Telekom"
        self.assertEqual(CompanySanitizer.get_primary_contributors(src_tent_first, include_tentative=False), [])
        self.assertEqual(CompanySanitizer.get_primary_contributors(src_tent_first, include_tentative=True), ["China Mobile"])

        src_bracket_first = "[China Mobile], Deutsche Telekom"
        self.assertEqual(CompanySanitizer.get_primary_contributors(src_bracket_first, include_tentative=False), [])
        self.assertEqual(CompanySanitizer.get_primary_contributors(src_bracket_first, include_tentative=True), ["China Mobile"])

        src_conf_first = "Deutsche Telekom, [China Mobile], Nokia?"
        self.assertEqual(CompanySanitizer.get_primary_contributors(src_conf_first, include_tentative=False), ["Deutsche Telekom"])

    def test_tdocs_table_and_proxy_model_tentative_filtering(self):
        rows = [
            {"TDoc": "S2-2600001", "Title": "Base pCR", "Source": "CATT, Deutsche Telekom, Nokia", "Type": "pCR", "Agenda Item": "5.1", "TDoc Status": "Revised"},
            {"TDoc": "S2-2600002", "Title": "Rev 1 with ?", "Source": "CATT, Deutsche Telekom?, Nokia", "Type": "pCR", "Agenda Item": "5.1", "TDoc Status": "Revised"},
            {"TDoc": "S2-2600003", "Title": "Rev 2 with []", "Source": "CATT, [Deutsche Telekom, Nokia]", "Type": "pCR", "Agenda Item": "5.1", "TDoc Status": "Agreed"},
        ]
        model = TDocsTableModel(data=rows)
        proxy = TDocsFilterProxyModel()
        proxy.setSourceModel(model)

        proxy.setTypeFilters({"pCR"})
        proxy.setAIFilters({"5.1"})
        proxy.setWIFilters({"(No WI)"})
        proxy.setStatusFilters({"Revised", "Agreed"})
        proxy.setMyStatusFilters({"⚪ Neutral"})

        # Filter for Deutsche Telekom (default: include_tentative_sources = False)
        proxy.setCompanyFilters({"Deutsche Telekom"})
        self.assertEqual(proxy.rowCount(), 1)
        self.assertEqual(proxy.data(proxy.index(0, 1), Qt.DisplayRole), "S2-2600001")

        # Enable include_tentative_sources -> matches all 3 rows!
        proxy.setIncludeTentativeSources(True)
        self.assertEqual(proxy.rowCount(), 3)

        # Verify tooltip on row 1 (S2-2600002) Source column (col 3)
        src_idx = model.index(1, 3)
        tooltip = model.data(src_idx, Qt.ToolTipRole)
        self.assertIsNotNone(tooltip)
        self.assertIn("Unconfirmed / Carry-over", tooltip)

    def test_contribution_results_table_model_source_tooltip(self):
        rows = [
            {"WG": "SA2", "Meeting": "SA2 #178", "TDoc": "S2-2600002", "Title": "Test", "Source": "CATT, [Nokia], Ericsson?"}
        ]
        model = ContributionResultsTableModel(rows)
        source_idx = model.index(0, 4)  # Col 4 is Source
        tooltip = model.data(source_idx, Qt.ToolTipRole)
        self.assertIsNotNone(tooltip)
        self.assertIn("Confirmed Contributors (1):", tooltip)
        self.assertIn("Unconfirmed / Carry-over (2):", tooltip)


if __name__ == "__main__":
    unittest.main()
