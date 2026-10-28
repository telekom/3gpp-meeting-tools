import unittest
import tempfile
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from src.modules.meetings.core.tdocs_parser import TDocsParser
from src.modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel

app = QApplication.instance() or QApplication([])


class TestTDocsAgendaOrderAndSubItems(unittest.TestCase):

    def setUp(self):
        self.mock_html = """<!DOCTYPE HTML>
<html>
<head><title>SA2#164 TdocsByAgenda</title></head>
<body>
<table border="1">
<tr>
    <th>AI</th>
    <th>TD#</th>
    <th>Type</th>
    <th>Doc For</th>
    <th>Subject</th>
    <th>Source</th>
    <th>Rel</th>
    <th>Work Item</th>
    <th>Comment</th>
    <th>Result</th>
</tr>
<!-- AI Header row: TD# == '-' and captured as AI Title -->
<tr>
    <td>9.1</td>
    <td>-</td>
    <td></td>
    <td></td>
    <td>System Enablers for Network Automation</td>
    <td></td>
    <td>Rel-19</td>
    <td>Session Chair</td>
    <td></td>
    <td></td>
</tr>
<!-- Sub-Item Separator 1: TD# == '-' -->
<tr>
    <td>9.1</td>
    <td>-</td>
    <td></td>
    <td></td>
    <td>Key Issue 1: High Availability</td>
    <td></td>
    <td></td>
    <td>-</td>
    <td></td>
    <td></td>
</tr>
<!-- TDoc 1 under Sub-Item 1 -->
<tr>
    <td>9.1</td>
    <td>S2-240100</td>
    <td>CR</td>
    <td>Discussion</td>
    <td>Discussion on High Availability</td>
    <td>Deutsche Telekom</td>
    <td>Rel-19</td>
    <td>FS_eNA_Ph4</td>
    <td>Reviewed by chair, needs revision</td>
    <td>not noted</td>
</tr>
<!-- TDoc 2 under Sub-Item 1 -->
<tr>
    <td>9.1</td>
    <td>S2-240101</td>
    <td>CR</td>
    <td>Agreement</td>
    <td>Solution for HA Architecture</td>
    <td>Orange</td>
    <td>Rel-19</td>
    <td>FS_eNA_Ph4</td>
    <td>Agreed in session</td>
    <td>agreed</td>
</tr>
<!-- Sub-Item Separator 2: TD# == '-' -->
<tr>
    <td>9.1</td>
    <td>-</td>
    <td></td>
    <td></td>
    <td>Key Issue 2: Energy Efficiency</td>
    <td></td>
    <td></td>
    <td>-</td>
    <td></td>
    <td></td>
</tr>
<!-- TDoc 3 under Sub-Item 2 -->
<tr>
    <td>9.1</td>
    <td>S2-240102</td>
    <td>Discussion Paper</td>
    <td>Discussion</td>
    <td>Energy saving considerations</td>
    <td>Ericsson</td>
    <td>Rel-19</td>
    <td>FS_eNA_Ph4</td>
    <td>Postponed to Tuesday</td>
    <td>noted</td>
</tr>
</table>
</body>
</html>
"""
        self.temp_file = tempfile.NamedTemporaryFile("w", suffix=".htm", delete=False, encoding="utf-8")
        self.temp_file.write(self.mock_html)
        self.temp_file.close()

    def tearDown(self):
        try:
            if os.path.exists(self.temp_file.name):
                os.remove(self.temp_file.name)
        except Exception:
            pass

    def _init_proxy_filters(self, proxy, model):
        proxy.setTypeFilters({str(r.get("Type", "") or "").strip() for r in model._data})
        proxy.setAIFilters({str(r.get("Agenda Item", "") or "").strip() for r in model._data})
        proxy.setWIFilters(set().union(*[model.get_related_wis(i) for i in range(model.rowCount())]))
        proxy.setStatusFilters({str(r.get("TDoc Status", "") or "").strip() for r in model._data})
        proxy.setMyStatusFilters({str(r.get("My Status", "") or "").strip() for r in model._data})
        proxy.setCompanyFilters(set().union(*[r.get("_Sanitized_Companies", ["Other"]) for r in model._data]))

    def test_parse_tdocs_by_agenda_sub_items_and_comment(self):
        """Verify TDocsParser extracts Sub-Items, sequence counters, and comments from singular 'Comment'."""
        parsed = TDocsParser.parse_tdocs_by_agenda(self.temp_file.name)
        self.assertEqual(len(parsed), 3)

        doc0 = parsed.get("S2-240100")
        self.assertIsNotNone(doc0)
        self.assertEqual(doc0.get("Sub-Item"), "Key Issue 1: High Availability")
        self.assertEqual(doc0.get("_sub_item_order"), 1)
        self.assertEqual(doc0.get("_agenda_order"), 1)
        self.assertEqual(doc0.get("Comments"), "Reviewed by chair, needs revision")
        self.assertEqual(doc0.get("Result"), "not noted")

        doc1 = parsed.get("S2-240101")
        self.assertIsNotNone(doc1)
        self.assertEqual(doc1.get("Sub-Item"), "Key Issue 1: High Availability")
        self.assertEqual(doc1.get("_sub_item_order"), 1)
        self.assertEqual(doc1.get("_agenda_order"), 2)
        self.assertEqual(doc1.get("Comments"), "Agreed in session")

        doc2 = parsed.get("S2-240102")
        self.assertIsNotNone(doc2)
        self.assertEqual(doc2.get("Sub-Item"), "Key Issue 2: Energy Efficiency")
        self.assertEqual(doc2.get("_sub_item_order"), 2)
        self.assertEqual(doc2.get("_agenda_order"), 3)
        self.assertEqual(doc2.get("Comments"), "Postponed to Tuesday")

    def test_merge_agenda_data_enriches_and_preserves(self):
        """Verify merge_agenda_data attaches Sub-Item and live agenda order without losing rows."""
        initial_rows = [
            {
                "TDoc": "S2-240100",
                "Title": "Discussion on High Availability",
                "Source": "Deutsche Telekom",
                "Agenda Item": "9.1",
                "TDoc sort order within agenda item": "5.0",  # Stale Excel order
            },
            {
                "TDoc": "S2-240102",
                "Title": "Energy saving considerations",
                "Source": "Ericsson",
                "Agenda Item": "9.1",
                "TDoc sort order within agenda item": "1.0",  # Stale Excel order
            }
        ]
        model = TDocsTableModel(data=initial_rows)
        self.assertEqual(model.rowCount(), 2)

        parsed = TDocsParser.parse_tdocs_by_agenda(self.temp_file.name)
        model.merge_agenda_data(parsed)

        # S2-240101 was not in initial Excel, so it should be appended
        self.assertEqual(model.rowCount(), 3)

        # Check S2-240100 got enriched with Sub-Item and Secretary Remarks
        doc0_idx = next(i for i, r in enumerate(model._data) if r["TDoc"] == "S2-240100")
        doc0 = model._data[doc0_idx]
        self.assertEqual(doc0["Sub-Item"], "Key Issue 1: High Availability")
        self.assertEqual(doc0["_agenda_order"], 1)
        self.assertEqual(doc0["Secretary Remarks"], "Reviewed by chair, needs revision")

        # Verify Qt.UserRole + 4 returns Sub-Item for Column 10 (Agenda Item)
        ai_index = model.index(doc0_idx, 10)
        self.assertEqual(model.data(ai_index, Qt.UserRole + 4), "Key Issue 1: High Availability")
        self.assertIn("Key Issue 1: High Availability", model.data(ai_index, Qt.ToolTipRole))

    def test_agenda_sorting_order_live_html_supersedes_excel(self):
        """Verify proxy model sorts by Agenda Item -> Sub-Item -> live HTML order superseding Excel."""
        rows = [
            # In Excel, S2-240102 has order "1.0", but in live HTML it's under Sub-Item 2 with order 3
            {
                "TDoc": "S2-240102",
                "Title": "Energy saving",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 2: Energy Efficiency",
                "_agenda_order": 3,
                "_sub_item_order": 2,
                "TDoc sort order within agenda item": "1.0",
            },
            # In Excel, S2-240100 has order "5.0", but in live HTML it's under Sub-Item 1 with order 1
            {
                "TDoc": "S2-240100",
                "Title": "Discussion on HA",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 1: High Availability",
                "_agenda_order": 1,
                "_sub_item_order": 1,
                "TDoc sort order within agenda item": "5.0",
            },
            # S2-240101 is also under Sub-Item 1 with order 2
            {
                "TDoc": "S2-240101",
                "Title": "HA Architecture",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 1: High Availability",
                "_agenda_order": 2,
                "_sub_item_order": 1,
            },
            # AI 9.2 item
            {
                "TDoc": "S2-240200",
                "Title": "Other topic",
                "Agenda Item": "9.2",
                "Sub-Item": "General",
                "_agenda_order": 1,
                "_sub_item_order": 1,
            }
        ]

        model = TDocsTableModel(data=rows)
        proxy = TDocsFilterProxyModel()
        proxy.setSourceModel(model)
        self._init_proxy_filters(proxy, model)
        proxy.setSortMode("AGENDA")

        # Sort by Agenda Item column (index 10)
        proxy.sort(10, Qt.AscendingOrder)

        result_tdocs = []
        for r in range(proxy.rowCount()):
            tdoc = proxy.data(proxy.index(r, 1), Qt.DisplayRole)
            result_tdocs.append(tdoc)

        # Expected order:
        # S2-240100 (AI 9.1, Sub-Item 1, live order 1)
        # S2-240101 (AI 9.1, Sub-Item 1, live order 2)
        # S2-240102 (AI 9.1, Sub-Item 2, live order 3)
        # S2-240200 (AI 9.2, Sub-Item 1, live order 1)
        self.assertEqual(result_tdocs, ["S2-240100", "S2-240101", "S2-240102", "S2-240200"])

    def test_sub_item_filtering_and_search(self):
        """Verify filtering by Sub-Item and search text including Sub-Item tokens."""
        rows = [
            {
                "TDoc": "S2-240100",
                "Title": "Discussion on HA",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 1: High Availability",
            },
            {
                "TDoc": "S2-240101",
                "Title": "HA Architecture",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 1: High Availability",
            },
            {
                "TDoc": "S2-240102",
                "Title": "Energy saving",
                "Agenda Item": "9.1",
                "Sub-Item": "Key Issue 2: Energy Efficiency",
            }
        ]

        model = TDocsTableModel(data=rows)
        proxy = TDocsFilterProxyModel()
        proxy.setSourceModel(model)
        self._init_proxy_filters(proxy, model)

        self.assertEqual(proxy.rowCount(), 3)

        # Filter by Sub-Item "Key Issue 2: Energy Efficiency"
        proxy.setSubItemFilters(["Key Issue 2: Energy Efficiency"])
        self.assertEqual(proxy.rowCount(), 1)
        self.assertEqual(proxy.data(proxy.index(0, 1), Qt.DisplayRole), "S2-240102")

        # Clear Sub-Item filter
        proxy.setSubItemFilters([])
        self.assertEqual(proxy.rowCount(), 3)

        # Search filter matching Sub-Item string
        proxy.setGlobalFilter("Energy Efficiency")
        self.assertEqual(proxy.rowCount(), 1)
        self.assertEqual(proxy.data(proxy.index(0, 1), Qt.DisplayRole), "S2-240102")

        # Search filter matching common token
        proxy.setGlobalFilter("Key Issue")
        self.assertEqual(proxy.rowCount(), 3)


if __name__ == "__main__":
    unittest.main()
