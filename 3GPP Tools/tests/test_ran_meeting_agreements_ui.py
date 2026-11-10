# --- File: tests/test_ran_meeting_agreements_ui.py ---
"""
Comprehensive unit tests for RAN Chairman's Notes Agreements UI visualization,
table model roles, proxy filtering, and the dedicated MeetingAgreementsDialog.
"""
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

# Ensure src is on sys.path
_root = Path(__file__).resolve().parent.parent
_src = _root / "src"
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from PyQt5.QtCore import Qt, QModelIndex
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QApplication

from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel
from modules.meetings.ui.meeting_agreements_dialog import (
    MeetingAgreementsDialog,
    AgreementsTableModel,
    _clean_agreement_preview,
)


class TestRanMeetingAgreementsUI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.sample_rows = [
            {
                "TDoc": "R2-2605760",
                "Title": "Report of [POST135][020][6G] Measurements",
                "Source": "Xiaomi, Ericsson",
                "Type": "discussion",
                "For": "Rel-20",
                "Agenda Item": "9.4.2",
                "Sub-Item": "Measurements",
                "TDoc Status": "available",
                "Secretary Remarks": (
                    "🟢 [AGREEMENT]:\n"
                    "Agreements\n"
                    "Proposal b: for study of Pain point area mobility, RAN2 adopt the following two procedures:\n"
                    "Procedure 1: the UE determines that it is within the pain point area with NW assistance."
                ),
            },
            {
                "TDoc": "R2-2607118",
                "Title": "Discussion on fast multi-carrier setup and activation",
                "Source": "OPPO",
                "Type": "discussion",
                "For": "Agreement",
                "Agenda Item": "9.4.2",
                "Sub-Item": "Measurements",
                "TDoc Status": "noted",
                "Secretary Remarks": (
                    "Observation 1: The CA setup latency occurring when UEs transition...\n"
                    "Proposal 2: RAN2 to study mechanism for fast multi-carrier setup."
                ),
            },
            {
                "TDoc": "R2-2607999",
                "Title": "Working Assumption on Unified Beam Indication",
                "Source": "Qualcomm",
                "Type": "discussion",
                "For": "Decision",
                "Agenda Item": "9.3.1",
                "Sub-Item": "Beam Management",
                "TDoc Status": "available",
                "Secretary Remarks": (
                    "🟡 [WORKING ASSUMPTION]:\n"
                    "Working assumption\n"
                    "1. Support unified TCI state indication for downlink and uplink."
                ),
            },
        ]

    def test_model_agreement_detection_and_tagging(self):
        """Verifies TDocsTableModel precomputes _has_agreement and updates TDoc Status."""
        model = TDocsTableModel(data=self.sample_rows)

        # Row 0 (R2-2605760) has 🟢 [AGREEMENT]
        row0 = model._data[0]
        self.assertTrue(row0.get("_has_agreement"))
        self.assertEqual(row0.get("_agreement_kind"), "Agreement")
        # Status should be promoted from "available" to "🟢 Agreement"
        self.assertIn("🟢", row0.get("TDoc Status", ""))
        self.assertIn("Agreement", row0.get("TDoc Status", ""))

        # Row 1 (R2-2607118) has plain secretary remarks (no agreement)
        row1 = model._data[1]
        self.assertFalse(row1.get("_has_agreement", False))
        self.assertEqual(row1.get("TDoc Status"), "noted")

        # Row 2 (R2-2607999) has 🟡 [WORKING ASSUMPTION]
        row2 = model._data[2]
        self.assertTrue(row2.get("_has_agreement"))
        self.assertEqual(row2.get("_agreement_kind"), "Working Assumption")
        self.assertIn("🟡", row2.get("TDoc Status", ""))

    def test_model_background_and_foreground_roles(self):
        """Verifies soft green / amber tinting and bold green status foreground in table data()."""
        model = TDocsTableModel(data=self.sample_rows)

        # Row 0: Agreement -> soft green background #F0FDF4
        idx0_remarks = model.index(0, model._headers.index("Secretary Remarks"))
        bg0 = model.data(idx0_remarks, Qt.BackgroundRole)
        self.assertIsInstance(bg0, QColor)
        self.assertEqual(bg0.name().upper(), "#F0FDF4")

        # Row 0: TDoc Status foreground -> bold green #15803D
        idx0_status = model.index(0, model._headers.index("TDoc Status"))
        fg0 = model.data(idx0_status, Qt.ForegroundRole)
        self.assertIsInstance(fg0, QColor)
        self.assertEqual(fg0.name().upper(), "#15803D")

        # Row 1: No agreement -> background is None
        idx1_remarks = model.index(1, model._headers.index("Secretary Remarks"))
        bg1 = model.data(idx1_remarks, Qt.BackgroundRole)
        self.assertIsNone(bg1)

        # Row 2: Working Assumption -> soft amber background #FEFCE8
        idx2_remarks = model.index(2, model._headers.index("Secretary Remarks"))
        bg2 = model.data(idx2_remarks, Qt.BackgroundRole)
        self.assertIsInstance(bg2, QColor)
        self.assertEqual(bg2.name().upper(), "#FEFCE8")

    def test_secretary_remarks_pill_and_tooltip_formatting(self):
        """Verifies Secretary Remarks renders styled HTML badge and stripped preview without ugly raw tags."""
        model = TDocsTableModel(data=self.sample_rows)

        idx0_remarks = model.index(0, model._headers.index("Secretary Remarks"))
        display_html = model.data(idx0_remarks, Qt.DisplayRole)

        # Must have executive pill badge
        self.assertIn("⚖️ AGREEMENT", display_html)
        self.assertIn("#DCFCE7", display_html)
        # Must have clean preview line
        self.assertIn("Proposal b: for study of Pain point area mobility", display_html)
        # Must NOT have raw bracket tag
        self.assertNotIn("🟢 [AGREEMENT]:", display_html)
        # Must have double click hint
        self.assertIn("Double-click to view", display_html)

        # Tooltip should return rich card
        tooltip_html = model.data(idx0_remarks, Qt.ToolTipRole)
        self.assertIn("3GPP Technical Agreement", tooltip_html)
        self.assertIn("R2-2605760", tooltip_html)
        self.assertIn("AI 9.4.2", tooltip_html)

    def test_proxy_model_agreements_only_filter(self):
        """Verifies TDocsFilterProxyModel filters table down to only agreement rows."""
        model = TDocsTableModel(data=self.sample_rows)
        proxy = TDocsFilterProxyModel()
        proxy.setSourceModel(model)

        # Populate proxy active filter facets as done by CheckableComboBox in TDocsWindow
        proxy.setTypeFilters({"discussion"})
        proxy.setAIFilters({"9.4.2", "9.3.1"})
        proxy.setWIFilters({"(No WI)"})
        proxy.setStatusFilters({str(r.get("TDoc Status")).strip() for r in model._data})
        unique_comps = {c for r in model._data for c in (r.get('_Sanitized_Companies') or ['Other'])}
        proxy.setCompanyFilters(unique_comps)
        proxy.setMyStatusFilters({"⚪ Neutral", ""})

        # Unfiltered: all 3 rows visible
        self.assertEqual(proxy.rowCount(), 3)

        # Filter to agreements only: rows 0 and 2 have agreements/WAs
        proxy.setAgreementsFilter(True)
        self.assertEqual(proxy.rowCount(), 2)

        visible_tdocs = {
            proxy.data(proxy.index(r, model._headers.index("TDoc")), Qt.DisplayRole)
            for r in range(proxy.rowCount())
        }
        self.assertEqual(visible_tdocs, {"R2-2605760", "R2-2607999"})
        self.assertNotIn("R2-2607118", visible_tdocs)

        # Reset filter: back to 3 rows
        proxy.setAgreementsFilter(False)
        self.assertEqual(proxy.rowCount(), 3)

    def test_meeting_agreements_dialog(self):
        """Verifies MeetingAgreementsDialog data loading, KPI chips, searching, and markdown export."""
        agreements = [
            {
                "id": "AG-9.4.2-1",
                "ai": "9.4.2",
                "wi": "FS_6G_Radio",
                "kind": "Agreement",
                "tdoc": "R2-2605760",
                "title": "Report of [POST135][020][6G] Measurements",
                "company": "Xiaomi, Ericsson",
                "text": "Proposal b: for study of Pain point area mobility, RAN2 adopt two procedures.",
            },
            {
                "id": "AG-9.3.1-1",
                "ai": "9.3.1",
                "wi": "NR_MIMO_Rel19",
                "kind": "Working Assumption",
                "tdoc": "R2-2607999",
                "title": "Working Assumption on Unified Beam Indication",
                "company": "Qualcomm",
                "text": "1. Support unified TCI state indication for downlink and uplink.",
            },
        ]

        dlg = MeetingAgreementsDialog(agreements=agreements, meeting_name="RAN2 135BIS")

        # Verify initial model row count
        self.assertEqual(dlg.model.rowCount(), 2)

        # Test searching
        dlg.search_input.setText("Pain point")
        self.assertEqual(dlg.model.rowCount(), 1)
        self.assertEqual(dlg.model.get_row(0)["tdoc"], "R2-2605760")

        dlg.search_input.clear()
        self.assertEqual(dlg.model.rowCount(), 2)

        # Test kind filter
        dlg.combo_kind.setCurrentIndex(2)  # WA Only
        self.assertEqual(dlg.model.rowCount(), 1)
        self.assertEqual(dlg.model.get_row(0)["tdoc"], "R2-2607999")

        dlg.combo_kind.setCurrentIndex(0)  # All
        self.assertEqual(dlg.model.rowCount(), 2)

        # Test markdown builder
        md = dlg._build_markdown_content()
        self.assertIn("# RAN2 135BIS — Agreements & Working Assumptions", md)
        self.assertIn("## Agenda Item 9.4.2", md)
        self.assertIn("`R2-2605760`", md)
        self.assertIn("`FS_6G_Radio`", md)
        self.assertIn("Proposal b: for study of Pain point area mobility", md)

        # Test focus_tdoc
        dlg.focus_tdoc("R2-2607999")
        selected_indexes = dlg.table.selectionModel().selectedRows()
        self.assertTrue(len(selected_indexes) > 0)
        selected_row = dlg.model.get_row(selected_indexes[0].row())
        self.assertEqual(selected_row["tdoc"], "R2-2607999")


    def test_sa_agreed_pcr_and_cr_detection(self):
        """Verifies SA2/SA/CT agreed pCRs, CRs, and LS out are recognized as agreements."""
        sa_rows = [
            {
                'TDoc': 'S2-2601234',
                'Title': 'Solution 1 for KI #3 on Ambient IoT',
                'Source': 'Deutsche Telekom',
                'Type': 'pCR',
                'For': 'Agreement',
                'Agenda Item': '6.3',
                'TDoc Status': 'agreed',
                'Secretary Remarks': 'Agreed off-line in drafting group 2',
            },
            {
                'TDoc': 'S2-2601235',
                'Title': 'CR on 23.501 AMF selection',
                'Source': 'Orange',
                'Type': 'CR',
                'For': 'Agreement',
                'Agenda Item': '7.1',
                'TDoc Status': 'approved',
                'Secretary Remarks': '',
            },
            {
                'TDoc': 'S2-2601236',
                'Title': 'Reply LS to RAN2 on 6G Architecture',
                'Source': 'Qualcomm',
                'Type': 'LS out',
                'For': 'Approval',
                'Agenda Item': '8.2',
                'TDoc Status': 'approved',
                'Secretary Remarks': 'Approved as revised',
            },
            {
                'TDoc': 'S2-2601237',
                'Title': 'Discussion on Edge Computing issues',
                'Source': 'Nokia',
                'Type': 'discussion',
                'For': 'Discussion',
                'Agenda Item': '9.1',
                'TDoc Status': 'noted',
                'Secretary Remarks': 'Noted without conclusion',
            },
            {
                'TDoc': 'S2-2601238',
                'Title': 'Competing Solution 2 for KI #3',
                'Source': 'Huawei',
                'Type': 'pCR',
                'For': 'Agreement',
                'Agenda Item': '6.3',
                'TDoc Status': 'not agreed',
                'Secretary Remarks': 'Not agreed, superseded by S2-2601234',
            },
        ]

        model = TDocsTableModel(data=sa_rows)

        # Row 0: Agreed pCR -> agreement
        self.assertTrue(model._data[0].get('_has_agreement'))
        self.assertEqual(model._data[0].get('_agreement_kind'), 'Agreed pCR')
        bg0 = model.data(model.index(0, model._headers.index('Secretary Remarks')), Qt.BackgroundRole)
        self.assertEqual(bg0.name().upper(), '#F0FDF4')
        disp0 = model.data(model.index(0, model._headers.index('Secretary Remarks')), Qt.DisplayRole)
        self.assertIn('📝 AGREED pCR', disp0)
        self.assertIn('Solution 1 for KI #3', disp0)

        # Row 1: Approved CR -> agreement
        self.assertTrue(model._data[1].get('_has_agreement'))
        self.assertEqual(model._data[1].get('_agreement_kind'), 'Approved CR')
        disp1 = model.data(model.index(1, model._headers.index('Secretary Remarks')), Qt.DisplayRole)
        self.assertIn('📄 APPROVED CR', disp1)
        self.assertIn('CR on 23.501 AMF selection', disp1)

        # Row 2: Approved LS out -> agreement
        self.assertTrue(model._data[2].get('_has_agreement'))
        self.assertEqual(model._data[2].get('_agreement_kind'), 'Approved LS')

        # Row 3: Noted discussion paper -> NOT agreement
        self.assertFalse(model._data[3].get('_has_agreement', False))
        bg3 = model.data(model.index(3, model._headers.index('Secretary Remarks')), Qt.BackgroundRole)
        self.assertIsNone(bg3)

        # Row 4: Not agreed pCR -> NOT agreement
        self.assertFalse(model._data[4].get('_has_agreement', False))
        bg4 = model.data(model.index(4, model._headers.index('Secretary Remarks')), Qt.BackgroundRole)
        self.assertIsNone(bg4)

    def test_sa_agreements_dialog_chips_and_filtering(self):
        """Verifies MeetingAgreementsDialog presents SA pCRs and CRs with dedicated badges and filters."""
        sa_agreements = [
            {
                'id': 'SA-1',
                'ai': '6.3',
                'wi': 'FS_Ambient_IoT',
                'kind': 'Agreed pCR',
                'tdoc': 'S2-2601234',
                'title': 'Solution 1 for KI #3 on Ambient IoT',
                'company': 'Deutsche Telekom',
                'text': 'Agreed off-line in drafting group 2',
            },
            {
                'id': 'SA-2',
                'ai': '7.1',
                'wi': '5G_Core_Enh',
                'kind': 'Approved CR',
                'tdoc': 'S2-2601235',
                'title': 'CR on 23.501 AMF selection',
                'company': 'Orange',
                'text': 'Approved with technical consensus',
            },
        ]

        dlg = MeetingAgreementsDialog(agreements=sa_agreements, meeting_name='SA2 160')
        self.assertEqual(dlg.model.rowCount(), 2)

        # Check pCR display
        pcr_kind_disp = dlg.model.data(dlg.model.index(0, 0), Qt.DisplayRole)
        self.assertEqual(pcr_kind_disp, '📝 Agreed pCR')

        # Check CR display
        cr_kind_disp = dlg.model.data(dlg.model.index(1, 0), Qt.DisplayRole)
        self.assertEqual(cr_kind_disp, '📄 Approved CR')

        # Filter by PCR
        dlg.model.apply_filter(kind_filter='PCR')
        self.assertEqual(dlg.model.rowCount(), 1)
        self.assertEqual(dlg.model.get_row(0)['tdoc'], 'S2-2601234')


if __name__ == "__main__":
    unittest.main()
