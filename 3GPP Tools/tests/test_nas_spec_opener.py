# --- File: tests/test_nas_spec_opener.py ---
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtWidgets import QApplication, QAction

from modules.nas.ui.nas_components import (
    NASInspectorWidget,
    NASVersionTreeWidget,
)
from modules.nas.ui.nas_tabs import NASTab
from modules.spec_search.core.spec_doc_resolver import (
    SpecDocResolver,
    letter_code_to_version,
    match_multipart_document_for_clause,
    version_to_letter_code,
)

# Ensure QApplication exists for GUI testing
app = QApplication.instance() or QApplication(sys.argv)


class TestNASSpecOpener(unittest.TestCase):
    """Verifies specification opening capabilities from the Protocols tab."""

    def setUp(self):
        self.dummy_nas_db = Path("src/db/3gpp_protocol_data.db")
        self.dummy_specs_db = Path("src/db/3gpp_data.db")

    def test_version_letter_code_bidirectional_conversion(self):
        """Verifies version_to_letter_code and letter_code_to_version symmetry."""
        cases = [
            ("20.0.0", "k00"),
            ("19.7.0", "j70"),
            ("16.20.0", "gk0"),
            ("15.7.0", "f70"),
            ("17.18.1", "hi1"),
        ]
        for ver, code in cases:
            self.assertEqual(version_to_letter_code(ver), code)
            self.assertEqual(letter_code_to_version(code), ver)

    def test_multipart_clause_matching(self):
        """Verifies multi-part specification splitting by clause and annex."""
        docs_24501 = [
            Path("24501-k00_0_cover.docx"),
            Path("24501-k00_1_Main-Body_s00_s04.docx"),
            Path("24501-k00_2_Main-Body_s05_s0504.docx"),
            Path("24501-k00_3_Main-Body_s0505_s0506.docx"),
            Path("24501-k00_4_Main-Body_s06_s08.docx"),
            Path("24501-k00_5_Main-Body_s09_s10.docx"),
            Path("24501-k00_6_Annexes_sA_sHistory.docx"),
        ]

        # Subclause 1.2 -> s00_s04
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "1.2").name,
            "24501-k00_1_Main-Body_s00_s04.docx",
        )
        # Subclause 5.1.2 -> s05_s0504
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "5.1.2").name,
            "24501-k00_2_Main-Body_s05_s0504.docx",
        )
        # Subclause 5.5.1 -> s0505_s0506
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "5.5.1").name,
            "24501-k00_3_Main-Body_s0505_s0506.docx",
        )
        # Message Clause 8.2.6 -> s06_s08
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "8.2.6").name,
            "24501-k00_4_Main-Body_s06_s08.docx",
        )
        # IE Clause 9.11.3.7 -> s09_s10
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "9.11.3.7").name,
            "24501-k00_5_Main-Body_s09_s10.docx",
        )
        # Annex A -> sA_sHistory
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "Annex A").name,
            "24501-k00_6_Annexes_sA_sHistory.docx",
        )
        # Clause D.6 -> sA_sHistory
        self.assertEqual(
            match_multipart_document_for_clause(docs_24501, "D.6").name,
            "24501-k00_6_Annexes_sA_sHistory.docx",
        )

    def test_version_tree_signals_and_context_menu(self):
        """Verifies that NASVersionTreeWidget emits open_spec_requested and populates menu."""
        tree = NASVersionTreeWidget()
        versions = [
            {"id": 1, "spec_number": "24.501", "version": "20.0.0"},
            {"id": 2, "spec_number": "24.501", "version": "19.7.0"},
        ]
        tree.populate(versions)

        spec_item = tree.topLevelItem(1)  # [CT1] TS 24.501
        self.assertIsNotNone(spec_item)
        version_item = spec_item.child(0)  # Rel-20 v20.0.0
        self.assertIsNotNone(version_item)

        # Test double-click emission
        emitted_signals = []
        tree.open_spec_requested.connect(lambda s, v, c, fmt: emitted_signals.append((s, v, c, fmt)))

        tree._on_item_double_clicked(version_item, 0)
        self.assertEqual(len(emitted_signals), 1)
        self.assertEqual(emitted_signals[0], ("24.501", "20.0.0", "", "word"))

        # Test spec-level double click
        tree._on_item_double_clicked(spec_item, 0)
        self.assertEqual(len(emitted_signals), 2)
        self.assertEqual(emitted_signals[1], ("24.501", "20.0.0", "", "word"))

    def test_inspector_open_spec_button_and_signals(self):
        """Verifies NASInspectorWidget open_spec_btn state, signals, and actions."""
        inspector = NASInspectorWidget()
        self.assertFalse(inspector.open_spec_btn.isEnabled())

        emitted_requests = []
        inspector.open_spec_requested.connect(lambda s, v, c, fmt: emitted_requests.append((s, v, c, fmt)))

        defs = [
            {
                "spec_number": "24.501",
                "version": "20.0.0",
                "ie_name": "5GS registration type",
                "clause": "9.11.3.7",
                "raw_description": "<div>5GS registration type definition</div>",
            }
        ]

        inspector.display_definitions(
            clause="9.11.3.7",
            ie_name="5GS registration type",
            spec_number="24.501",
            defs=defs,
            containing_msgs=[],
        )

        self.assertTrue(inspector.open_spec_btn.isEnabled())
        self.assertIn("TS 24.501", inspector.open_spec_btn.text())

        # Test clicking the button
        inspector._on_open_spec_btn_clicked()
        self.assertEqual(len(emitted_requests), 1)
        self.assertEqual(emitted_requests[0], ("24.501", "20.0.0", "9.11.3.7", "word"))

        # Test clearing
        inspector.clear_display()
        self.assertFalse(inspector.open_spec_btn.isEnabled())

    @patch("modules.spec_search.core.spec_doc_resolver.SpecDocResolver.open_spec")
    def test_nas_tab_open_specification_dispatch(self, mock_open_spec):
        """Verifies that NASTab dispatches spec opening with correct arguments."""
        mock_open_spec.return_value = True

        tab = NASTab(self.dummy_nas_db, self.dummy_specs_db)
        tab._current_message_spec = "24.501"
        tab._current_message_clause = "8.2.6"

        tab._update_open_spec_button()
        self.assertTrue(tab.open_spec_btn.isEnabled())
        self.assertIn("24.501", tab.open_spec_btn.text())

        res = tab.open_specification(
            spec_number="24.501",
            version="20.0.0",
            clause_number="8.2.6",
            doc_format="word",
        )
        self.assertTrue(res)
        mock_open_spec.assert_called_once()
        call_kwargs = mock_open_spec.call_args[1]
        self.assertEqual(call_kwargs["spec_number"], "24.501")
        self.assertEqual(call_kwargs["version"], "20.0.0")
        self.assertEqual(call_kwargs["clause_number"], "8.2.6")
        self.assertEqual(call_kwargs["doc_format"], "word")


if __name__ == "__main__":
    unittest.main()
