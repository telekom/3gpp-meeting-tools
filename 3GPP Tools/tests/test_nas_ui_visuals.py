import sys
import unittest
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.nas.ui.nas_dialogs import (
    PROTOCOL_SPEC_CATALOG,
    get_spec_metadata,
    get_spec_title,
    get_spec_tooltip,
)
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.nas.ui.nas_components import (
    NASInspectorWidget,
    NASMessageItemDelegate,
    NASVersionCardsWidget,
    NASVersionTreeWidget,
)

# Ensure QApplication exists for Qt widgets
app = QApplication.instance() or QApplication(sys.argv)


class TestNASUiVisuals(unittest.TestCase):
    """Verifies visual enhancements, specification clue resolution, and model rendering in Protocols UI."""

    def test_protocol_spec_catalog_and_metadata(self):
        """Verifies that key 3GPP specifications resolve to rich descriptive metadata."""
        specs_to_test = ["24.501", "24.301", "29.244", "29.281", "38.331", "38.413", "38.415"]

        for s_num in specs_to_test:
            meta = get_spec_metadata(s_num)
            self.assertIn("acronym", meta)
            self.assertIn("wg", meta)
            self.assertIn("title", meta)
            self.assertIn("full_title", meta)
            self.assertIn("description", meta)
            self.assertTrue(len(meta["description"]) > 10)

            # Verify tooltip contains Working Group and description
            tooltip = get_spec_tooltip(s_num)
            self.assertIn(meta["wg"], tooltip)
            self.assertIn(meta["description"], tooltip)

    def test_heuristic_spec_metadata_fallback(self):
        """Verifies that uncatalogued specs gracefully infer series and WG."""
        meta = get_spec_metadata("24.999")
        self.assertEqual(meta["wg"], "CT1")
        self.assertEqual(meta["tech"], "Core")

        meta_ran = get_spec_metadata("38.399")
        self.assertEqual(meta_ran["wg"], "RAN2")

    def test_version_tree_population_and_clues(self):
        """Verifies that version tree items display Working Group badges and Release tags."""
        tree = NASVersionTreeWidget()
        versions = [
            {"id": 1, "spec_number": "24.501", "version": "20.0.0"},
            {"id": 2, "spec_number": "24.501", "version": "19.7.0"},
            {"id": 3, "spec_number": "29.244", "version": "20.0.0"},
        ]
        tree.populate(versions)

        # Top level has master toggle + 2 specs
        self.assertEqual(tree.topLevelItemCount(), 3)

        # Spec 24.501 should show [CT1] and 5GS NAS
        spec_item_24501 = tree.topLevelItem(1)
        self.assertIn("CT1", spec_item_24501.text(0))
        self.assertIn("24.501", spec_item_24501.text(0))

        # Child version should display Rel-20
        child_v20 = spec_item_24501.child(0)
        self.assertIn("Rel-20", child_v20.text(0))
        self.assertIn("v20.0.0", child_v20.text(0))

        # Test quick selection helpers
        tree.select_none()
        self.assertEqual(len(tree.get_selected_version_ids()), 0)

        tree.select_all()
        self.assertEqual(len(tree.get_selected_version_ids()), 3)

        tree.select_latest_only()
        selected_ids = tree.get_selected_version_ids()
        # Should select latest v20.0.0 for 24.501 (id 1) and latest v20.0.0 for 29.244 (id 3)
        self.assertEqual(len(selected_ids), 2)
        self.assertIn(1, selected_ids)
        self.assertIn(3, selected_ids)

    def test_evolution_matrix_headers_with_release(self):
        """Verifies that table headers format version columns with clear Release badges."""
        df = pd.DataFrame([
            {
                "iei": "5E",
                "ie_name": "PDU session identity 2",
                "type_reference": "9.11.4.10",
                "presence": "M",
                "format": "V",
                "length": "1",
                "spec_number": "24.501",
                "version": "19.7.0",
                "order_index": 1,
            },
            {
                "iei": "5E",
                "ie_name": "PDU session identity 2",
                "type_reference": "9.11.4.10",
                "presence": "M",
                "format": "V",
                "length": "1",
                "spec_number": "24.501",
                "version": "20.0.0",
                "order_index": 1,
            },
        ])
        model = NASEvolutionMatrixModel(df)
        self.assertGreaterEqual(model.columnCount(), 4)

        # Columns 0, 1, 2 are IEI, Field, Type
        self.assertEqual(model.headerData(0, Qt.Horizontal, Qt.DisplayRole), "IEI / ID")
        self.assertEqual(model.headerData(1, Qt.Horizontal, Qt.DisplayRole), "Field / Information Element")
        self.assertEqual(model.headerData(2, Qt.Horizontal, Qt.DisplayRole), "Type / Reference")

        # Column 3+ should contain Release tag
        v_header = model.headerData(3, Qt.Horizontal, Qt.DisplayRole)
        self.assertTrue("Rel-19" in v_header or "Rel-20" in v_header)

    def test_inspector_widget_stack_and_empty_state(self):
        """Verifies that the inspector card manages its empty state and populated state."""
        inspector = NASInspectorWidget()
        self.assertEqual(inspector.stack.currentIndex(), 0)  # Starts on empty state
        self.assertFalse(inspector.copy_btn.isEnabled())

        inspector.display_definitions(
            clause="8.2.1",
            ie_name="Test IE",
            spec_number="24.501",
            defs=[{
                "ie_name": "Test IE",
                "spec_number": "24.501",
                "version": "20.0.0",
                "raw_description": "<p>Sample description</p>",
            }],
            containing_msgs=[],
        )
        self.assertEqual(inspector.stack.currentIndex(), 1)  # Switches to content view
        self.assertTrue(inspector.copy_btn.isEnabled())

        inspector.clear_display()
        self.assertEqual(inspector.stack.currentIndex(), 0)  # Reverts to empty state
        self.assertFalse(inspector.copy_btn.isEnabled())

    def test_asn1_prose_summary_and_direction_extraction(self):
        """Verifies extraction of message description and Direction from Clause 9.2 / 6.2 HTML."""
        from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser

        sample_html = (
            '<p style="margin: 3px 0;">This message is sent by the AMF to transfer updated information '
            'for an NG-C interface instance.</p>'
            '<p style="margin: 3px 0; color: #0369A1; font-weight: bold;">Direction: AMF → NG-RAN node</p>'
            '<table><tr><th>IE/Group Name</th><th>Presence</th></tr></table>'
        )
        desc, direction = BaseAsn1DocxParser._extract_prose_summary_and_direction(sample_html)
        self.assertEqual(
            desc,
            "This message is sent by the AMF to transfer updated information for an NG-C interface instance.",
        )
        self.assertEqual(direction, "AMF → NG-RAN node")

    def test_nas_db_message_context_and_ui_banner(self):
        """Verifies DB storage/fallback of message description & direction and NASTab banner/inspector population."""
        import tempfile
        from PyQt5.QtWidgets import QListWidgetItem
        from modules.nas.core.nas_db import NASDatabase
        from modules.nas.ui.nas_tabs import NASTab

        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "test_nas.db"
            db = NASDatabase(db_path=db_path)

            # 1. Insert a parsed NGAP message with explicit description & direction
            messages = [
                {
                    "clause": "9.2.6.7",
                    "message_name": "AMFConfigurationUpdate",
                    "table_caption": "AMFConfigurationUpdate ASN.1 PDU Definition (Clause 9.2.6.7)",
                    "description": "This message is sent by the AMF to transfer updated information for an NG-C interface instance.",
                    "direction": "AMF → NG-RAN node",
                    "ies": [
                        {
                            "iei": "0",
                            "information_element": "MessageType",
                            "field_path": "MessageType",
                            "depth": 0,
                            "type_reference": "MessageType (9.3.1.1)",
                            "presence": "M",
                            "format": "IE",
                            "length": "reject",
                        }
                    ],
                },
                {
                    # 2. Legacy-style message without explicit description column, testing fallback from ie_definitions
                    "clause": "9.2.6.8",
                    "message_name": "AMFConfigurationUpdateAcknowledge",
                    "table_caption": "AMFConfigurationUpdateAcknowledge ASN.1 PDU Definition (Clause 9.2.6.8)",
                    "ies": [
                        {
                            "iei": "0",
                            "information_element": "MessageType",
                            "field_path": "MessageType",
                            "depth": 0,
                            "type_reference": "MessageType (9.3.1.1)",
                            "presence": "M",
                            "format": "IE",
                            "length": "ignore",
                        }
                    ],
                },
            ]
            ie_defs = [
                {
                    "clause": "9.2.6.7",
                    "ie_name": "AMFConfigurationUpdate",
                    "raw_description": (
                        "<h2>AMFConfigurationUpdate (Clause 9.2.6.7)</h2>"
                        "<p>This message is sent by the AMF to transfer updated information for an NG-C interface instance.</p>"
                        "<p>Direction: AMF → NG-RAN node</p>"
                    ),
                    "structure_table": "[]",
                },
                {
                    "clause": "9.2.6.8",
                    "ie_name": "AMFConfigurationUpdateAcknowledge",
                    "raw_description": (
                        "<h2>AMFConfigurationUpdateAcknowledge (Clause 9.2.6.8)</h2>"
                        "<div>Specification Details &amp; Tabular Definition:</div>"
                        "<p>This message is sent by the NG-RAN node to acknowledge the AMF CONFIGURATION UPDATE.</p>"
                        "<p>Direction: NG-RAN node → AMF</p>"
                        "<table><tr><td>MessageType</td></tr></table>"
                    ),
                    "structure_table": "[]",
                },
            ]

            ver_id = db.insert_parsed_spec("38.413", "19.3.0", messages, ie_defs)

            # Verify get_message_context_info for explicit columns
            ctx1 = db.get_message_context_info("AMFConfigurationUpdate", version_ids=[ver_id])
            self.assertIn("transfer updated information", ctx1["description"])
            self.assertEqual(ctx1["direction"], "AMF → NG-RAN node")

            # Verify get_message_context_info fallback from ie_definitions
            ctx2 = db.get_message_context_info("AMFConfigurationUpdateAcknowledge", version_ids=[ver_id])
            self.assertIn("acknowledge the AMF CONFIGURATION UPDATE", ctx2["description"])
            self.assertEqual(ctx2["direction"], "NG-RAN node → AMF")

            # Verify deep search matches message-level description
            matched = db.get_messages_by_ie_search(
                "NG-C interface instance",
                version_ids=[ver_id],
                search_descriptions=True,
            )
            self.assertEqual(len(matched), 1)
            self.assertEqual(matched[0]["message_name"], "AMFConfigurationUpdate")

            # Verify NASTab UI updates banner & inspector when message is clicked
            from core.ui.page_ribbon import PageRibbonCard

            tab = NASTab(nas_db_path=db_path)
            self.assertIsInstance(tab.ribbon_card, PageRibbonCard)
            tab.resize(1280, 800)
            tab.show()
            app.processEvents()
            self.assertLessEqual(tab.ribbon_card.height(), 80)
            self.assertLessEqual(tab.stats_specs_lbl.height(), 36)
            self.assertFalse(tab.fetch_btn.icon().isNull())
            self.assertFalse(tab.import_file_btn.icon().isNull())
            self.assertFalse(tab.manage_btn.icon().isNull())
            self.assertFalse(tab.open_spec_btn.icon().isNull())
            self.assertFalse(tab.copy_matrix_btn.icon().isNull())
            tab.selected_version_ids = [ver_id]

            item = QListWidgetItem("AMFConfigurationUpdate")
            item.setData(Qt.UserRole, "AMFConfigurationUpdate")
            item.setData(Qt.UserRole + 1, matched[0])
            item.setData(Qt.UserRole + 2, "38.413")
            item.setData(Qt.UserRole + 3, "9.2.6.7")

            tab._on_message_clicked(item)
            self.assertFalse(tab.msg_context_banner.isHidden())
            self.assertIn("AMF → NG-RAN node", tab.msg_direction_badge.text())
            self.assertIn("NG-C interface instance", tab.msg_description_lbl.text())
            self.assertEqual(tab.inspector.stack.currentIndex(), 1)
            self.assertIn("AMFConfigurationUpdate", tab.inspector.title_lbl.text())

    def test_version_cards_widget_pills_filtering_and_quick_selectors(self):
        """Verifies NASVersionCardsWidget cards, release pills, filtering, and quick-selection chips."""
        cards = NASVersionCardsWidget()
        versions = [
            {"id": 1, "spec_number": "24.501", "version": "20.0.0"},
            {"id": 2, "spec_number": "24.501", "version": "19.7.0"},
            {"id": 3, "spec_number": "24.501", "version": "19.6.0"},
            {"id": 4, "spec_number": "38.413", "version": "17.9.0"},
        ]
        cards.populate(versions)
        self.assertEqual(len(cards._cards), 2)
        self.assertEqual(cards.topLevelItemCount(), 3)

        # Latest/Rel selects newest version per major release (20.0.0 + 19.7.0 + 17.9.0)
        cards.select_latest_per_release()
        self.assertEqual(set(cards.get_selected_version_ids()), {1, 2, 4})

        # Latest Only selects newest single version per spec (20.0.0 + 17.9.0)
        cards.select_latest_only()
        self.assertEqual(set(cards.get_selected_version_ids()), {1, 4})

        # Rel-18+ selects only Rel-18/19/20 (ids 1, 2, 3; excludes Rel-17 id 4)
        cards.select_releases_min(18)
        self.assertEqual(set(cards.get_selected_version_ids()), {1, 2, 3})

        # Select All / None
        cards.select_none()
        self.assertEqual(cards.get_selected_version_ids(), [])
        cards.select_all()
        self.assertEqual(len(cards.get_selected_version_ids()), 4)

        # Quick text filter by acronym / spec number (show() needed so isHidden/isVisible reflects state)
        cards.show()
        cards.filter_tree("NGAP")
        self.assertTrue(cards._cards["24.501"].isHidden())
        self.assertFalse(cards._cards["38.413"].isHidden())
        cards.filter_tree("")
        self.assertFalse(cards._cards["24.501"].isHidden())
        cards.close()

    def test_two_line_message_delegate_and_diff_filter_pills(self):
        """Verifies 2-line NASMessageItemDelegate sizing and interactive diff filter pills."""
        import tempfile
        from PyQt5.QtWidgets import QStyleOptionViewItem
        from modules.nas.core.nas_db import NASDatabase
        from modules.nas.ui.nas_tabs import NASTab

        delegate = NASMessageItemDelegate()
        hint = delegate.sizeHint(QStyleOptionViewItem(), None)
        self.assertGreaterEqual(hint.height(), 42)

        df = pd.DataFrame([
            {
                "iei": "1", "ie_name": "SharedField", "type_reference": "9.1",
                "presence": "M", "format": "V", "length": "1",
                "spec_number": "24.501", "version": "19.7.0", "order_index": 1,
            },
            {
                "iei": "1", "ie_name": "SharedField", "type_reference": "9.1",
                "presence": "M", "format": "V", "length": "1",
                "spec_number": "24.501", "version": "20.0.0", "order_index": 1,
            },
            {
                "iei": "2", "ie_name": "AddedField", "type_reference": "9.2",
                "presence": "O", "format": "TV", "length": "2",
                "spec_number": "24.501", "version": "20.0.0", "order_index": 2,
            },
            {
                "iei": "3", "ie_name": "ModifiedField", "type_reference": "9.3",
                "presence": "O", "format": "TLV", "length": "2-4",
                "spec_number": "24.501", "version": "19.7.0", "order_index": 3,
            },
            {
                "iei": "3", "ie_name": "ModifiedField", "type_reference": "9.3",
                "presence": "O", "format": "TLV", "length": "2-8",
                "spec_number": "24.501", "version": "20.0.0", "order_index": 3,
            },
        ])
        full_model = NASEvolutionMatrixModel(df)
        self.assertEqual(full_model.rowCount(), 3)
        self.assertEqual(full_model.diff_counts, {"added": 1, "modified": 1, "removed": 0})

        with tempfile.TemporaryDirectory() as tmpdir:
            tab = NASTab(nas_db_path=Path(tmpdir) / "nas_diff_ui.db")
            self.assertIsInstance(tab.version_tree, NASVersionCardsWidget)
            self.assertTrue(tab.ribbon_card.isHidden())
            tab._current_matrix_df = df
            tab._rebuild_matrix_model()
            self.assertEqual(tab.matrix_table.model().rowCount(), 3)
            self.assertIn("(1)", tab.btn_diff_added.text())
            self.assertIn("(1)", tab.btn_diff_modified.text())

            tab.btn_diff_added.setChecked(True)
            self.assertEqual(tab.matrix_table.model().rowCount(), 1)
            tab.btn_diff_modified.setChecked(True)
            self.assertEqual(tab.matrix_table.model().rowCount(), 2)
            tab.btn_diff_added.setChecked(False)
            tab.btn_diff_modified.setChecked(False)
            self.assertEqual(tab.matrix_table.model().rowCount(), 3)
            tab.close()


if __name__ == "__main__":
    unittest.main()
