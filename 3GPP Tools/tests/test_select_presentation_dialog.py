"""Unit tests for SelectPresentationDialog in PowerPoint PlantUML Export."""

import sys
import unittest
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from modules.puml2visio.ui.select_presentation_dialog import (
    SelectPresentationDialog,
    PresentationItemWidget,
    _format_display_path,
)

# Ensure QApplication singleton exists for tests
app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestSelectPresentationDialog(unittest.TestCase):
    """Test suite verifying ergonomics, selection state, and path formatting of SelectPresentationDialog."""

    def setUp(self):
        self.sample_presentations = [
            {
                "name": "2026E00XXX Diagrams.pptx",
                "full_name": "C:/Projects/3GPP/TSG-SA/WG2/7 Service-aware RAN CN interface/2026E00XXX Diagrams.pptx",
                "slides_count": 12,
                "is_active": True,
                "saved": True,
            },
            {
                "name": "Presentation1",
                "full_name": "Presentation1",
                "slides_count": 1,
                "is_active": False,
                "saved": False,
            },
            {
                "name": "RAN3_Architecture.pptx",
                "full_name": "D:/Standards/RAN3_Architecture.pptx",
                "slides_count": 45,
                "is_active": False,
                "saved": True,
            },
        ]

    def test_path_formatting(self):
        """Verify file path truncation prevents card overflow while preserving context."""
        # Unsaved document
        self.assertEqual(
            _format_display_path("Presentation1", "Presentation1"),
            "New presentation (not saved to disk yet)",
        )
        # Short path
        short_path = "C:/docs/test.pptx"
        self.assertEqual(_format_display_path(short_path, "test.pptx"), short_path)

        # Long path truncation
        long_path = "C:/Users/colomikunoj/Documents/Long_Directory_Name/Deep_Nested/2026E00XXX Diagrams.pptx"
        truncated = _format_display_path(long_path, "2026E00XXX Diagrams.pptx", max_chars=40)
        self.assertTrue("..." in truncated)
        self.assertLessEqual(len(truncated), 40)
        self.assertTrue(truncated.startswith("C:/Users/colomi"))

    def test_item_widget_size_hints(self):
        """Ensure item widget provides ample height to prevent font clipping."""
        widget = PresentationItemWidget(self.sample_presentations[0])
        hint = widget.sizeHint()
        min_hint = widget.minimumSizeHint()
        self.assertGreaterEqual(hint.height(), 68)
        self.assertGreaterEqual(min_hint.height(), 64)

    def test_dialog_initialization_and_preselection(self):
        """Verify the active deck is automatically preselected upon initialization."""
        dialog = SelectPresentationDialog(self.sample_presentations)
        # Active deck is index 0
        self.assertEqual(dialog.list_widget.currentRow(), 0)
        selected = dialog.get_selected_presentation()
        self.assertIsNotNone(selected)
        self.assertEqual(selected["name"], "2026E00XXX Diagrams.pptx")
        self.assertEqual(selected["insert_at"], "end")
        self.assertIn("Slide 13", dialog.radio_end.text())

    def test_selection_change_updates_labels(self):
        """Verify selecting another presentation updates the slide placement text and visual indicator."""
        dialog = SelectPresentationDialog(self.sample_presentations)
        # Select second item (Presentation1, 1 slide)
        dialog.list_widget.setCurrentRow(1)
        selected = dialog.get_selected_presentation()
        self.assertEqual(selected["name"], "Presentation1")
        self.assertIn("Slide 2", dialog.radio_end.text())

        # Check visual indicator on widgets
        item0_widget = dialog.list_widget.itemWidget(dialog.list_widget.item(0))
        item1_widget = dialog.list_widget.itemWidget(dialog.list_widget.item(1))
        self.assertFalse(item0_widget._is_selected)
        self.assertTrue(item1_widget._is_selected)
        self.assertEqual(item1_widget.status_pill.text(), "✓ Selected")

    def test_slide_placement_preference(self):
        """Verify user can toggle between appending at end and inserting after active slide."""
        dialog = SelectPresentationDialog(self.sample_presentations)
        # Default is end
        selected = dialog.get_selected_presentation()
        self.assertEqual(selected["insert_at"], "end")

        # Toggle to after active slide
        dialog.radio_after_active.setChecked(True)
        selected_after = dialog.get_selected_presentation()
        self.assertEqual(selected_after["insert_at"], "after_active")

    def test_new_presentation_action(self):
        """Verify user can divert to creating a new presentation."""
        dialog = SelectPresentationDialog(self.sample_presentations)
        self.assertFalse(dialog.is_new_presentation_requested())
        dialog._on_new_presentation_clicked()
        self.assertTrue(dialog.is_new_presentation_requested())

    def test_filtering(self):
        """Verify search filter hides non-matching presentations."""
        dialog = SelectPresentationDialog(self.sample_presentations)
        self.assertFalse(dialog.filter_edit.isHidden())

        # Filter by 'RAN3'
        dialog._on_filter_changed("RAN3")
        self.assertTrue(dialog.list_widget.item(0).isHidden())
        self.assertTrue(dialog.list_widget.item(1).isHidden())
        self.assertFalse(dialog.list_widget.item(2).isHidden())

        # Selected should switch to visible item
        self.assertEqual(dialog.list_widget.currentRow(), 2)

        # Clear filter
        dialog._on_filter_changed("")
        self.assertFalse(dialog.list_widget.item(0).isHidden())
        self.assertFalse(dialog.list_widget.item(1).isHidden())
        self.assertFalse(dialog.list_widget.item(2).isHidden())


if __name__ == "__main__":
    unittest.main()
