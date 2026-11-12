"""Unit tests for PlantUML Editor drag & drop and file loading capabilities."""

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtCore import QMimeData, QPoint, Qt, QUrl
from PyQt5.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent
from PyQt5.QtWidgets import QApplication

# Ensure QApplication is initialized for widget tests
app = QApplication.instance()
if app is None:
    app = QApplication([])

from modules.puml2visio.ui.ui_components import (
    PUML_EXTENSIONS,
    SUPPORTED_EDITOR_EXTENSIONS,
    VISIO_EXTENSIONS,
    CodeDropTextEdit,
)
from modules.puml2visio.ui.ui_tabs import CodeEditorTab
from modules.puml2visio.ui.svg_preview_widget import EmbeddedSvgPreview


class TestCodeEditorDragDrop(unittest.TestCase):
    """Test suite for drag & drop file loading in PlantUML editor."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_path = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_extension_detection(self):
        """Verify extension checkers recognize all PlantUML and Visio formats."""
        for ext in ('.puml', '.plantuml', '.pu', '.iuml', '.wsd', '.txt'):
            path = f"C:/docs/my_diagram{ext}"
            self.assertTrue(CodeDropTextEdit.is_puml_file(path), f"Failed for {ext}")
            self.assertTrue(CodeDropTextEdit.is_supported_file(path), f"Failed for {ext}")
            self.assertFalse(CodeDropTextEdit.is_visio_file(path), f"Failed for {ext}")

        for ext in ('.vsdx', '.vsdm'):
            path = f"C:/docs/diagram{ext}"
            self.assertTrue(CodeDropTextEdit.is_visio_file(path), f"Failed for {ext}")
            self.assertTrue(CodeDropTextEdit.is_supported_file(path), f"Failed for {ext}")
            self.assertFalse(CodeDropTextEdit.is_puml_file(path), f"Failed for {ext}")

        # Unsupported files
        for ext in ('.docx', '.png', '.pdf', '.exe', ''):
            path = f"C:/docs/file{ext}"
            self.assertFalse(CodeDropTextEdit.is_supported_file(path), f"Should reject {ext}")

    def test_load_puml_file_content(self):
        """Verify CodeDropTextEdit loads a .puml file from disk and supports undo."""
        editor = CodeDropTextEdit()
        editor.setPlainText("@startuml\ninitial state\n@enduml")

        test_file = self.temp_path / "sequence.puml"
        puml_content = "@startuml\nAlice -> Bob : Authentication Request\nBob --> Alice : Response\n@enduml"
        test_file.write_text(puml_content, encoding="utf-8")

        success = editor.load_file(test_file)
        self.assertTrue(success)
        self.assertEqual(editor.toPlainText().strip(), puml_content.strip())
        self.assertIsNone(editor.error_line)

        # Test single undo step restores initial content
        editor.undo()
        self.assertEqual(editor.toPlainText().strip(), "@startuml\ninitial state\n@enduml")

    def test_load_puml_file_with_bom(self):
        """Verify CodeDropTextEdit correctly handles UTF-8 with BOM."""
        editor = CodeDropTextEdit()
        test_file = self.temp_path / "diagram_bom.puml"
        puml_content = "@startuml\nparticipant UE\nparticipant gNB\n@enduml"
        test_file.write_text(puml_content, encoding="utf-8-sig")

        success = editor.load_file(test_file)
        self.assertTrue(success)
        self.assertEqual(editor.toPlainText().strip(), puml_content.strip())
        # Ensure BOM is not included as a leading character
        self.assertFalse(editor.toPlainText().startswith("\ufeff"))

    def test_editor_drag_and_drop_event(self):
        """Verify dragEnterEvent and dropEvent on CodeDropTextEdit emit file_dropped."""
        editor = CodeDropTextEdit()
        test_file = self.temp_path / "sample.puml"
        test_file.write_text("@startuml\n@enduml", encoding="utf-8")

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(test_file))])

        dropped_files = []
        editor.file_dropped.connect(dropped_files.append)

        # Simulate Drag Enter
        enter_event = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        editor.dragEnterEvent(enter_event)
        self.assertTrue(enter_event.isAccepted())
        self.assertIn("dashed", editor.styleSheet())

        # Simulate Drag Move
        move_event = QDragMoveEvent(QPoint(15, 15), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        editor.dragMoveEvent(move_event)
        self.assertTrue(move_event.isAccepted())

        # Simulate Drop
        drop_event = QDropEvent(QPoint(15, 15), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        editor.dropEvent(drop_event)
        self.assertTrue(drop_event.isAccepted())
        self.assertNotIn("dashed", editor.styleSheet())
        self.assertEqual(len(dropped_files), 1)
        self.assertEqual(Path(dropped_files[0]).resolve(), test_file.resolve())

    def test_editor_rejects_unsupported_drag(self):
        """Verify dragEnterEvent ignores files with unsupported extensions."""
        editor = CodeDropTextEdit()
        bad_file = self.temp_path / "report.docx"
        bad_file.write_text("dummy", encoding="utf-8")

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(bad_file))])

        enter_event = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        editor.dragEnterEvent(enter_event)
        self.assertFalse(enter_event.isAccepted())
        self.assertNotIn("dashed", editor.styleSheet())

    def test_code_editor_tab_drag_drop_and_load(self):
        """Verify CodeEditorTab forwards drops and manages current file path."""
        tab = CodeEditorTab()
        test_file = self.temp_path / "network_flow.puml"
        content = "@startuml\nUE -> AMF : Registration Request\n@enduml"
        test_file.write_text(content, encoding="utf-8")

        dropped = []
        tab.file_dropped.connect(dropped.append)

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(test_file))])

        drop_event = QDropEvent(QPoint(50, 50), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        tab.dropEvent(drop_event)
        self.assertTrue(drop_event.isAccepted())
        self.assertEqual(len(dropped), 1)

        # Direct load on tab
        ok = tab.load_file(test_file)
        self.assertTrue(ok)
        self.assertEqual(tab.get_text(), content.strip())
        self.assertEqual(tab.get_current_file_path(), test_file)

    def test_svg_preview_forwards_drag_drop(self):
        """Verify EmbeddedSvgPreview forwards dropped PlantUML files to file_dropped."""
        preview = EmbeddedSvgPreview()
        test_file = self.temp_path / "preview_drop.puml"
        test_file.write_text("@startuml\n@enduml", encoding="utf-8")

        dropped = []
        preview.file_dropped.connect(dropped.append)

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(test_file))])

        drop_event = QDropEvent(QPoint(20, 20), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
        preview.dropEvent(drop_event)
        self.assertTrue(drop_event.isAccepted())
        self.assertEqual(len(dropped), 1)
        self.assertEqual(Path(dropped[0]).resolve(), test_file.resolve())

    def test_main_window_dispatch_puml(self):
        """Verify _on_editor_file_dropped routes .puml files to load_puml_file."""
        from main_window import DragDropUI
        # Create a mock DragDropUI instance without invoking full GUI startup
        win = MagicMock(spec=DragDropUI)
        win.extract_code_from_visio = MagicMock()
        win.load_puml_file = MagicMock()

        # Call unbound method with mock win
        DragDropUI._on_editor_file_dropped(win, "C:/test/diagram.puml")
        win.load_puml_file.assert_called_once_with("C:/test/diagram.puml")
        win.extract_code_from_visio.assert_not_called()

    def test_main_window_dispatch_visio(self):
        """Verify _on_editor_file_dropped routes .vsdx files to extract_code_from_visio."""
        from main_window import DragDropUI
        win = MagicMock(spec=DragDropUI)
        win.extract_code_from_visio = MagicMock()
        win.load_puml_file = MagicMock()

        DragDropUI._on_editor_file_dropped(win, "C:/test/diagram.vsdx")
        win.extract_code_from_visio.assert_called_once_with("C:/test/diagram.vsdx")
        win.load_puml_file.assert_not_called()

    def test_main_window_load_puml_file(self):
        """Verify DragDropUI.load_puml_file invokes tab load_file, updates preview, and logs."""
        from main_window import DragDropUI
        win = MagicMock(spec=DragDropUI)
        win.code_tab = MagicMock()
        win.code_tab.load_file.return_value = True
        win.live_preview = MagicMock()
        win.log_message = MagicMock()
        win.status_bar = MagicMock()

        test_file = self.temp_path / "actual.puml"
        test_file.write_text("@startuml\n@enduml", encoding="utf-8")

        DragDropUI.load_puml_file(win, str(test_file))
        win.code_tab.load_file.assert_called_once_with(test_file)
        win.live_preview.update_now.assert_called_once()
        win.log_message.assert_called_once()
        self.assertIn("Loaded PlantUML diagram", win.log_message.call_args[0][0])

    def test_code_editor_phase2_ribbon_and_icons(self):
        from core.ui.page_ribbon import PageRibbonCard

        tab = CodeEditorTab()
        self.assertIsInstance(tab.ribbon_card, PageRibbonCard)
        for btn in (
            tab.insert_tpl_btn,
            tab.insert_elem_btn,
            tab.palette_btn,
            tab.find_btn,
            tab.docs_btn,
            tab.ai_gen_btn,
            tab.preview_dock_btn,
            tab.clear_btn,
            tab.undo_btn,
            tab.copy_code_btn,
            tab.live_view_btn,
            tab.planttext_btn,
            tab.copy_btn,
            tab.open_folder_btn,
            tab.export_btn,
        ):
            self.assertFalse(btn.icon().isNull())


if __name__ == "__main__":
    unittest.main()
