"""
Tests for global UI font scaling, DisplaySettingsDialog, and live reader zoom.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QTableWidget, QTextBrowser

from core.ui.command_palette import build_static_command_items
from core.ui.display_settings_dialog import DisplaySettingsDialog
from core.ui.font_scale import (
    DEFAULT_FONT_SCALE,
    FONT_SCALE_PRESETS,
    apply_font_scale,
    clamp_font_scale,
    get_font_scale,
    install_font_scaling,
    load_font_scale,
    preset_for_scale,
    register_scaled_row_height,
    save_font_scale,
    scale_html_fonts,
    scale_pt,
    scale_px,
    scale_row_height,
    scale_stylesheet,
    scaled_font,
    zoom_text_widget,
)


class _DummyMainWindow:
    def __init__(self) -> None:
        self.display_opened = 0

    def open_display_settings(self) -> None:
        self.display_opened += 1


class TestFontScale(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)
        install_font_scaling(cls.app)

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.settings = QSettings(str(Path(self._tmp.name) / "shell.ini"), QSettings.IniFormat)
        apply_font_scale(DEFAULT_FONT_SCALE, app=self.app, save=False)

    def tearDown(self) -> None:
        apply_font_scale(DEFAULT_FONT_SCALE, app=self.app, save=False)
        self._tmp.cleanup()

    def test_clamp_and_presets(self) -> None:
        self.assertEqual(clamp_font_scale("invalid"), 1.0)
        self.assertEqual(clamp_font_scale(0.2), 0.85)
        self.assertEqual(clamp_font_scale(3.0), 1.60)
        self.assertEqual(clamp_font_scale("1.25"), 1.25)

        self.assertIsNotNone(preset_for_scale(1.25))
        self.assertEqual(preset_for_scale(1.25).key, "large")  # type: ignore[union-attr]
        self.assertIsNone(preset_for_scale(1.33))
        self.assertGreaterEqual(len(FONT_SCALE_PRESETS), 5)

    def test_scale_helpers_and_stylesheet_rewriting(self) -> None:
        self.assertEqual(scale_px(13, scale=1.25), 16)
        self.assertEqual(scale_pt(8, scale=1.25), 10)
        self.assertEqual(scale_row_height(44, scale=1.25), 55)
        self.assertEqual(scale_row_height(44, scale=0.90), 40)

        qss = (
            "QFrame { max-height: 1px; }"
            "QPushButton { font-size: 12px; min-height: 20px; max-height: 22px; }"
            "QLabel { font-size: 0px; }"
        )
        scaled = scale_stylesheet(qss, 1.25)
        self.assertIn("max-height: 1px;", scaled)  # Dividers untouched
        self.assertIn("font-size: 15px;", scaled)
        self.assertIn("min-height: 25px;", scaled)
        self.assertIn("max-height: 28px;", scaled)
        self.assertIn("font-size: 0px;", scaled)

    def test_settings_round_trip(self) -> None:
        self.assertEqual(load_font_scale(self.settings), 1.0)
        save_font_scale(1.40, self.settings)
        self.assertEqual(load_font_scale(self.settings), 1.40)

    def test_no_double_scaling_and_live_refresh(self) -> None:
        apply_font_scale(1.25, app=self.app, save=False)
        self.assertEqual(get_font_scale(), 1.25)
        self.assertEqual(self.app.font().pixelSize(), 16)

        lbl1 = QLabel("One")
        lbl1.setStyleSheet("QLabel { font-size: 12px; }")
        self.assertIn("font-size: 15px;", lbl1.styleSheet())

        # Copying scaled styleSheet onto another widget must NOT double-scale
        lbl2 = QPushButton("Two")
        lbl2.setStyleSheet(lbl1.styleSheet())
        self.assertIn("font-size: 15px;", lbl2.styleSheet())
        self.assertNotIn("font-size: 19px;", lbl2.styleSheet())

        tbl = QTableWidget(1, 1)
        register_scaled_row_height(tbl, 44)
        self.assertEqual(tbl.verticalHeader().defaultSectionSize(), 55)

        # Live switch to 1.40 and back to 1.00
        apply_font_scale(1.40, app=self.app, save=False)
        self.assertIn("font-size: 17px;", lbl1.styleSheet())
        self.assertIn("font-size: 17px;", lbl2.styleSheet())
        self.assertEqual(tbl.verticalHeader().defaultSectionSize(), 62)

        apply_font_scale(1.00, app=self.app, save=False)
        self.assertIn("font-size: 12px;", lbl1.styleSheet())
        self.assertEqual(tbl.verticalHeader().defaultSectionSize(), 44)

    def test_display_settings_dialog_and_command_palette(self) -> None:
        dlg = DisplaySettingsDialog(settings=self.settings)
        self.assertEqual(dlg.scale_slider.value(), 100)

        dlg._preset_radios["large"].setChecked(True)
        self.assertEqual(dlg.scale_slider.value(), 125)
        self.assertEqual(dlg.pct_badge.text(), "125%")
        self.assertIn("Body 16px", dlg.pv_metrics_label.text())
        self.assertIn("font-size: 16px;", dlg.pv_title.styleSheet())
        self.assertIn("font-size: 15px;", dlg.pv_body.styleSheet())

        dlg._on_apply_clicked()
        self.assertEqual(get_font_scale(), 1.25)
        self.assertEqual(load_font_scale(self.settings), 1.25)
        # Active scale is now 1.25 — preview must NOT double-scale (16 * 1.25 = 20)!
        self.assertIn("font-size: 16px;", dlg.pv_title.styleSheet())
        dlg._preset_radios["xlarge"].setChecked(True)
        self.assertIn("font-size: 18px;", dlg.pv_title.styleSheet())
        self.assertIn("font-size: 17px;", dlg.pv_body.styleSheet())

        dlg._on_reset_clicked()
        self.assertIn("font-size: 13px;", dlg.pv_title.styleSheet())
        dlg._on_ok_clicked()
        self.assertEqual(get_font_scale(), 1.0)
        self.assertEqual(load_font_scale(self.settings), 1.0)

        dummy = _DummyMainWindow()
        items = build_static_command_items(dummy)
        disp_items = [it for it in items if it.id == "sys:display_settings"]
        self.assertEqual(len(disp_items), 1)
        disp_items[0].callback()  # type: ignore[misc]
        self.assertEqual(dummy.display_opened, 1)

    def test_checkable_combobox_size_hint_scales(self) -> None:
        from core.ui.ui_components import SPEC_FILTER_COMBO_STYLE
        from modules.meetings.ui.tdocs_components import CheckableComboBox

        apply_font_scale(1.00, app=self.app, save=False)
        combo = CheckableComboBox("Working Group")
        combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        combo.addItems(["SA1", "SA2", "SA5"])
        hint_100 = combo.sizeHint()
        text_w_100 = combo.fontMetrics().horizontalAdvance("Working Group: All")
        self.assertGreaterEqual(hint_100.width(), text_w_100 + 32)

        apply_font_scale(1.40, app=self.app, save=False)
        hint_140 = combo.sizeHint()
        text_w_140 = combo.fontMetrics().horizontalAdvance("Working Group: All")
        self.assertGreater(hint_140.width(), hint_100.width())
        self.assertGreaterEqual(hint_140.width(), text_w_140 + 40)
        self.assertGreater(hint_140.height(), hint_100.height())

    def test_reader_text_zoom_controls(self) -> None:
        browser = QTextBrowser()
        self.assertEqual(zoom_text_widget(browser, 1), 1)
        self.assertEqual(zoom_text_widget(browser, 2), 3)
        self.assertEqual(zoom_text_widget(browser, -1), 2)
        self.assertEqual(zoom_text_widget(browser, 0, reset=True), 0)
        self.assertIsNone(zoom_text_widget(QLabel("not zoomable"), 1))

    def test_scaled_font_uses_pixel_size_consistently(self) -> None:
        from PyQt5.QtGui import QFont

        f_100 = scaled_font(base_px=12, bold=True, scale=1.00)
        self.assertEqual(f_100.pixelSize(), 12)
        self.assertEqual(f_100.pointSize(), -1)
        self.assertTrue(f_100.bold())

        f_140 = scaled_font(QFont(), base_px=12, family="Consolas", scale=1.40)
        self.assertEqual(f_140.pixelSize(), 17)
        self.assertEqual(f_140.pointSize(), -1)
        self.assertEqual(f_140.family(), "Consolas")

    def test_table_delegates_and_models_stay_pixel_proportional(self) -> None:
        import ast
        from core.ui.theme import T
        from modules.meetings.ui.meetings_delegates import (
            DynamicHoverMenuDelegate,
            TDocsButtonDelegate,
            WGTagDelegate,
        )
        from modules.meetings.ui.tdoc_delegates import AgendaItemDelegate, TDocActionDelegate
        from modules.work_items.ui.work_items_model import WorkItemsTableModel
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QImage, QPainter
        from PyQt5.QtWidgets import QStyleOptionViewItem

        # Source invariant: UI delegate/model painters must not reintroduce DPI-sensitive setPointSize / QFont(..., N)
        checked_paths = (
            SRC_DIR / "modules" / "meetings" / "ui" / "meetings_delegates.py",
            SRC_DIR / "modules" / "meetings" / "ui" / "tdoc_delegates.py",
            SRC_DIR / "modules" / "work_items" / "ui" / "work_items_delegates.py",
            SRC_DIR / "modules" / "work_items" / "ui" / "work_items_model.py",
            SRC_DIR / "modules" / "nas" / "ui" / "nas_components.py",
            SRC_DIR / "modules" / "spec_search" / "ui" / "spec_search_components.py",
        )
        for src_path in checked_paths:
            src_text = src_path.read_text(encoding="utf-8")
            self.assertNotIn("setPointSize", src_text, f"{src_path.name} must use setPixelSize/scaled_font, not setPointSize")
            tree = ast.parse(src_text)
            for node in ast.walk(tree):
                if (
                    isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Name)
                    and node.func.id == "QFont"
                    and len(node.args) >= 2
                ):
                    self.fail(f"{src_path.name}:{node.lineno} calls multi-arg QFont(..., pointSize); use scaled_font()")

        # Live runtime check: Meetings delegates & WorkItems Code font match T.SIZE_TABLE / T.SIZE_CAPTION pixel scale
        wi_model = WorkItemsTableModel([{"code": "990012", "acronym": "6G", "name": "Test"}])
        apply_font_scale(0.90, app=self.app, save=False)
        wi_font_90 = wi_model.data(wi_model.index(0, 1), Qt.FontRole)
        self.assertEqual(wi_font_90.pointSize(), -1)
        self.assertEqual(wi_font_90.pixelSize(), scale_px(T.SIZE_TABLE, scale=0.90))

        apply_font_scale(1.40, app=self.app, save=False)
        wi_font_140 = wi_model.data(wi_model.index(0, 1), Qt.FontRole)
        self.assertEqual(wi_font_140.pointSize(), -1)
        self.assertEqual(wi_font_140.pixelSize(), scale_px(T.SIZE_TABLE, scale=1.40))

        # Smoke-paint Meetings & TDocs delegates under active scaling with pixel-sized app font
        class _DummyParent:
            def __init__(self, tbl):
                self.table = tbl

        tbl = QTableWidget(1, 3)
        dummy_parent = _DummyParent(tbl)
        img = QImage(160, 48, QImage.Format_ARGB32)
        painter = QPainter(img)
        opt = QStyleOptionViewItem()
        opt.rect = img.rect()
        opt.widget = tbl
        opt.font = self.app.font()

        for delegate in (
            WGTagDelegate(dummy_parent),
            TDocsButtonDelegate(dummy_parent),
            DynamicHoverMenuDelegate(dummy_parent),
            TDocActionDelegate(tbl),
            AgendaItemDelegate(tbl),
        ):
            delegate.paint(painter, opt, tbl.model().index(0, 0))
            self.assertEqual(painter.font().pointSize(), -1)
        painter.end()

    def test_inline_html_font_scaling_and_live_refresh(self) -> None:
        from PyQt5.QtGui import QTextDocument
        from PyQt5.QtWidgets import QStyleOptionViewItem, QTableWidgetItem, QTextEdit

        apply_font_scale(1.25, app=self.app, save=False)

        # 1. QLabel constructor & setText with inline HTML font-size
        lbl = QLabel("<span style='color:#334155; font-size: 12px;'>Rich Label</span>")
        self.assertIn("font-size: 15px;", lbl.text())
        # Passing lbl.text() to another label must not double-scale
        lbl_copy = QLabel(lbl.text())
        self.assertIn("font-size: 15px;", lbl_copy.text())

        # 2. QTextBrowser & QTextEdit setHtml
        browser = QTextBrowser()
        browser.setHtml("<div style='font-size: 12px;'>Clause Text</div>")
        self.assertEqual(
            scale_html_fonts(browser.property("_font_scale_base_html"), 1.25),
            "<div style='font-size: 15px;'>Clause Text</div>",
        )

        editor = QTextEdit()
        editor.setHtml("<p style='font-size: 11px;'>Notes</p>")
        self.assertIsNotNone(editor.property("_font_scale_base_html"))
        editor.setPlainText("Plain")
        self.assertIsNone(editor.property("_font_scale_base_html"))

        # 3. QTextDocument setHtml (used by custom delegates)
        doc = QTextDocument()
        doc.setHtml("<div style='font-size: 12px;'>Delegate Cell</div>")
        self.assertIn("font-size:15px", doc.toHtml().replace(" ", ""))

        # 4. Live scale switch to 1.40 and back to 0.90 updates QLabel, QTextBrowser, and bold QTableWidgetItems
        tbl = QTableWidget(1, 2)
        tbl.setStyleSheet("QTableWidget { font-size: 11px; }")
        it_plain = QTableWidgetItem("Plain")
        it_bold = QTableWidgetItem("Bold")
        f_bold = it_bold.font()
        f_bold.setBold(True)
        it_bold.setFont(f_bold)
        tbl.setItem(0, 0, it_plain)
        tbl.setItem(0, 1, it_bold)

        apply_font_scale(1.40, app=self.app, save=False)
        self.assertIn("font-size: 17px;", lbl.text())
        self.assertIn("font-size:17px", browser.toHtml().replace(" ", ""))

        apply_font_scale(0.90, app=self.app, save=False)
        self.assertIn("font-size: 11px;", lbl.text())
        self.assertIn("font-size:11px", browser.toHtml().replace(" ", ""))

        delg = tbl.itemDelegate()
        opt0 = QStyleOptionViewItem()
        opt0.font = tbl.font()
        delg.initStyleOption(opt0, tbl.model().index(0, 0))
        opt1 = QStyleOptionViewItem()
        opt1.font = tbl.font()
        delg.initStyleOption(opt1, tbl.model().index(0, 1))
        self.assertEqual(opt0.font.pixelSize(), opt1.font.pixelSize())
        self.assertTrue(opt1.font.bold())


if __name__ == "__main__":
    unittest.main()
