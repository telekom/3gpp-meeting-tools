"""
Tests for global UI font scaling, DisplaySettingsDialog, and live reader zoom.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

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
    scale_pt,
    scale_px,
    scale_row_height,
    scale_stylesheet,
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


if __name__ == "__main__":
    unittest.main()
