"""
Regression tests for TDocsWindow maximized responsive layout, multi-line row visibility,
and CheckableComboBox checklist popup indicator rendering under Fusion style.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import Qt, QRect, QSettings
from PyQt5.QtGui import QColor, QImage, QPainter
from PyQt5.QtWidgets import QApplication, QHBoxLayout, QStyle, QStyleOptionViewItem, QWidget

from core.ui.font_scale import apply_font_scale, install_font_scaling, scale_px
from core.ui.theme import C
from core.ui.ui_components import GLOBAL_STYLE, SPEC_FILTER_COMBO_STYLE
from modules.meetings.ui.tdocs_components import CheckableComboBox, CheckableComboItemDelegate
from modules.meetings.ui.tdocs_window import TDocsWindow


class TestTDocsWindowLayoutAndChecklist(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)
        cls.app.setStyle("Fusion")
        cls.app.setStyleSheet(GLOBAL_STYLE)
        install_font_scaling(cls.app)
        apply_font_scale(1.0, app=cls.app, save=False)

    def tearDown(self) -> None:
        apply_font_scale(1.0, app=self.app, save=False)

    def test_checkable_combobox_compact_minimum_size_and_clean_popup(self) -> None:
        container = QWidget()
        container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        lay = QHBoxLayout(container)
        combo = CheckableComboBox("WG")
        combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        combo.addItems([
            "VeryLongCompanyOrWorkingGroupNameThatShouldNotInflateButtonMinimumWidth_01",
            "SA2",
            "RAN2",
        ])
        lay.addWidget(combo)

        self.assertLess(combo.minimumSizeHint().width(), 160)
        self.assertEqual(combo.view().objectName(), "checkableComboPopupView")
        self.assertIsInstance(combo.view().itemDelegate(), CheckableComboItemDelegate)

        # Render first checked row and verify checkbox indicator is painted with ACCENT_PRIMARY (not solid black)
        img = QImage(220, 32, QImage.Format_ARGB32)
        img.fill(QColor(C.BG_CARD))
        painter = QPainter(img)
        opt = QStyleOptionViewItem()
        opt.widget = combo.view()
        opt.rect = QRect(0, 0, 220, 32)
        opt.state = QStyle.State_Enabled
        idx = combo.model().index(1, 0)
        opt.text = str(idx.data(Qt.DisplayRole) or "")
        combo.view().itemDelegate().paint(painter, opt, idx)
        painter.end()

        # Sample inside the rounded checkbox region (pad=8, box_sz=15 -> x in 10..20, y in 10..22)
        accent = QColor(C.ACCENT_PRIMARY)
        accent_hits = 0
        black_hits = 0
        for x in range(10, 22):
            for y in range(10, 22):
                px = QColor(img.pixel(x, y))
                if px == accent:
                    accent_hits += 1
                if px == QColor(0, 0, 0):
                    black_hits += 1
        self.assertGreater(accent_hits, 20, "Checked indicator should use C.ACCENT_PRIMARY fill")
        self.assertEqual(black_hits, 0, "Checked indicator must never paint solid black squares")

    def test_tdocs_window_responsive_width_and_column_visibility(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            mtg_dir = Path(tmp) / "SA2_178" / "Agenda"
            mtg_dir.mkdir(parents=True, exist_ok=True)
            dummy_xls = mtg_dir / "TDocs.xlsx"
            dummy_xls.write_text("", encoding="utf-8")

            rows = [
                {
                    "TDoc": "S2-2611863",
                    "Title": "[KI#1.1, Bullet all, Conclusion] Interim conclusion for KI#1.1 NAS Routing Alt 2 with long wrapped title.",
                    "Source": "Nokia, Huawei, HiSilicon, T-Mobile USA, AT&T, Interdigital, China Mobile, CUCC, Xiaomi, vivo, OPPO",
                    "Type": "pCR",
                    "For": "Approval",
                    "Agenda Item": "20.6.1",
                    "Sub-Item": "NAS Routing/transport",
                    "Related WIs": "FS_6G_ARC",
                    "TDoc Status": "available",
                    "Secretary Remarks": "Baseline\nRevised in pre-meeting 6G baseline assignments session to S2-2612555.",
                    "Related TDocs": "S2-2612555",
                },
                {
                    "TDoc": "S2-2611895",
                    "Title": "6G NAS Routing Options",
                    "Source": "Nokia",
                    "Type": "discussion",
                    "For": "Information",
                    "Agenda Item": "20.6.1",
                    "Sub-Item": "NAS Routing/transport",
                    "Related WIs": "FS_6G_ARC",
                    "TDoc Status": "available",
                    "Secretary Remarks": "",
                },
            ]
            test_settings = QSettings(str(Path(tmp) / "tdocs_filters.ini"), QSettings.IniFormat)
            win = TDocsWindow(
                {"wg_name": "SA2", "meeting_number": "178", "location": "Calgary"},
                rows,
                str(dummy_xls),
                filter_settings=test_settings,
            )
            win.sub_item_combo.setVisible(True)
            self.app.processEvents()

            # Status readouts live in bottom footer bar, freeing top header width
            self.assertIs(win.count_lbl.parent(), win._footer_status_widget)
            self.assertIs(win.last_mod_lbl.parent(), win._footer_status_widget)
            self.assertIsNot(win.count_lbl.parent(), win._header_container)

            # Window must fit easily within standard laptop/desktop widths without pushing columns off-screen
            self.assertLessEqual(win.minimumSizeHint().width(), 1050)

            win.resize(1920, 1040)
            win.show()
            self.app.processEvents()

            # Both top header bar and filter bar sit cleanly on 1 single row when maximized at 1920px
            self.assertLessEqual(win._header_container.height(), scale_px(44))
            self.assertLessEqual(win._filter_frame.height(), scale_px(52))

            # Narrowing the window wraps both bars cleanly without locking horizontal width
            win.resize(1024, 768)
            self.app.processEvents()
            self.assertGreater(win._header_container.heightForWidth(1000), scale_px(40))
            self.assertGreater(win._filter_frame.heightForWidth(1000), scale_px(48))

            win.resize(1920, 1040)
            self.app.processEvents()

            hdr = win.table.horizontalHeader()
            last_col = win.model.columnCount() - 1
            right_edge = hdr.sectionViewportPosition(last_col) + hdr.sectionSize(last_col)
            self.assertLessEqual(
                right_edge,
                win.table.viewport().width() + 2,
                f"Rightmost column ({last_col}) must stay inside 1920px viewport",
            )
            self.assertGreaterEqual(
                hdr.sectionSize(7),
                scale_px(400),
                "Secretary Remarks must expand generously when maximized",
            )

            # At default 1400x750 (1.0x and 1.25x scale), Secretary Remarks keeps >= scale_px(225) floor
            win.resize(1400, 750)
            self.app.processEvents()
            self.assertGreaterEqual(hdr.sectionSize(7), scale_px(225))
            self.assertGreaterEqual(hdr.sectionSize(11), scale_px(86))
            self.assertLessEqual(
                hdr.sectionViewportPosition(last_col) + hdr.sectionSize(last_col),
                win.table.viewport().width() + 2,
            )

            apply_font_scale(1.25, app=self.app, save=False)
            self.app.processEvents()
            self.assertGreaterEqual(
                hdr.sectionSize(7),
                scale_px(225),
                "Secretary Remarks must stay readable under scaled fonts instead of collapsing",
            )
            apply_font_scale(1.0, app=self.app, save=False)
            self.app.processEvents()

            # Manual column resize marks _user_resized_cols and survives subsequent window resizes
            self.assertFalse(win._user_resized_cols)
            hdr.resizeSection(7, 310)
            self.app.processEvents()
            self.assertTrue(win._user_resized_cols)
            win.resize(1600, 750)
            self.app.processEvents()
            self.assertEqual(hdr.sectionSize(7), 310)

            # Verify HtmlDelegate word-wraps long remarks instead of clipping horizontally
            delegate = win.html_delegate
            opt = QStyleOptionViewItem()
            opt.widget = win.table
            opt.font = win.table.font()
            opt.rect = QRect(0, 0, 600, 40)
            wide_h = delegate.sizeHint(opt, win.proxy.index(0, 7)).height()
            opt.rect = QRect(0, 0, 180, 40)
            narrow_h = delegate.sizeHint(opt, win.proxy.index(0, 7)).height()
            self.assertGreater(narrow_h, wide_h)

            # Multi-line row 0 expands above baseline 48px while brief row 1 remains compact
            win._adjust_table_row_heights()
            self.assertGreater(win.table.rowHeight(0), scale_px(48))
            self.assertEqual(win.table.rowHeight(1), scale_px(48))
            win.close()


if __name__ == "__main__":
    unittest.main()
