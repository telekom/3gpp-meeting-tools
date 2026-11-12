"""Unit tests for Phase 1.1 SVG icons and collapsible grouped NavRail."""

from __future__ import annotations

import sys
import unittest

from PyQt5.QtWidgets import QApplication, QTabWidget, QWidget

from core.ui.icons import get_available_icon_names, lucide_icon, render_svg_pixmap
from core.ui.nav_rail import DEFAULT_NAV_ITEMS, NavRail
from core.ui.theme import B, S


class TestIconsAndNavRail(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_all_nav_icons_render_crisply(self):
        available = set(get_available_icon_names())
        for spec in DEFAULT_NAV_ITEMS:
            self.assertIn(spec.icon_name, available)
            pm = render_svg_pixmap(spec.icon_name, size=16)
            self.assertFalse(pm.isNull(), msg=f"Null pixmap for {spec.icon_name}")
            self.assertEqual(pm.devicePixelRatio(), 2.0)
            icon = lucide_icon(spec.icon_name, size=16)
            self.assertFalse(icon.isNull())

    def test_nav_rail_syncs_bidirectionally_with_hidden_tab_widget(self):
        tabs = QTabWidget()
        tabs.tabBar().hide()
        for spec in DEFAULT_NAV_ITEMS:
            tabs.addTab(QWidget(), spec.label)

        rail = NavRail()
        rail.tab_selected.connect(tabs.setCurrentIndex)
        tabs.currentChanged.connect(rail.set_active_index)

        # Clicking button in rail selects matching tab index
        rail._buttons_by_index[6].click()
        self.assertEqual(tabs.currentIndex(), 6)
        self.assertEqual(rail.active_index(), 6)

        # Programmatic tab switch (like internal deep links) highlights matching rail item
        tabs.setCurrentIndex(4)
        self.assertEqual(rail.active_index(), 4)
        self.assertTrue(rail._buttons_by_index[4].isChecked())
        self.assertFalse(tabs.tabBar().isVisibleTo(tabs))

    def test_mnemonic_ampersand_escaped_on_rail_button(self):
        rail = NavRail()
        btn = rail._buttons_by_index[1]
        self.assertEqual(btn.text(), "Visio && Office")
        self.assertIn("PowerPoint & Stencils", btn.toolTip())
        self.assertNotIn("&&", btn.toolTip())

    def test_manual_toggle_and_responsive_hysteresis(self):
        rail = NavRail()
        signals = []
        rail.collapsed_changed.connect(signals.append)

        # Expanded by default at normal desktop width
        self.assertFalse(rail.is_collapsed)
        self.assertEqual(rail.width(), S.RAIL_EXPANDED_WIDTH)

        # Shrink below 940 -> auto-collapses to 56px
        rail.apply_responsive_width(B.COMPACT_COLLAPSE_PX - 10)
        self.assertTrue(rail.is_collapsed)
        self.assertEqual(rail.width(), S.RAIL_COLLAPSED_WIDTH)
        self.assertEqual(signals, [True])

        # Stay inside hysteresis zone (950) -> stays collapsed (no jitter)
        rail.apply_responsive_width(950)
        self.assertTrue(rail.is_collapsed)

        # Widen >= 970 -> expands back to 224px
        rail.apply_responsive_width(B.COMPACT_EXPAND_PX + 20)
        self.assertFalse(rail.is_collapsed)
        self.assertEqual(rail.width(), S.RAIL_EXPANDED_WIDTH)

        # Manual collapse persists even when resizing above 970 after narrow
        rail.toggle_collapsed()
        self.assertTrue(rail.is_collapsed)
        self.assertTrue(rail.manual_collapsed)
        rail.apply_responsive_width(900)
        rail.apply_responsive_width(1200)
        self.assertTrue(rail.is_collapsed)

    def test_footer_buttons_register_and_collapse(self):
        from PyQt5.QtWidgets import QMenu, QPushButton

        rail = NavRail()
        help_btn = QPushButton("Help && Docs")
        help_btn.setMenu(QMenu(rail))
        settings_btn = QPushButton("Settings")
        settings_btn.setMenu(QMenu(rail))

        rail.register_footer_button(help_btn, "Help && Docs")
        rail.register_footer_button(settings_btn, "Settings")
        self.assertEqual(help_btn.text(), "Help && Docs")
        self.assertEqual(settings_btn.text(), "Settings")

        rail.set_collapsed(True, manual=True)
        self.assertEqual(help_btn.text(), "")
        self.assertEqual(settings_btn.text(), "")
        self.assertEqual(help_btn.property("collapsed"), "true")

        rail.set_collapsed(False, manual=True)
        self.assertEqual(help_btn.text(), "Help && Docs")
        self.assertEqual(settings_btn.text(), "Settings")


if __name__ == "__main__":
    unittest.main()
