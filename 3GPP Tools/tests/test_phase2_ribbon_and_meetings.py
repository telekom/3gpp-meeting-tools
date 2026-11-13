# --- File: tests/test_phase2_ribbon_and_meetings.py ---
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QSizePolicy

from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon
from core.ui.toast import ToastManager, notify_info, notify_success
from modules.meetings.ui.meetings_actions import on_merger_finished, on_tdocs_caching_finished
from modules.meetings.ui.ui_tabs import MeetingsTab


class _HostWithToast(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.toast_manager = ToastManager(self)
        self.recorded = []

    def show_toast(self, title: str, message: str = "", variant: str = "info", actions=None, duration_ms: int = 6000):
        self.recorded.append((title, message, variant, actions))
        return self.toast_manager.show_toast(
            title=title,
            message=message,
            variant=variant,
            actions=actions,
            duration_ms=duration_ms,
        )


class TestPageRibbonAndToastHelpers(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_page_ribbon_tiers_and_lucide_icons(self) -> None:
        ribbon = PageRibbonCard("Test Workspace", "Subtitle hint")
        self.assertIsNotNone(ribbon.title_label)
        self.assertEqual(ribbon.sizePolicy().verticalPolicy(), QSizePolicy.Maximum)
        self.assertEqual(ribbon.tier1_widget.objectName(), "ribbonTier1")
        self.assertEqual(ribbon.primary_container.objectName(), "ribbonPrimaryContainer")
        self.assertEqual(ribbon.tier2_widget.objectName(), "ribbonTier2")
        self.assertTrue(ribbon.tier2_widget.isHidden())

        btn = QPushButton("Run")
        apply_lucide_icon(btn, "play", size=14)
        self.assertFalse(btn.icon().isNull())
        ribbon.add_primary_widget(btn)

        util_btn = QPushButton("Settings")
        ribbon.add_utility_widget(util_btn)
        self.assertTrue(ribbon.tier2_widget.isHidden())

        flt_btn = QPushButton("SA")
        ribbon.add_filter_widget(flt_btn)
        self.assertFalse(ribbon.tier2_widget.isHidden())
        self.assertFalse(ribbon.divider.isHidden())
        ribbon.close()

    def test_notify_helpers_dispatch_to_host_or_noop_standalone(self) -> None:
        standalone = QPushButton("Alone")
        self.assertIsNone(notify_success(standalone, "Done", "Saved cleanly"))
        self.assertIsNone(notify_info(standalone, "Info", "Notice"))

        host = _HostWithToast()
        child = QPushButton("Hosted", host)
        host.setCentralWidget(child)

        toast = notify_success(child, "Export Complete", "Wrote out.xlsx", actions=[("Open", lambda: None)])
        self.assertIsNotNone(toast)
        self.assertEqual(len(host.recorded), 1)
        self.assertEqual(host.recorded[0][0], "Export Complete")
        self.assertEqual(host.recorded[0][2], "success")
        host.toast_manager.dismiss_all()
        host.close()


class TestMeetingsPhase2Revamp(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication(sys.argv)
        cls.db_path = Path("src/db/3gpp_data.db")

    def test_meetings_tab_ribbon_card_icons_and_toasts(self) -> None:
        host = _HostWithToast()
        tab = MeetingsTab(self.db_path)
        host.setCentralWidget(tab)

        self.assertTrue(hasattr(tab, "ribbon_card"))
        self.assertTrue(hasattr(tab, "table_card"))
        self.assertFalse(tab.btn_open_last.icon().isNull())
        self.assertFalse(tab.sync_menu_btn.icon().isNull())
        self.assertFalse(tab.reports_menu_btn.icon().isNull())
        self.assertFalse(tab.btn_config.icon().isNull())
        self.assertFalse(tab.btn_reset_filters.icon().isNull())

        tab._on_agenda_download_finished(True, "Downloaded Agenda.csv")
        on_tdocs_caching_finished(tab, True, "Cached 12 documents")
        on_merger_finished(tab, True, "Merged 42 rows.\nSaved to:\nC:/temp/merged.xlsx")

        self.assertEqual(len(host.recorded), 3)
        self.assertEqual(host.recorded[0][0], "Agenda Download")
        self.assertEqual(host.recorded[1][0], "Caching Complete")
        self.assertEqual(host.recorded[2][0], "Merged TDocs Exported")
        self.assertEqual(host.recorded[2][3][0][0], "Open Excel")

        host.toast_manager.dismiss_all()
        tab.close()
        host.close()


if __name__ == "__main__":
    unittest.main()
