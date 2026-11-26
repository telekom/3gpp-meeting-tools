import ctypes
import os
import sys
import unittest
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QWidget

from core.ui.ui_components import (
    apply_native_window_icon,
    create_app_icon,
    get_app_icon_path,
    install_window_icon_manager,
)


class TestTaskbarIconReliability(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)
        cls.app.setWindowIcon(create_app_icon())
        install_window_icon_manager(cls.app)

    def test_app_icon_has_all_standard_dpi_sizes(self):
        icon = create_app_icon()
        self.assertFalse(icon.isNull(), "Application QIcon must not be null")

        sizes = {(s.width(), s.height()) for s in icon.availableSizes()}
        for expected in (16, 20, 24, 32, 40, 48, 64, 128, 256):
            self.assertIn(
                (expected, expected),
                sizes,
                f"create_app_icon() should provide {expected}x{expected} resolution for Windows DPI scaling",
            )

    def test_ico_file_exists_and_is_multi_resolution(self):
        ico_path = get_app_icon_path()
        self.assertIsNotNone(ico_path)
        self.assertTrue(Path(ico_path).is_file())
        data = Path(ico_path).read_bytes()
        self.assertGreater(len(data), 4096)
        # ICO header: reserved=0, type=1, count>=5
        import struct

        reserved, ico_type, count = struct.unpack("<HHH", data[:6])
        self.assertEqual(reserved, 0)
        self.assertEqual(ico_type, 1)
        self.assertGreaterEqual(count, 5, "3gpp_tools.ico must embed multiple icon frames (16px..256px)")

    @unittest.skipUnless(os.name == "nt", "Windows-only native HICON / IPropertyStore test")
    def test_native_win32_hicon_and_property_store_on_windows(self):
        import ctypes.wintypes as wt

        user32 = ctypes.windll.user32
        user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
        user32.SendMessageW.restype = wt.LPARAM
        WM_GETICON = 0x007F
        ICON_SMALL = 0
        ICON_BIG = 1

        # 1. Test Main Window
        main_win = QMainWindow()
        apply_native_window_icon(main_win, set_class_icon=True)
        main_win.show()
        QApplication.processEvents()

        hwnd_main = int(main_win.winId())
        self.assertNotEqual(hwnd_main, 0)
        self.assertFalse(main_win.windowIcon().isNull())
        self.assertNotEqual(user32.SendMessageW(hwnd_main, WM_GETICON, ICON_BIG, 0), 0)
        self.assertNotEqual(user32.SendMessageW(hwnd_main, WM_GETICON, ICON_SMALL, 0), 0)

        # 2. Test standalone QDialog that calls setWindowFlags(Qt.Window) without explicit setWindowIcon
        floating_dlg = QDialog(None)
        floating_dlg.setWindowFlags(Qt.Window)
        floating_dlg.show()
        QApplication.processEvents()

        hwnd_dlg = int(floating_dlg.winId())
        self.assertNotEqual(hwnd_dlg, 0)
        self.assertFalse(
            floating_dlg.windowIcon().isNull(),
            "Global _WindowIconEventFilter should automatically stamp top-level Qt.Window dialogs",
        )
        self.assertNotEqual(user32.SendMessageW(hwnd_dlg, WM_GETICON, ICON_BIG, 0), 0)
        self.assertNotEqual(user32.SendMessageW(hwnd_dlg, WM_GETICON, ICON_SMALL, 0), 0)

        # 3. Test standalone QWidget window (e.g. TDocsWindow / SessionPlanWindow pattern)
        standalone_widget = QWidget(None)
        standalone_widget.setWindowFlags(Qt.Window)
        standalone_widget.show()
        QApplication.processEvents()

        hwnd_w = int(standalone_widget.winId())
        self.assertFalse(standalone_widget.windowIcon().isNull())
        self.assertNotEqual(user32.SendMessageW(hwnd_w, WM_GETICON, ICON_BIG, 0), 0)
        self.assertNotEqual(user32.SendMessageW(hwnd_w, WM_GETICON, ICON_SMALL, 0), 0)

        standalone_widget.close()
        floating_dlg.close()
        main_win.close()


if __name__ == "__main__":
    unittest.main()
