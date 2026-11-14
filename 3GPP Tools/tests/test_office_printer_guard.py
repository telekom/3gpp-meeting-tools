# --- File: tests/test_office_printer_guard.py ---
"""
Unit tests for core.utils.office_printer_guard.
"""

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from core.utils import office_printer_guard as guard


class TestOfficePrinterGuard(unittest.TestCase):
    def setUp(self):
        guard._GUARD_DEPTH = 0
        guard._SAVED_DEFAULT_PRINTER = None

    def test_find_safe_local_virtual_printer_prefers_print_to_pdf(self):
        fake_printers = [
            {"pPrinterName": "Brother HL-L2340D series Printer", "pPortName": "WSD-1234", "Attributes": 0x240},
            {"pPrinterName": "OneNote (Desktop)", "pPortName": "nul:", "Attributes": 0x240},
            {"pPrinterName": "Microsoft Print to PDF", "pPortName": "PORTPROMPT:", "Attributes": 0x240},
        ]
        with patch.object(guard.os, "name", "nt"), \
             patch.dict("sys.modules", {"win32print": MagicMock(PRINTER_ENUM_LOCAL=2, EnumPrinters=MagicMock(return_value=fake_printers))}):
            self.assertEqual(guard.find_safe_local_virtual_printer(), "Microsoft Print to PDF")

    def test_is_default_printer_risky_detects_wsd_and_virtual(self):
        fake_printers = [
            {"pPrinterName": "Brother HL-L2340D series Printer", "pPortName": "WSD-1234", "Attributes": 0x240},
            {"pPrinterName": "Microsoft Print to PDF", "pPortName": "PORTPROMPT:", "Attributes": 0x240},
        ]
        fake_win32 = MagicMock(
            PRINTER_ENUM_LOCAL=2,
            EnumPrinters=MagicMock(return_value=fake_printers),
            GetDefaultPrinter=MagicMock(return_value="Brother HL-L2340D series Printer"),
        )
        with patch.object(guard.os, "name", "nt"), patch.dict("sys.modules", {"win32print": fake_win32}):
            risky, name = guard.is_default_printer_risky()
            self.assertTrue(risky)
            self.assertEqual(name, "Brother HL-L2340D series Printer")

        fake_win32.GetDefaultPrinter.return_value = "Microsoft Print to PDF"
        with patch.object(guard.os, "name", "nt"), patch.dict("sys.modules", {"win32print": fake_win32}):
            risky, name = guard.is_default_printer_risky()
            self.assertFalse(risky)
            self.assertEqual(name, "Microsoft Print to PDF")

    def test_temporary_virtual_default_printer_restores_on_exception_and_nesting(self):
        fake_printers = [
            {"pPrinterName": "Brother HL-L2340D series Printer", "pPortName": "WSD-1234", "Attributes": 0x240},
            {"pPrinterName": "Microsoft Print to PDF", "pPortName": "PORTPROMPT:", "Attributes": 0x240},
        ]
        fake_win32 = MagicMock(
            PRINTER_ENUM_LOCAL=2,
            EnumPrinters=MagicMock(return_value=fake_printers),
            GetDefaultPrinter=MagicMock(return_value="Brother HL-L2340D series Printer"),
        )
        with patch.object(guard.os, "name", "nt"), patch.dict("sys.modules", {"win32print": fake_win32}):
            with self.assertRaises(RuntimeError):
                with guard.temporary_virtual_default_printer() as outer_sw:
                    self.assertEqual(outer_sw, "Microsoft Print to PDF")
                    fake_win32.SetDefaultPrinter.assert_called_once_with("Microsoft Print to PDF")
                    with guard.temporary_virtual_default_printer() as inner_sw:
                        self.assertEqual(inner_sw, "Microsoft Print to PDF")
                        # Still only called once (re-entrant)
                        self.assertEqual(fake_win32.SetDefaultPrinter.call_count, 1)
                    # Not yet restored on inner exit
                    self.assertEqual(fake_win32.SetDefaultPrinter.call_count, 1)
                    raise RuntimeError("boom")

            # Must restore original default printer on outermost exception exit
            self.assertEqual(fake_win32.SetDefaultPrinter.call_args_list[-1][0][0], "Brother HL-L2340D series Printer")
            self.assertEqual(guard._GUARD_DEPTH, 0)
            self.assertIsNone(guard._SAVED_DEFAULT_PRINTER)

    def test_configure_word_safe_printer_sets_dialog_without_changing_sys_default(self):
        fake_printers = [
            {"pPrinterName": "Microsoft Print to PDF", "pPortName": "PORTPROMPT:", "Attributes": 0x240},
        ]
        fake_win32 = MagicMock(
            PRINTER_ENUM_LOCAL=2,
            EnumPrinters=MagicMock(return_value=fake_printers),
        )
        mock_dlg = MagicMock()
        mock_word = MagicMock()
        mock_word.Dialogs.return_value = mock_dlg

        with patch.object(guard.os, "name", "nt"), patch.dict("sys.modules", {"win32print": fake_win32}):
            res = guard.configure_word_safe_printer(mock_word)

        self.assertEqual(res, "Microsoft Print to PDF")
        mock_word.Dialogs.assert_called_once_with(guard.WD_DIALOG_FILE_PRINT_SETUP)
        self.assertEqual(mock_dlg.Printer, "Microsoft Print to PDF")
        self.assertTrue(mock_dlg.DoNotSetAsSysDefault)
        mock_dlg.Execute.assert_called_once()


if __name__ == "__main__":
    unittest.main()
