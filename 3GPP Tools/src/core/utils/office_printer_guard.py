# --- File: src/core/utils/office_printer_guard.py ---
"""
Guards headless Office COM automation (Word, PowerPoint) against offline or
unreachable network/WSD default printers.

When Windows' default printer points to an offline WSD/network printer,
starting Microsoft Word or PowerPoint via COM and opening/converting a
document can trigger a blocking modal dialog ("Waiting for printer connection..."
/ "Es wird auf die Druckerverbindung gewartet..."), stalling background exports.

This module inspects local printers offline via ``win32print.EnumPrinters`` (no
network calls) and provides:
1. ``configure_word_safe_printer(word_app)`` — switches a Word COM session to a
   local virtual printer (e.g. ``Microsoft Print to PDF``) via
   ``Dialogs(wdDialogFilePrintSetup)`` with ``DoNotSetAsSysDefault=True``.
2. ``temporary_virtual_default_printer()`` — re-entrant context manager that
   temporarily switches the Windows default printer to a local virtual printer
   only when the current default printer is a network/WSD/offline printer, and
   restores the original default printer unconditionally on exit.
"""

from __future__ import annotations

import atexit
from contextlib import contextmanager
import logging
import os
import threading
from typing import Iterator, Optional, Tuple

logger = logging.getLogger(__name__)

# Ordered by ubiquity and reliability on Windows 10/11
PREFERRED_VIRTUAL_PRINTERS = (
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "Foxit PDF Editor Printer",
    "Adobe PDF",
    "OneNote (Desktop)",
    "Send To OneNote 16",
    "Send To OneNote",
)

SAFE_VIRTUAL_PORTS = (
    "PORTPROMPT:",
    "FILE:",
    "NUL:",
    "SHRFAX:",
    "FOXIT_PDF:",
    "FOXIT_EVERNOTE:",
    "ADOBEPDF:",
)

RISKY_PORT_PREFIXES = (
    "WSD",
    "IP_",
    "\\\\",
    "HTTP",
    "IPP",
    "CNBJNP",
)

WD_DIALOG_FILE_PRINT_SETUP = 97
PRINTER_ATTRIBUTE_NETWORK = 0x00000010
PRINTER_ATTRIBUTE_WORK_OFFLINE = 0x00000400

_GUARD_LOCK = threading.RLock()
_GUARD_DEPTH = 0
_SAVED_DEFAULT_PRINTER: Optional[str] = None
_ACTIVE_VIRTUAL_PRINTER: Optional[str] = None
_ATEXIT_REGISTERED = False


def _restore_saved_printer_on_exit() -> None:
    """Safety net restoring the user's default printer if the process exits abnormally."""
    global _SAVED_DEFAULT_PRINTER, _ACTIVE_VIRTUAL_PRINTER, _GUARD_DEPTH
    if os.name != "nt":
        return
    with _GUARD_LOCK:
        saved = _SAVED_DEFAULT_PRINTER
        _SAVED_DEFAULT_PRINTER = None
        _ACTIVE_VIRTUAL_PRINTER = None
        _GUARD_DEPTH = 0
        if saved:
            try:
                import win32print
                win32print.SetDefaultPrinter(saved)
            except Exception:
                pass


def find_safe_local_virtual_printer() -> Optional[str]:
    """
    Returns the name of an installed local virtual printer (e.g. Microsoft Print to PDF)
    using purely local spooler metadata without querying network ports.
    """
    if os.name != "nt":
        return None
    try:
        import win32print
        local_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
    except Exception as exc:
        logger.debug("Could not enumerate local printers: %s", exc)
        return None

    by_name_lower = {}
    fallback_virtual: Optional[str] = None
    for info in local_printers:
        name = str(info.get("pPrinterName") or "").strip()
        port = str(info.get("pPortName") or "").strip()
        if not name:
            continue
        by_name_lower[name.casefold()] = (name, port)
        port_upper = port.upper()
        if fallback_virtual is None and any(port_upper.startswith(p) for p in SAFE_VIRTUAL_PORTS):
            fallback_virtual = name

    for pref in PREFERRED_VIRTUAL_PRINTERS:
        hit = by_name_lower.get(pref.casefold())
        if hit:
            return hit[0]

    return fallback_virtual


def is_default_printer_risky() -> Tuple[bool, Optional[str]]:
    """
    Checks whether the current Windows default printer is a network, WSD, IP, or
    work-offline printer that may stall Office COM initialization/open calls.

    Returns ``(is_risky, default_printer_name)``.
    """
    if os.name != "nt":
        return False, None
    try:
        import win32print
        default_name = str(win32print.GetDefaultPrinter() or "").strip()
        if not default_name:
            return False, None

        if default_name.startswith("\\\\"):
            return True, default_name

        local_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
        matched = None
        for info in local_printers:
            if str(info.get("pPrinterName") or "").strip().casefold() == default_name.casefold():
                matched = info
                break

        if matched is None:
            # Default printer is likely a per-user network connection (not local)
            return True, default_name

        attrs = int(matched.get("Attributes") or 0)
        port = str(matched.get("pPortName") or "").strip()
        port_upper = port.upper()

        if attrs & (PRINTER_ATTRIBUTE_WORK_OFFLINE | PRINTER_ATTRIBUTE_NETWORK):
            return True, default_name

        if any(port_upper.startswith(vp) for vp in SAFE_VIRTUAL_PORTS):
            return False, default_name

        if any(port_upper.startswith(rp) for rp in RISKY_PORT_PREFIXES):
            return True, default_name

        return False, default_name
    except Exception as exc:
        logger.debug("Default printer inspection skipped: %s", exc)
        return False, None


def configure_word_safe_printer(word_app, log: Optional[logging.Logger] = None) -> Optional[str]:
    """
    Switches a headless Word COM session to a safe local virtual printer using
    ``Dialogs(wdDialogFilePrintSetup)`` with ``DoNotSetAsSysDefault=True`` so
    the user's Windows default printer is left untouched.
    """
    if os.name != "nt" or word_app is None:
        return None
    virtual_printer = find_safe_local_virtual_printer()
    if not virtual_printer:
        return None

    try:
        dlg = word_app.Dialogs(WD_DIALOG_FILE_PRINT_SETUP)
        dlg.Printer = virtual_printer
        dlg.DoNotSetAsSysDefault = True
        dlg.Execute()
        (log or logger).debug("Configured Word COM session printer to '%s' (system default untouched).", virtual_printer)
        return virtual_printer
    except Exception as exc:
        (log or logger).debug("Word Dialogs(FilePrintSetup) skipped (%s).", exc)
        return None


@contextmanager
def temporary_virtual_default_printer(log: Optional[logging.Logger] = None) -> Iterator[Optional[str]]:
    """
    Thread-safe, re-entrant context manager that temporarily sets Windows' default
    printer to a local virtual printer ONLY if the active default printer is a
    network/WSD/offline printer, preventing Office COM modals such as
    'Waiting for printer connection...'. Always restores the original default
    printer when the outermost scope exits.
    """
    global _GUARD_DEPTH, _SAVED_DEFAULT_PRINTER, _ACTIVE_VIRTUAL_PRINTER, _ATEXIT_REGISTERED

    if os.name != "nt":
        yield None
        return

    switched_to: Optional[str] = None
    with _GUARD_LOCK:
        if not _ATEXIT_REGISTERED:
            atexit.register(_restore_saved_printer_on_exit)
            _ATEXIT_REGISTERED = True

        if _GUARD_DEPTH == 0:
            risky, current_default = is_default_printer_risky()
            if risky and current_default:
                virtual_printer = find_safe_local_virtual_printer()
                if virtual_printer and virtual_printer.casefold() != current_default.casefold():
                    try:
                        import win32print
                        win32print.SetDefaultPrinter(virtual_printer)
                        _SAVED_DEFAULT_PRINTER = current_default
                        _ACTIVE_VIRTUAL_PRINTER = virtual_printer
                        switched_to = virtual_printer
                        (log or logger).info(
                            "Temporarily switched Windows default printer from '%s' to virtual '%s' for Office COM export.",
                            current_default,
                            virtual_printer,
                        )
                    except Exception as exc:
                        (log or logger).debug("Could not temporarily switch default printer: %s", exc)
        else:
            switched_to = _ACTIVE_VIRTUAL_PRINTER

        _GUARD_DEPTH += 1

    try:
        yield switched_to
    finally:
        with _GUARD_LOCK:
            _GUARD_DEPTH = max(0, _GUARD_DEPTH - 1)
            if _GUARD_DEPTH == 0:
                _ACTIVE_VIRTUAL_PRINTER = None
                if _SAVED_DEFAULT_PRINTER:
                    restore_name = _SAVED_DEFAULT_PRINTER
                    _SAVED_DEFAULT_PRINTER = None
                    try:
                        import win32print
                        win32print.SetDefaultPrinter(restore_name)
                        (log or logger).debug("Restored Windows default printer to '%s'.", restore_name)
                    except Exception as exc:
                        (log or logger).warning("Failed to restore Windows default printer '%s': %s", restore_name, exc)
