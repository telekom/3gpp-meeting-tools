"""Word WdSaveFormat constants, NTFS attribute cleanup, and safe COM lifecycle utilities."""

from __future__ import annotations

import ctypes
import faulthandler
import gc
import logging
import os
import stat
import subprocess
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Optional

import pythoncom

from core.utils.office_printer_guard import configure_word_safe_printer

logger = logging.getLogger(__name__)

# Word WdSaveFormat Constants
WD_FORMAT_DOC = 0
WD_FORMAT_XML_DOCX = 12      # wdFormatXMLDocument (.docx)
WD_FORMAT_DOC_DEFAULT = 16   # wdFormatDocumentDefault
WD_FORMAT_PDF = 17           # wdFormatPDF
WD_FORMAT_XPS = 18           # wdFormatXPS
WD_FORMAT_RTF = 6            # wdFormatRTF
WD_FORMAT_HTML = 8           # wdFormatHTML
WD_FORMAT_TXT = 2            # wdFormatText


def _sanitize_file_attributes(file_path: Path) -> None:
    """Removes read-only flags and NTFS Zone.Identifier streams using native Win32 APIs."""
    p = Path(file_path).resolve()
    if not p.exists():
        return

    try:
        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass

    if os.name == "nt":
        # Reset file attributes (removes Read-Only / Hidden flags)
        try:
            ctypes.windll.kernel32.SetFileAttributesW(str(p), 0x80)  # FILE_ATTRIBUTE_NORMAL
        except Exception:
            pass

        # Unlink the NTFS Zone.Identifier alternate data stream directly
        try:
            zone_stream = f"{str(p)}:Zone.Identifier"
            ctypes.windll.kernel32.DeleteFileW(zone_stream)
        except Exception:
            pass

        # Fallback via PowerShell Unblock-File
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", f'Unblock-File -LiteralPath "{str(p)}"'],
                creationflags=0x08000000,  # CREATE_NO_WINDOW
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=2,
            )
        except Exception:
            pass


def _is_rpc_disconnect_error(exc: BaseException) -> bool:
    """Detect Word COM/RPC disconnects after WINWORD becomes unavailable."""
    rpc_hresults = {
        -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
        -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE
        -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
        -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    }
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int) and hresult in rpc_hresults:
        return True
    msg = str(exc).lower()
    return any(marker in msg for marker in (
        "0x800706be", "0x800706ba", "0x80010108", "0x80010001",
        "remote procedure call failed", "rpc server is unavailable",
        "object invoked has disconnected", "call was rejected by callee",
    ))


@contextmanager
def _suppress_faulthandler_for_word_com():
    """
    Suppress only faulthandler's process-wide native dump during fragile Word
    export calls. Python/pywin32 exceptions still propagate normally.
    """
    try:
        was_enabled = faulthandler.is_enabled()
    except Exception:
        was_enabled = False
    try:
        if was_enabled:
            try:
                faulthandler.disable()
            except Exception:
                pass
        yield
    finally:
        if was_enabled:
            try:
                faulthandler.enable()
            except Exception:
                pass


def _safe_com_cleanup(word=None, doc=None, new_doc=None, com_server_dead: bool = False) -> None:
    """
    Safely closes Word documents, quits Word, releases COM pointers,
    and uninitializes the COM apartment without triggering RPC SEH faults.
    """
    # Temporarily shield faulthandler so harmless first-chance RPC server
    # disconnects during WINWORD exit are not dumped to stderr as fatal errors.
    was_faulthandler_enabled = False
    try:
        if faulthandler.is_enabled():
            was_faulthandler_enabled = True
            faulthandler.disable()
    except Exception:
        pass

    try:
        # A disconnected WINWORD RPC server must not receive further COM calls.
        if com_server_dead:
            new_doc = None
            doc = None
            word = None

        # 1. Wait for any background printing/exporting to complete
        if word is not None and not com_server_dead:
            try:
                while word.BackgroundPrintingStatus > 0:
                    time.sleep(0.05)
            except Exception:
                pass

        # 2. Safely close documents
        for d in (() if com_server_dead else (new_doc, doc)):
            if d is not None:
                try:
                    d.Close(SaveChanges=False)
                except Exception:
                    pass

        # 3. Safely close any lingering Protected View windows
        if word is not None and not com_server_dead:
            try:
                if hasattr(word, "ProtectedViewWindows"):
                    while word.ProtectedViewWindows.Count > 0:
                        word.ProtectedViewWindows.Item(1).Close()
            except Exception:
                pass

            # 4. Quit Word Application
            try:
                word.Quit(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            except Exception:
                pass

    finally:
        # 5. Clear all COM references in local scope
        new_doc = None
        doc = None
        word = None

        # 6. Force garbage collection so win32com wrapper destructors execute
        try:
            gc.collect()
        except Exception:
            pass

        # 7. Pump any remaining COM messages
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        # 8. Uninitialize COM apartment
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        # 9. Restore faulthandler if it was active
        if was_faulthandler_enabled:
            try:
                faulthandler.enable()
            except Exception:
                pass


def _is_valid_pdf(path: Path) -> bool:
    try:
        if not path.exists() or path.stat().st_size < 5:
            return False
        with path.open("rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def _configure_unattended_word(word, log: Optional[logging.Logger] = None) -> None:
    """Applies unattended, modal-free COM configuration (including virtual printer binding)."""
    word.Visible = False
    word.DisplayAlerts = 0
    configure_word_safe_printer(word, log=log or logger)
    try:
        word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        word.Options.ConfirmConversions = False
        word.Options.DoNotPromptForConvert = True
        word.Options.WarnBeforeSavingPrintingSendingMarkup = False
        word.Options.SaveInterval = 0
        word.Options.UpdateLinksAtOpen = False
        word.Options.UpdateFieldsAtPrint = False
        word.Options.UpdateLinksAtPrint = False
        word.Options.PrintBackground = False
    except Exception:
        pass
