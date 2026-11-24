"""OpenXML formatting restore, file-lock waiting, and Word COM labeling helpers."""

from __future__ import annotations

import logging
import os
import sys
import time
import zipfile
from pathlib import Path

import pythoncom
import win32com.client

from core.utils.office_printer_guard import (
    configure_word_safe_printer,
    temporary_virtual_default_printer,
)
from core.utils.paths import to_extended_path
from modules.word_tools.core.sensitivity_label import set_sensitivity_label

logger = logging.getLogger(__name__)


def copy_formatting_parts(
    source: Path,
    target: Path,
    attempts: int = 12,
    delay: float = 0.5,
) -> None:
    """Restore source formatting parts after Word has fully released the output.

    Office and OneDrive can retain a file handle briefly after Document.Close()/
    Application.Quit(). Retry only sharing/permission failures; other ZIP/package
    errors are real failures and should surface immediately.
    """
    parts = (
        "word/styles.xml",
        "word/stylesWithEffects.xml",
        "word/theme/theme1.xml",
        "word/fontTable.xml",
    )
    tmp = target.with_name(target.name + ".formatting.tmp")
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            if tmp.exists():
                tmp.unlink()
            with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(
                target, "r"
            ) as dst, zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as out:
                source_names = set(src.namelist())
                replaced = set()
                for item in dst.infolist():
                    name = item.filename
                    if name in parts and name in source_names:
                        out.writestr(item, src.read(name))
                        replaced.add(name)
                    else:
                        out.writestr(item, dst.read(name))
                for name in parts:
                    if name in source_names and name not in replaced:
                        out.writestr(name, src.read(name))
                        replaced.add(name)
            os.replace(tmp, target)
            logger.info(
                "   -> Restored source formatting package parts: %s",
                ", ".join(sorted(replaced)),
            )
            return
        except PermissionError as exc:
            last_exc = exc
            try:
                if tmp.exists():
                    tmp.unlink()
            except OSError:
                pass
            if attempt == attempts:
                break
            logger.info(
                "   -> Output is still locked; retrying formatting restore (%d/%d)...",
                attempt,
                attempts,
            )
            time.sleep(delay)
    raise RuntimeError(
        f"Microsoft Word/OneDrive did not release the output file in time: {target}"
    ) from last_exc


def wait_for_file_unlocked(file_path: Path, max_wait: float = 6.0, delay: float = 0.2) -> bool:
    """Waits for Word or background indexers to release locks on file_path."""
    ext_path = to_extended_path(file_path)
    start = time.time()
    while time.time() - start < max_wait:
        try:
            with open(ext_path, "a+b"):
                return True
        except (PermissionError, OSError):
            time.sleep(delay)
    return False


def configure_word(word) -> None:
    """Apply safe headless automation flags on a Word COM Application instance."""
    word.Visible = False
    word.DisplayAlerts = 0
    configure_word_safe_printer(word, log=logger)
    try:
        word.AutomationSecurity = 3
        word.Options.ConfirmConversions = False
        word.Options.DoNotPromptForConvert = True
        word.Options.WarnBeforeSavingPrintingSendingMarkup = False
        word.Options.SaveInterval = 0
        word.Options.PrintBackground = False
    except Exception:
        pass


def is_rpc_disconnect(exc: BaseException) -> bool:
    """Return True if an exception reflects a dead or busy COM RPC server."""
    hresult = getattr(exc, "hresult", None)
    if hresult in (-2147023170, -2147023174, -2147417848, -2147418111):
        return True
    msg = str(exc).lower()
    return any(
        marker in msg
        for marker in (
            "0x800706be",
            "0x800706ba",
            "0x80010108",
            "0x80010001",
            "remote procedure call failed",
            "rpc server is unavailable",
            "object invoked has disconnected",
            "call was rejected by callee",
        )
    )


def apply_label_in_fresh_word(output: Path) -> None:
    """Apply the corporate label only after package patching is complete.

    A dedicated Word instance isolates sensitivity-label add-in/policy failures
    from the merge session. Label failure is deliberately fatal: returning an
    unlabeled corporate document would be misleading.
    """
    word = None
    doc = None
    server_dead = False
    facade = sys.modules.get("modules.word_tools.core.word_excerpt_extractor")
    apply_label = (
        getattr(facade, "set_sensitivity_label", set_sensitivity_label)
        if facade is not None
        else set_sensitivity_label
    )
    try:
        with temporary_virtual_default_printer(log=logger):
            word = win32com.client.DispatchEx("Word.Application")
            configure_word(word)
            logger.info("   -> Reopening finalized excerpt for sensitivity labeling...")
            doc = word.Documents.Open(
                os.path.normpath(str(output)),
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            apply_label(doc)
            try:
                from modules.word_tools.core.styles_3gpp import promote_3gpp_styles_in_word

                promote_3gpp_styles_in_word(doc)
            except Exception as e:
                logger.debug("Failed promoting 3GPP styles during labeling: %s", e)
            doc.Save()
            doc.Close(SaveChanges=False)
            doc = None
            word.Quit(SaveChanges=0)
            word = None
        logger.info("   -> Sensitivity label applied successfully.")
    except Exception as exc:
        server_dead = is_rpc_disconnect(exc)
        raise RuntimeError(
            "The excerpt was created and formatted, but Microsoft Word failed "
            "while applying the required corporate sensitivity label. The output "
            "was not reported as successfully completed."
        ) from exc
    finally:
        if not server_dead:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception as exc:
                    if is_rpc_disconnect(exc):
                        server_dead = True
            if word is not None and not server_dead:
                try:
                    word.Quit(SaveChanges=0)
                except Exception:
                    pass
        doc = None
        word = None
        try:
            import gc

            gc.collect()
        except Exception:
            pass
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        wait_for_file_unlocked(output)
