# --- File: src/modules/word_tools/core/document_opener.py ---
import logging
import os
import platform
import webbrowser
from pathlib import Path
from typing import Union

from modules.word_tools.core.sensitivity_label import set_sensitivity_label

logger = logging.getLogger(__name__)

WORD_EXTENSIONS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"}


def _open_with_default_application(path: Path) -> None:
    """Open a file with the operating system's default application."""
    if hasattr(os, "startfile"):
        os.startfile(str(path))
    else:
        webbrowser.open(path.resolve().as_uri())


def _open_word_document(path: Path) -> None:
    """
    Open a Word document through COM, apply the configured sensitivity label
    when the document is currently unlabeled, persist a newly applied label,
    and leave the document open for the user.

    This function is intentionally Windows-only. Callers should normally use
    ``open_document`` so that non-Windows platforms and non-Word files retain
    their normal OS-opening behavior.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        # Dispatch attaches to an existing Word instance when possible and
        # starts Word when necessary. Unlike DispatchEx, this avoids creating
        # a separate Word process for every document the user opens.
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True

        doc = word.Documents.Open(
            FileName=os.path.normpath(str(path)),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=True,
        )

        label_changed = set_sensitivity_label(doc, return_changed=True)
        if label_changed:
            # The normal Word Tools configuration deliberately does not force
            # every caller to save. For an interactively opened document,
            # however, the newly assigned label must be persisted immediately.
            doc.Save()
            logger.info(
                "Applied configured sensitivity label and saved Word document: %s",
                path.name,
            )

        doc.Activate()
        word.Visible = True
    finally:
        pythoncom.CoUninitialize()


def open_document(file_path: Union[str, Path]) -> None:
    """
    Open a local document using the application's centralized opening policy.

    Word documents are opened through Microsoft Word COM so the configured
    sensitivity label can be applied. Other files use the operating system's
    default application. If Word automation fails, the Word file is still
    opened normally so sensitivity-label handling never blocks user access.
    """
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Document does not exist: {path}")

    is_word_document = path.suffix.lower() in WORD_EXTENSIONS
    if is_word_document and platform.system() == "Windows":
        try:
            _open_word_document(path)
            return
        except Exception as exc:
            logger.warning(
                "Could not open '%s' through Word COM / sensitivity-label handling; "
                "falling back to the default application: %s",
                path.name,
                exc,
                exc_info=True,
            )

    _open_with_default_application(path)
