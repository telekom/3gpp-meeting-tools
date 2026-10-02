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


def list_open_word_documents():
    """Return metadata for documents in the currently running Word instance."""
    if platform.system() != "Windows":
        return []

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return []

        documents = []
        for doc in word.Documents:
            try:
                name = str(doc.Name or "").strip() or "Untitled Document"
                full_name = str(doc.FullName or "").strip()
                saved = bool(doc.Saved)
                documents.append({
                    "name": name,
                    "full_name": full_name,
                    "modified": not saved,
                })
            except Exception:
                continue
        return documents
    finally:
        pythoncom.CoUninitialize()


def close_open_word_documents(excluded_paths=None, discard_changes: bool = False):
    """
    Close documents in the active Word instance while leaving exclusions open.

    Normal mode uses wdPromptToSaveChanges so Word retains its normal save
    confirmation for modified documents. Destructive mode uses
    wdDoNotSaveChanges explicitly. The Word application itself is never quit;
    if all documents close, Word decides its normal UI/session behaviour.
    """
    summary = {"closed": 0, "excluded": 0, "failed": []}
    if platform.system() != "Windows":
        return summary

    excluded = {str(x).casefold() for x in (excluded_paths or []) if x}

    import pythoncom
    import win32com.client

    wd_do_not_save_changes = 0
    wd_prompt_to_save_changes = -2

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return summary

        # Snapshot in reverse order because closing mutates Documents.
        docs = [word.Documents.Item(i) for i in range(word.Documents.Count, 0, -1)]
        for doc in docs:
            try:
                full_name = str(doc.FullName or "").strip()
                name = str(doc.Name or "").strip() or "Untitled Document"
                key = full_name.casefold() if full_name else name.casefold()
                if key in excluded:
                    summary["excluded"] += 1
                    continue

                save_option = wd_do_not_save_changes if discard_changes else wd_prompt_to_save_changes
                doc.Close(SaveChanges=save_option)
                summary["closed"] += 1
            except Exception as exc:
                summary["failed"].append(f"{name}: {exc}")

        return summary
    finally:
        pythoncom.CoUninitialize()
