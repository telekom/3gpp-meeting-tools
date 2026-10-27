# --- File: src/modules/word_tools/core/document_opener.py ---
import logging
import os
import platform
import webbrowser
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple, Union

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
    import gc
    import pythoncom
    import win32com.client

    word = None
    doc = None
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
        doc = word = None
        gc.collect()
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


def _snapshot_com_document(doc, doc_name: str, full_name: str, exists: bool) -> Optional[dict]:
    """
    Create a temporary disk snapshot of an unsaved or modified Word COM document
    using SaveCopyAs so the user's open session is not mutated.
    """
    import tempfile

    temp_dir = Path(tempfile.gettempdir()) / "3gpp_active_word_snapshots"
    temp_dir.mkdir(parents=True, exist_ok=True)
    safe_base = "".join(c if c.isalnum() or c in "._- " else "_" for c in doc_name).strip() or "Active_Document"
    snapshot_path = temp_dir / f"Active_{safe_base}"
    if snapshot_path.suffix.lower() not in WORD_EXTENSIONS:
        snapshot_path = snapshot_path.with_suffix(".docx")

    try:
        if hasattr(doc, "SaveCopyAs"):
            doc.SaveCopyAs(str(snapshot_path))
            logger.info("Created temporary snapshot of Word document: %s", snapshot_path)
            return {
                "name": doc_name,
                "display_name": f"🪟 {doc_name}*",
                "full_name": full_name or doc_name,
                "path": str(snapshot_path),
                "is_snapshot": True,
                "is_active": True,
                "modified": True,
            }
    except Exception as snap_err:
        logger.warning("SaveCopyAs snapshot failed for '%s': %s", doc_name, snap_err)

    if exists and full_name:
        return {
            "name": doc_name,
            "display_name": f"🪟 {doc_name}",
            "full_name": full_name,
            "path": full_name,
            "is_snapshot": False,
            "is_active": True,
            "modified": not bool(getattr(doc, "Saved", True)),
        }
    return None


def get_active_word_document(
    allow_unsaved_snapshot: bool = True,
    target_identifier: Optional[str] = None,
) -> Optional[dict]:
    """
    Return metadata and a usable disk path for the currently active Microsoft Word
    document (or a specific open document matching ``target_identifier``).

    If the document has unsaved edits or has never been saved to disk and
    ``allow_unsaved_snapshot`` is True, safely creates a temporary snapshot via
    ``SaveCopyAs`` so comparisons reflect the user's latest in-memory text without
    forcing a save on their working file.

    Returns:
        dict with keys: 'name', 'display_name', 'full_name', 'path',
        'is_snapshot', 'is_active', 'modified', or None if no document is open.
    """
    if platform.system() != "Windows":
        return None

    import gc
    import pythoncom
    import win32com.client

    word = None
    doc = None
    active_doc = None
    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            logger.info("No active Word.Application instance detected via COM.")
            return None

        try:
            if word.Documents.Count == 0:
                logger.info("Word application is running, but 0 documents are currently open.")
                return None
        except Exception:
            return None

        try:
            active_doc = word.ActiveDocument
        except Exception:
            active_doc = None

        if target_identifier:
            needle = target_identifier.strip().casefold()
            for candidate in word.Documents:
                try:
                    c_name = str(candidate.Name or "").strip()
                    c_full = str(candidate.FullName or "").strip()
                    if needle in (c_name.casefold(), c_full.casefold()):
                        doc = candidate
                        break
                except Exception:
                    continue
            if doc is None:
                doc = active_doc
        else:
            doc = active_doc

        if doc is None:
            return None

        doc_name = str(doc.Name or "").strip() or "Active_Document.docx"
        full_name = str(doc.FullName or "").strip()
        is_saved = bool(getattr(doc, "Saved", True))
        exists = False
        if full_name and not full_name.lower().startswith(("http://", "https://")):
            try:
                exists = Path(full_name).is_file()
            except Exception:
                exists = False

        is_active = False
        if active_doc is not None:
            try:
                is_active = str(active_doc.FullName or "") == full_name and str(active_doc.Name or "") == doc_name
            except Exception:
                is_active = (doc == active_doc)

        is_locked = False
        if exists:
            try:
                with open(full_name, "rb") as test_f:
                    test_f.read(16)
            except (PermissionError, OSError):
                is_locked = True

        if not exists or not is_saved or is_locked:
            if not allow_unsaved_snapshot:
                return {
                    "name": doc_name,
                    "display_name": f"🪟 {doc_name}*",
                    "full_name": full_name,
                    "path": full_name if exists else None,
                    "is_snapshot": False,
                    "is_active": is_active,
                    "modified": not is_saved,
                    "unsaved": not exists,
                    "is_locked": is_locked,
                }
            snap = _snapshot_com_document(doc, doc_name, full_name, exists)
            if snap is not None:
                snap["is_active"] = is_active
                snap["is_locked"] = is_locked
                return snap
            return None

        return {
            "name": doc_name,
            "display_name": f"🪟 {doc_name}",
            "full_name": full_name,
            "path": full_name,
            "is_snapshot": False,
            "is_active": is_active,
            "modified": False,
        }
    finally:
        active_doc = doc = word = None
        gc.collect()
        pythoncom.CoUninitialize()


def list_open_word_documents() -> List[dict]:
    """Return metadata for documents in the currently running Word instance."""
    if platform.system() != "Windows":
        return []

    import gc
    import pythoncom
    import win32com.client

    word = None
    doc = None
    active_doc = None
    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return []

        active_key = ""
        try:
            if word.Documents.Count > 0 and word.ActiveDocument is not None:
                active_doc = word.ActiveDocument
                active_key = f"{str(active_doc.FullName or '').strip()}|{str(active_doc.Name or '').strip()}".casefold()
        except Exception:
            active_key = ""

        documents = []
        for doc in word.Documents:
            try:
                name = str(doc.Name or "").strip() or "Untitled Document"
                full_name = str(doc.FullName or "").strip()
                saved = bool(doc.Saved)
                key = f"{full_name}|{name}".casefold()
                is_active = bool(active_key and key == active_key)
                exists = False
                if full_name and not full_name.lower().startswith(("http://", "https://")):
                    try:
                        exists = Path(full_name).is_file()
                    except Exception:
                        exists = False
                documents.append({
                    "name": name,
                    "full_name": full_name,
                    "modified": not saved,
                    "is_active": is_active,
                    "exists_on_disk": exists,
                })
            except Exception:
                continue

        # Sort so the active document is always listed first
        documents.sort(key=lambda d: (not d.get("is_active", False), d.get("name", "").lower()))
        return documents
    finally:
        active_doc = doc = word = None
        gc.collect()
        pythoncom.CoUninitialize()


def select_best_word_document(
    extracted_paths: Sequence[Union[str, Path]],
    target_tdoc: str = "",
) -> Optional[Path]:
    """
    Given a list of extracted file paths from a TDoc archive, select the best
    Word document (.docx, .doc, .docm) for comparison, prioritizing files whose
    stem matches ``target_tdoc`` and preferring .docx over legacy .doc when both exist.
    """
    if not extracted_paths:
        return None

    word_files: List[Path] = []
    for p in extracted_paths:
        path_obj = Path(p)
        if path_obj.suffix.lower() in WORD_EXTENSIONS and not path_obj.name.startswith("~$"):
            word_files.append(path_obj)

    if not word_files:
        return None

    if len(word_files) == 1:
        return word_files[0]

    clean_target = (target_tdoc or "").strip().lower()

    def _score(p: Path):
        name_lower = p.stem.lower()
        exact_match = 1 if (clean_target and name_lower == clean_target) else 0
        prefix_match = 1 if (clean_target and clean_target in name_lower) else 0
        is_docx = 1 if p.suffix.lower() == ".docx" else 0
        # Avoid cover sheets or presentation companions if main CR/spec doc exists
        is_cover = 1 if "cover" in name_lower else 0
        return (exact_match, prefix_match, -is_cover, is_docx)

    word_files.sort(key=_score, reverse=True)
    return word_files[0]


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

    import gc
    import pythoncom
    import win32com.client

    wd_do_not_save_changes = 0
    wd_prompt_to_save_changes = -2

    word = None
    docs = None
    doc = None
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
        doc = docs = word = None
        gc.collect()
        pythoncom.CoUninitialize()


def ensure_readable_word_document(
    file_path: Union[str, Path],
    logger_fn: Optional[Callable[[str], None]] = None,
) -> Tuple[Path, Optional[Path]]:
    """
    Ensures that a Word document (.docx, .doc, .docm) can be opened and read without
    file-sharing lock violations ([Errno 13] Permission denied), such as when the document
    is currently open in Microsoft Word or locked by OneDrive sync.

    1. Checks if the file can be opened directly in binary read mode.
    2. If opening fails with PermissionError / OSError or if Microsoft Word currently holds
       an open document matching this path/filename:
       - Uses Word COM SaveCopyAs to take a non-destructive temporary snapshot.
       - Returns (snapshot_path, snapshot_path) so the caller can convert the snapshot
         and delete it upon completion.
    3. If Word COM is not running or doesn't have the document, attempts a headless Word COM
       read-only open and SaveCopyAs fallback.
    4. If all snapshotting attempts fail, returns the original path (to let the underlying
       exception surface naturally).

    Returns:
        Tuple[Path, Optional[Path]]: (usable_path, temp_file_to_cleanup)
    """
    path = Path(file_path).resolve()
    log = logger_fn or (lambda msg: logger.info(msg))

    if not path.is_file():
        return path, None

    # 1. Quick check: can we open and read the first bytes?
    is_locked = False
    try:
        with open(path, "rb") as test_f:
            test_f.read(32)
    except (PermissionError, OSError) as exc:
        is_locked = True
        log(f"⚠️ Document '{path.name}' is locked by another process ({exc}). Attempting snapshot recovery...")

    # 2. Check if Microsoft Word is running on Windows and has this document open
    if platform.system() == "Windows":
        import tempfile
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            word = None
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                word = None

            matched_doc = None
            if word is not None:
                try:
                    target_name = path.name.casefold()
                    target_full = str(path).casefold()
                    for d in word.Documents:
                        try:
                            d_name = str(d.Name or "").strip().casefold()
                            d_full = str(d.FullName or "").strip().casefold()
                            if target_full == d_full or target_name == d_name:
                                matched_doc = d
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

            if matched_doc is not None:
                # If it's open in Word, snapshot it if locked or unsaved!
                is_saved = bool(getattr(matched_doc, "Saved", True))
                if is_locked or not is_saved:
                    log(f"🪟 Document '{matched_doc.Name}' is currently open in Word. Taking live snapshot...")
                    temp_dir = Path(tempfile.gettempdir()) / "3gpp_word_snapshots"
                    temp_dir.mkdir(parents=True, exist_ok=True)
                    safe_base = "".join(c if c.isalnum() or c in "._- " else "_" for c in path.stem).strip() or "Doc"
                    snap_name = f"Snapshot_{safe_base}_{os.getpid()}{path.suffix}"
                    snap_path = temp_dir / snap_name
                    try:
                        matched_doc.SaveCopyAs(str(snap_path))
                        if snap_path.is_file() and snap_path.stat().st_size > 0:
                            log(f"✅ Created live unlocked snapshot: {snap_name}")
                            return snap_path, snap_path
                    except Exception as snap_err:
                        log(f"⚠️ SaveCopyAs failed on open Word document: {snap_err}")

            # 3. If file is locked and active Word was not open or didn't have it,
            # try headless Word COM read-only open
            if is_locked:
                try:
                    log(f"⏳ Attempting headless Word COM snapshot for locked file '{path.name}'...")
                    word_headless = win32com.client.DispatchEx("Word.Application")
                    word_headless.Visible = False
                    word_headless.DisplayAlerts = 0
                    doc = word_headless.Documents.Open(
                        FileName=str(path),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                    )
                    temp_dir = Path(tempfile.gettempdir()) / "3gpp_word_snapshots"
                    temp_dir.mkdir(parents=True, exist_ok=True)
                    safe_base = "".join(c if c.isalnum() or c in "._- " else "_" for c in path.stem).strip() or "Doc"
                    snap_name = f"Snapshot_{safe_base}_{os.getpid()}{path.suffix}"
                    snap_path = temp_dir / snap_name
                    doc.SaveCopyAs(str(snap_path))
                    doc.Close(SaveChanges=False)
                    word_headless.Quit()
                    if snap_path.is_file() and snap_path.stat().st_size > 0:
                        log(f"✅ Created unlocked snapshot via Word COM: {snap_name}")
                        return snap_path, snap_path
                except Exception as com_err:
                    log(f"⚠️ Headless Word COM snapshot could not open locked file: {com_err}")
        finally:
            pythoncom.CoUninitialize()

    return path, None
