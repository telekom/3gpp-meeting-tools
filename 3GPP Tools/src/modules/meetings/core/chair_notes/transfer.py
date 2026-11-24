"""HTTP directory probing, atomic `.part` download streaming, and ZIP extraction."""

import logging
import os
import time
import zipfile
from pathlib import Path
from typing import List, Tuple

from core.utils.paths import to_extended_path

from .selector import (
    SUBPATHS,
    looks_like_directory_listing,
    parse_file_list,
    prefer_clean_variants,
)


def format_byte_size(num_bytes: int) -> str:
    """Format byte counts compactly for toolbar progress labels."""
    if num_bytes < 1024 * 1024:
        return f"{max(1, num_bytes // 1024)} KB"
    return f"{num_bytes / (1024 * 1024):.1f} MB"


def locate_chair_notes_dir(
    session,
    candidate_base_urls: list,
    subpaths=None,
    progress_callback=None,
    cancel_check=None,
) -> Tuple[str, str]:
    """Probe candidate base URLs across Chairs_Notes / Chair_Notes subpaths."""
    paths_to_probe = SUBPATHS if subpaths is None else subpaths
    for base_url in candidate_base_urls:
        if cancel_check and cancel_check():
            return None, ""
        if not base_url:
            continue
        clean_base = base_url.rstrip("/")
        for sub in paths_to_probe:
            if cancel_check and cancel_check():
                return None, ""

            probe_url = f"{clean_base}/{sub}"
            try:
                if progress_callback:
                    progress_callback(f"Probing {probe_url[:35]}...")
                resp = session.get(probe_url, timeout=15)
                if resp.status_code == 200 and looks_like_directory_listing(resp.text):
                    parsed = parse_file_list(resp.text, probe_url)
                    if parsed or "ChairNotes" in resp.text or "Directory Listing" in resp.text:
                        logging.info(f"[ChairNotes] Found active directory listing at: {probe_url}")
                        return probe_url, resp.text
            except Exception as e:
                logging.debug(f"[ChairNotes] Failed to probe {probe_url}: {e}")
    return None, ""


def extract_chair_notes_zip(zip_path: Path, target_dir: Path) -> List[str]:
    """Extract `.docx`, `.doc`, `.htm`, `.html` files from a Chairman's Notes ZIP archive."""
    os.makedirs(to_extended_path(target_dir), exist_ok=True)
    extracted_files: List[Tuple[str, str]] = []
    try:
        with zipfile.ZipFile(to_extended_path(zip_path), "r") as zf:
            for member in zf.infolist():
                if member.is_dir():
                    continue
                member_name = os.path.basename(member.filename)
                if not member_name or member_name.startswith("~$") or member_name.startswith("."):
                    continue
                ext = Path(member_name).suffix.lower()
                if ext not in (".docx", ".doc", ".htm", ".html"):
                    continue

                out_path = target_dir / member_name
                with zf.open(member) as src, open(to_extended_path(out_path), "wb") as dst:
                    dst.write(src.read())
                extracted_files.append((member_name, str(out_path)))
                logging.info(f"[ChairNotes] Extracted from ZIP: {out_path}")
    except Exception as e:
        logging.warning(f"[ChairNotes] Could not extract ZIP {zip_path}: {e}")
        return []

    preferred = prefer_clean_variants(extracted_files)
    return [fname for fname, _ in preferred]


def download_files_sync(
    session,
    files_to_download: List[Tuple[str, str]],
    target_dir: Path,
    progress_callback=None,
    cancel_check=None,
    overwrite: bool = True,
    zip_extractor=None,
) -> List[str]:
    """Download `(filename, url)` tuples atomically via `.part` files with throttled progress."""
    os.makedirs(to_extended_path(target_dir), exist_ok=True)
    extract_fn = zip_extractor or extract_chair_notes_zip
    downloaded: List[str] = []
    total_files = len(files_to_download)

    for idx, (filename, file_url) in enumerate(files_to_download, 1):
        if cancel_check and cancel_check():
            break

        if progress_callback:
            progress_callback(f"({idx}/{total_files}) {filename}")
        dest_path = target_dir / filename
        part_path = target_dir / f"{filename}.part"

        if not overwrite and dest_path.exists() and dest_path.stat().st_size > 0:
            logging.info(f"[ChairNotes] Skipping already downloaded file: {dest_path}")
            if dest_path.suffix.lower() == ".zip":
                extracted_names = extract_fn(dest_path, target_dir)
                if extracted_names:
                    downloaded.extend(extracted_names)
                else:
                    downloaded.append(filename)
            else:
                downloaded.append(filename)
            continue

        completed_file = False
        try:
            with session.get(file_url, stream=True, timeout=60) as dl_resp:
                dl_resp.raise_for_status()
                try:
                    total_bytes = int(dl_resp.headers.get("Content-Length") or 0)
                except (TypeError, ValueError):
                    total_bytes = 0

                bytes_read = 0
                last_emit = time.monotonic()

                with open(to_extended_path(part_path), "wb") as f:
                    for chunk in dl_resp.iter_content(chunk_size=65536):
                        if cancel_check and cancel_check():
                            return downloaded
                        if chunk:
                            f.write(chunk)
                            bytes_read += len(chunk)
                            if progress_callback:
                                now = time.monotonic()
                                if now - last_emit >= 0.25:
                                    last_emit = now
                                    if total_bytes > 0:
                                        pct = min(100, int(bytes_read * 100 / total_bytes))
                                        progress_callback(
                                            f"({idx}/{total_files}) {pct}% "
                                            f"({format_byte_size(bytes_read)}/{format_byte_size(total_bytes)})"
                                        )
                                    else:
                                        progress_callback(
                                            f"({idx}/{total_files}) {format_byte_size(bytes_read)}"
                                        )

            os.replace(to_extended_path(part_path), to_extended_path(dest_path))
            completed_file = True
            logging.info(f"[ChairNotes] Saved: {dest_path}")

            if dest_path.suffix.lower() == ".zip":
                if progress_callback:
                    progress_callback(f"({idx}/{total_files}) Extracting...")
                extracted_names = extract_fn(dest_path, target_dir)
                if extracted_names:
                    downloaded.extend(extracted_names)
                else:
                    downloaded.append(filename)
            else:
                downloaded.append(filename)

        except Exception as dl_err:
            logging.error(f"[ChairNotes] Error downloading {filename}: {dl_err}")
        finally:
            if not completed_file:
                try:
                    ext_part = to_extended_path(part_path)
                    if os.path.exists(ext_part):
                        os.remove(ext_part)
                except OSError:
                    pass

    return downloaded
