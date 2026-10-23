# --- File: src/modules/meetings/core/tdoc_file_handler.py ---
from __future__ import annotations

import logging
import os
import shutil
import zipfile
from pathlib import Path
from typing import Callable, List, Optional

from core.network.session import NetworkSession
from core.utils.paths import to_extended_path


class TDocFileHandler:
    @staticmethod
    def download_and_extract_tdoc(
        target_filename: str,
        base_url: str,
        tdoc_dir: Path,
        timeout: int = 15,
        cancel_check: Optional[Callable[[], bool]] = None,
    ) -> List[Path]:
        """
        Downloads a TDoc ZIP file from the 3GPP FTP/HTTP server and extracts its contents.
        Applies atomic downloads (.part rename), archive integrity checks, and intelligent
        renaming to prevent base/revision collisions and corrupted archives on interruption.

        :param target_filename: The specific filename to download (e.g., 'S2-260123r01')
        :param base_url: The FTP/HTTP directory URL containing the ZIP.
        :param tdoc_dir: The local Path object where the files should be saved.
        :param timeout: Network timeout in seconds per URL attempt.
        :param cancel_check: Optional callback returning True if download should be cancelled.
        :return: A list of Path objects pointing to the extracted documents.
        """
        if cancel_check and cancel_check():
            raise InterruptedError(f"Download cancelled for {target_filename}")

        zip_path = tdoc_dir / f"{target_filename}.zip"
        part_path = tdoc_dir / f"{target_filename}.zip.part"
        ext_zip_path = to_extended_path(zip_path)
        ext_part_path = to_extended_path(part_path)
        ext_tdoc_dir = to_extended_path(tdoc_dir)

        # 1. Check existing local archive validity
        archive_ready = False
        if os.path.exists(ext_zip_path):
            if os.path.getsize(ext_zip_path) > 0 and zipfile.is_zipfile(ext_zip_path):
                archive_ready = True
                logging.debug(f"💾 [CACHE HIT] Found valid local archive: {zip_path.name}")
            else:
                logging.warning(f"⚠️ [CORRUPT ARCHIVE] Found corrupt/empty {zip_path.name}. Re-downloading...")
                try:
                    os.remove(ext_zip_path)
                except OSError:
                    pass

        # 2. Download if missing or corrupt
        if not archive_ready:
            os.makedirs(ext_tdoc_dir, exist_ok=True)
            dl_url = base_url.rstrip('/') + f"/{target_filename}.zip"

            logging.info(f"🌐 [HTTP GET] Requesting: {dl_url}")

            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            try:
                response = session.get(dl_url, stream=True, timeout=timeout)
                response.raise_for_status()

                total_bytes = 0
                with open(ext_part_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=32768):
                        if cancel_check and cancel_check():
                            f.close()
                            if os.path.exists(ext_part_path):
                                try:
                                    os.remove(ext_part_path)
                                except OSError:
                                    pass
                            raise InterruptedError(f"Download cancelled for {target_filename}")
                        if chunk:
                            f.write(chunk)
                            total_bytes += len(chunk)

                if os.path.exists(ext_part_path):
                    os.replace(ext_part_path, ext_zip_path)

                size_kb = total_bytes / 1024.0
                logging.info(f"✅ [HTTP 200] Successfully downloaded {target_filename}.zip ({size_kb:.1f} KB) from {dl_url}")
            except Exception:
                if os.path.exists(ext_part_path):
                    try:
                        os.remove(ext_part_path)
                    except OSError:
                        pass
                raise

        # 3. Extract and Rename
        extracted_files = []
        with zipfile.ZipFile(ext_zip_path, 'r') as z:
            for info in z.infolist():
                if cancel_check and cancel_check():
                    break
                if '__MACOSX' in info.filename or info.filename.startswith('._'):
                    continue

                if info.filename.lower().endswith(('.doc', '.docx', '.pdf', '.ppt', '.pptx')):
                    original_name = Path(info.filename).name

                    if target_filename.lower() not in original_name.lower():
                        safe_name = f"{target_filename}_{original_name}"
                    else:
                        safe_name = original_name

                    out_path = tdoc_dir / safe_name
                    ext_out_path = to_extended_path(out_path)
                    temp_out_path = to_extended_path(tdoc_dir / f"{safe_name}.part")

                    if not os.path.exists(ext_out_path):
                        try:
                            with z.open(info.filename) as source, open(temp_out_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                            os.replace(temp_out_path, ext_out_path)
                            logging.info(f"📂 [EXTRACTED] Saved: {out_path.name}")
                        except Exception:
                            if os.path.exists(temp_out_path):
                                try:
                                    os.remove(temp_out_path)
                                except OSError:
                                    pass
                            raise
                    else:
                        logging.debug(f"📄 [FILE READY] Already extracted: {out_path.name}")

                    extracted_files.append(out_path)

        return extracted_files
