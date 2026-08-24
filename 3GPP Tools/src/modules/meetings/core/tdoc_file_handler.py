# --- File: src/modules/meetings/core/tdoc_file_handler.py ---
import logging
import shutil
import zipfile
from pathlib import Path

from core.network.session import NetworkSession


class TDocFileHandler:
    @staticmethod
    def download_and_extract_tdoc(target_filename: str, base_url: str, tdoc_dir: Path, timeout: int = 6) -> list:
        """
        Downloads a TDoc ZIP file from the 3GPP FTP/Local server and extracts its contents.
        Applies intelligent renaming to prevent base/revision collisions.

        :param target_filename: The specific filename to download (e.g., 'S2-260123r01')
        :param base_url: The FTP/HTTP directory URL containing the ZIP.
        :param tdoc_dir: The local Path object where the files should be saved.
        :param timeout: Network timeout in seconds per URL attempt.
        :return: A list of Path objects pointing to the extracted documents.
        """
        zip_path = tdoc_dir / f"{target_filename}.zip"

        # 1. Download if missing
        if not zip_path.exists():
            tdoc_dir.mkdir(parents=True, exist_ok=True)
            dl_url = base_url.rstrip('/') + f"/{target_filename}.zip"

            logging.info(f"🌐 [HTTP GET] Requesting: {dl_url}")

            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            response = session.get(dl_url, stream=True, timeout=timeout)
            response.raise_for_status()

            total_bytes = 0
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=16384):
                    if chunk:
                        f.write(chunk)
                        total_bytes += len(chunk)

            size_kb = total_bytes / 1024.0
            logging.info(f"✅ [HTTP 200] Successfully downloaded {target_filename}.zip ({size_kb:.1f} KB) from {dl_url}")
        else:
            logging.info(f"💾 [CACHE HIT] Found existing local archive: {zip_path.name} (Skipping network request)")

        # 2. Extract and Rename
        extracted_files = []
        with zipfile.ZipFile(zip_path, 'r') as z:
            for info in z.infolist():
                if '__MACOSX' in info.filename or info.filename.startswith('._'):
                    continue

                if info.filename.lower().endswith(('.doc', '.docx', '.pdf', '.ppt', '.pptx')):
                    original_name = Path(info.filename).name

                    if target_filename.lower() not in original_name.lower():
                        safe_name = f"{target_filename}_{original_name}"
                    else:
                        safe_name = original_name

                    out_path = tdoc_dir / safe_name

                    if not out_path.exists():
                        with z.open(info.filename) as source, open(out_path, 'wb') as target:
                            shutil.copyfileobj(source, target)
                        logging.info(f"📂 [EXTRACTED] Saved: {out_path.name}")
                    else:
                        logging.debug(f"📄 [FILE READY] Already extracted: {out_path.name}")

                    extracted_files.append(out_path)

        return extracted_files