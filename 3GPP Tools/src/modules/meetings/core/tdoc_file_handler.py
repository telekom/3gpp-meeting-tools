# --- File: src/modules/meetings/core/tdoc_file_handler.py ---
import shutil
import zipfile
from pathlib import Path

from core.network.session import NetworkSession


class TDocFileHandler:
    @staticmethod
    def download_and_extract_tdoc(target_filename: str, base_url: str, tdoc_dir: Path) -> list:
        """
        Downloads a TDoc ZIP file from the 3GPP FTP and extracts its contents.
        Applies intelligent renaming to prevent base/revision collisions.

        :param target_filename: The specific filename to download (e.g., 'S2-260123r01')
        :param base_url: The FTP directory URL containing the ZIP.
        :param tdoc_dir: The local Path object where the files should be saved.
        :return: A list of Path objects pointing to the extracted documents.
        """
        zip_path = tdoc_dir / f"{target_filename}.zip"

        # 1. Download if missing
        if not zip_path.exists():
            tdoc_dir.mkdir(parents=True, exist_ok=True)
            dl_url = base_url.rstrip('/') + f"/{target_filename}.zip"

            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)
            response = session.get(dl_url, stream=True, timeout=30)
            response.raise_for_status()

            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=16384):
                    if chunk:
                        f.write(chunk)

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

                    extracted_files.append(out_path)

        return extracted_files