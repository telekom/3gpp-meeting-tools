# --- File: src/modules/meetings/core/tdocs_package_exporter.py ---
import datetime
import logging
import re
import shutil
from pathlib import Path

import requests
from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.url_router import URLRouter

logger = logging.getLogger(__name__)


class TDocsPackageExporterThread(QThread):
    """Download/reuse, extract, and copy a set of base TDocs into one export package."""

    progress = pyqtSignal(str)
    # (success, export_dir_or_error, failed_tdoc_ids)
    finished = pyqtSignal(bool, str, object)

    def __init__(
        self,
        meeting_dir: Path,
        rows_data: list,
        mtg_info: dict,
        main_ftp_url: str = "",
        scope_label: str = "visible",
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.rows_data = list(rows_data or [])
        self.mtg_info = mtg_info or {}
        self.main_ftp_url = str(main_ftp_url or "").rstrip("/")
        self.scope_label = str(scope_label or "visible")

    def run(self):
        try:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            export_dir = self.meeting_dir / "Export" / f"TDocs_{timestamp}"
            export_dir.mkdir(parents=True, exist_ok=False)

            total = len(self.rows_data)
            failures = []
            is_active = bool(self.mtg_info.get("is_active_sync", False))
            wg_name = str(self.mtg_info.get("wg_name", "")).strip()
            folder_name = str(
                self.mtg_info.get("folder_name")
                or self.mtg_info.get("meeting_number", "")
            ).strip()

            logger.info(
                "[TDocs Package Export] Starting %s export of %d TDoc(s) to %s",
                self.scope_label, total, export_dir,
            )

            for idx, row in enumerate(self.rows_data, start=1):
                if self.isInterruptionRequested():
                    self.finished.emit(False, "TDoc package export was cancelled.", failures)
                    return

                tdoc_id = str(row.get("TDoc", "")).strip().upper()
                if not tdoc_id:
                    continue

                # The package operation intentionally exports the base/selected row only.
                # If a revision-like ID ever appears as a row, normalize it to its base.
                base_match = re.match(
                    r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$",
                    tdoc_id,
                    re.IGNORECASE,
                )
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id
                target_filename = base_tdoc

                self.progress.emit(f"TDoc {idx}/{total}: {target_filename}")

                routes = URLRouter.build_priority_url_list(
                    wg_name=wg_name,
                    folder_name=folder_name,
                    main_ftp_url=self.main_ftp_url,
                    is_active_sync=is_active,
                    target_filename=target_filename,
                )

                source_dir = self.meeting_dir / base_tdoc
                extracted_files = self._resolve_tdoc(target_filename, source_dir, routes)
                if not extracted_files:
                    failures.append(target_filename)
                    logger.warning("[TDocs Package Export] Could not retrieve %s", target_filename)
                    continue

                target_dir = export_dir / target_filename
                target_dir.mkdir(parents=True, exist_ok=True)

                copied = 0
                for source_path in extracted_files:
                    source_path = Path(source_path)
                    if not source_path.is_file():
                        continue
                    target_path = target_dir / source_path.name
                    shutil.copy2(str(source_path), str(target_path))
                    copied += 1

                if copied == 0:
                    failures.append(target_filename)
                    logger.warning("[TDocs Package Export] No extracted files available to copy for %s", target_filename)

            logger.info(
                "[TDocs Package Export] File stage complete: %d requested, %d failed, output=%s",
                total, len(failures), export_dir,
            )
            self.finished.emit(True, str(export_dir), failures)

        except FileExistsError:
            self.finished.emit(False, "Could not create a unique export folder. Please retry the export.", [])
        except Exception as exc:
            logger.error("[TDocs Package Export] Export failed: %s", exc, exc_info=True)
            self.finished.emit(False, str(exc), [])

    def _resolve_tdoc(self, target_filename: str, source_dir: Path, routes: list) -> list:
        """Reuse the normal TDoc cache/extraction helper and normal route priority."""
        last_error = "No candidate routes available."

        # A cached ZIP is route-independent. TDocFileHandler will skip the network
        # request when it finds this archive, so one route is enough to trigger extraction.
        cached_zip = source_dir / f"{target_filename}.zip"
        if cached_zip.exists():
            try:
                return TDocFileHandler.download_and_extract_tdoc(
                    target_filename,
                    routes[0] if routes else "http://cache.invalid",
                    source_dir,
                    timeout=6,
                )
            except Exception as exc:
                last_error = str(exc)
                logger.warning(
                    "[TDocs Package Export] Cached archive for %s could not be extracted: %s",
                    target_filename, exc,
                )

        for route in routes:
            if self.isInterruptionRequested():
                return []
            try:
                files = TDocFileHandler.download_and_extract_tdoc(
                    target_filename, route, source_dir, timeout=6
                )
                if files:
                    return files
            except requests.exceptions.HTTPError as exc:
                if exc.response is not None and exc.response.status_code == 404:
                    last_error = f"404 Not Found at {route}"
                    continue
                last_error = str(exc)
            except requests.exceptions.ConnectTimeout:
                last_error = f"Connection timed out at {route}"
            except Exception as exc:
                last_error = str(exc)

        logger.warning(
            "[TDocs Package Export] All routes exhausted for %s. Last error: %s",
            target_filename, last_error,
        )
        return []
