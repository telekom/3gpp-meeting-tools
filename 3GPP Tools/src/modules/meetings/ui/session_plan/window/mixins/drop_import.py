# --- File: src/modules/meetings/ui/session_plan/window/mixins/drop_import.py ---
"""
Drag & drop ingestion (incl. Outlook attachments), manual import via URL/file/paste,
session-plan download, and the snapshot diff dialog.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import os
import re
from pathlib import Path
from typing import List, Tuple
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffEngine
from modules.meetings.core.session_plan.session_plan_models import SessionPlanSnapshot
from modules.meetings.ui.session_plan.session_plan_dialogs import SessionPlanDiffDialog, SessionPlanImportDialog
from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
from modules.meetings.ui.session_plan.window.constants import logger


class DropImportMixin:
    """Drag & drop ingestion (incl. Outlook attachments), manual import via URL/file/paste,"""

    # --- Drag & Drop Direct Ingestion ---

    def _is_valid_drop_data(self, mime_data) -> bool:
        """Determines if the drag payload contains a session plan file, Outlook attachment, or table text."""
        if not mime_data:
            return False

        # 1. Local URLs (Files dragged from Explorer, Desktop, or temp folder)
        if mime_data.hasUrls():
            for url in mime_data.urls():
                p = Path(url.toLocalFile())
                ext = p.suffix.lower()
                if ext in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"]:
                    return True
            return False

        # 2. Virtual Outlook attachments (FileGroupDescriptorW / FileContents)
        formats = mime_data.formats()
        for f in formats:
            if "FileGroupDescriptor" in f:
                return True

        # 3. Text or HTML table dragged from email body
        if mime_data.hasText():
            txt = mime_data.text()
            if any(k in txt.lower() for k in ("stream 1", "stream 2", "main room", "breakout 1", "breakout 2", "q1", "q2", "monday")):
                return True

        return False

    def _extract_outlook_attachments(self, mime_data) -> List[Path]:
        """Extracts virtual file attachments dragged directly from Microsoft Outlook."""
        import struct
        formats = mime_data.formats()
        fgd_format = None
        is_unicode = True
        for f in formats:
            if "FileGroupDescriptorW" in f:
                fgd_format = f
                is_unicode = True
                break
            elif "FileGroupDescriptor" in f:
                fgd_format = f
                is_unicode = False
                break

        if not fgd_format:
            return []

        fc_format = None
        for f in formats:
            if "FileContents" in f:
                fc_format = f
                break

        if not fc_format:
            return []

        descriptor_bytes = bytes(mime_data.data(fgd_format))
        if len(descriptor_bytes) < 4:
            return []

        num_files = struct.unpack("<I", descriptor_bytes[:4])[0]
        extracted = []
        contents_bytes = bytes(mime_data.data(fc_format))
        rec_size = 592 if is_unicode else 332
        name_offset = 72

        temp_dir = self.meeting_dir / "Agenda" / "SessionPlans" / "Dropped"
        temp_dir.mkdir(parents=True, exist_ok=True)

        for i in range(min(num_files, 5)):
            rec_start = 4 + i * rec_size
            if rec_start + rec_size > len(descriptor_bytes):
                break

            raw_name = descriptor_bytes[rec_start + name_offset : rec_start + rec_size]
            if is_unicode:
                filename = raw_name.decode("utf-16le", errors="ignore").split("\0")[0].strip()
            else:
                filename = raw_name.decode("mbcs", errors="ignore").split("\0")[0].strip()

            ext = Path(filename).suffix.lower()
            if ext in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"]:
                target_path = temp_dir / filename
                target_path.write_bytes(contents_bytes)
                extracted.append(target_path)
                break  # Typically 1 attachment dropped at a time

        return extracted

    def _extract_drop_payloads(self, mime_data) -> List[Tuple[str, str]]:
        """
        Extracts imported items as a list of (mode, payload) tuples:
        - ('file', file_path_str) for local files or extracted Outlook attachments
        - ('paste', text_payload) for dragged table text
        """
        results = []

        # 1. Local files
        if mime_data.hasUrls():
            for url in mime_data.urls():
                p = Path(url.toLocalFile())
                if p.suffix.lower() in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"] and p.is_file():
                    results.append(("file", str(p)))

        # 2. Outlook virtual attachments
        if not results:
            outlook_files = self._extract_outlook_attachments(mime_data)
            for of in outlook_files:
                results.append(("file", str(of)))

        # 3. Dragged email text or HTML table
        if not results and mime_data.hasText():
            txt = mime_data.text().strip()
            if any(k in txt.lower() for k in ("stream 1", "stream 2", "main room", "breakout 1", "breakout 2", "q1", "q2", "monday")):
                results.append(("paste", txt))

        return results

    def dragEnterEvent(self, event):
        if self._is_valid_drop_data(event.mimeData()):
            if hasattr(self, "drop_overlay") and self.drop_overlay:
                self.drop_overlay.setGeometry(self.rect())
                self.drop_overlay.show()
                self.drop_overlay.raise_()
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if self._is_valid_drop_data(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        if hasattr(self, "drop_overlay"):
            self.drop_overlay.hide()
        event.accept()

    def dropEvent(self, event):
        if hasattr(self, "drop_overlay"):
            self.drop_overlay.hide()

        payloads = self._extract_drop_payloads(event.mimeData())
        if payloads:
            event.acceptProposedAction()
            for mode, payload in payloads:
                self._handle_manual_import(mode, payload)
        else:
            event.ignore()

    # --- Ingestion Entry Points ---

    def _import_via_url(self):
        dlg = SessionPlanImportDialog(self.meeting_number, self)
        dlg.import_requested.connect(self._handle_manual_import)
        dlg.exec_()

    def _import_via_file(self):
        f_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Session Plan Document",
            str(self.meeting_dir / "Agenda" / "SessionPlans"),
            "Session Plan Files (*.docx *.doc *.zip *.htm *.html *.txt);;All Files (*.*)",
        )
        if f_path:
            self._handle_manual_import("file", f_path)

    def _import_via_paste(self):
        dlg = SessionPlanImportDialog(self.meeting_number, self)
        dlg.import_requested.connect(self._handle_manual_import)
        dlg.exec_()

    def _handle_manual_import(self, mode: str, payload: str):
        try:
            if mode == "file":
                p = Path(payload)

                # Check if dropped file is an agenda.csv
                if p.name.lower() in ("agenda.csv", "agenda.txt"):
                    target_csv = self.meeting_dir / "Agenda" / "agenda.csv"
                    import shutil
                    shutil.copy2(p, target_csv)
                    self.agenda_map = AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")
                    self._render_current_views()
                    QMessageBox.information(self, "Agenda Updated", f"Successfully loaded {len(self.agenda_map)} Agenda Items from agenda.csv!")
                    return

                snap = self._import_local_session_plan_path(p)
            elif mode == "paste":
                snap = SessionPlanParser.parse_text_or_markdown(payload, meeting_number=self.meeting_number)
            else:
                raw_input = payload.strip()
                if re.match(r"^https?://", raw_input, re.IGNORECASE):
                    downloaded = self._download_session_plan_url(raw_input)
                    m_tdoc = re.search(r"([A-Za-z]\d-\d{6,8})", downloaded.name)
                    tdoc_hint = m_tdoc.group(1).upper() if m_tdoc else "SessionPlan"
                    snap = self._import_local_session_plan_path(downloaded, tdoc_hint=tdoc_hint)
                else:
                    # TDoc# lookup across local meeting server (if active) and 3GPP FTP
                    from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
                    from modules.meetings.core.url_router import URLRouter
                    from core.network.network_state import NetworkState

                    tdoc_candidate = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", raw_input.upper()).strip(" .")
                    if not tdoc_candidate:
                        QMessageBox.warning(self, "Invalid Input", "Please enter a valid 3GPP URL or TDoc number.")
                        return

                    wg_name = self.mtg_info.get("wg_name", "").upper()
                    candidate_urls = []
                    if NetworkState.get_instance().is_local_active():
                        candidate_urls.append(URLRouter._get_local_server_base(wg_name) + "/Docs")
                    docs_url = self.mtg_info.get("docs_folder_url", "")
                    if docs_url:
                        if not docs_url.startswith("http"):
                            docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip("/")
                        candidate_urls.append(docs_url)

                    tdoc_dir = self.meeting_dir / tdoc_candidate
                    extracted = []
                    for u in candidate_urls:
                        try:
                            extracted = TDocFileHandler.download_and_extract_tdoc(tdoc_candidate, u, tdoc_dir, timeout=10)
                        except Exception as dl_err:
                            logger.warning(f"[SessionPlanWindow] TDoc download attempt failed from {u}: {dl_err}")
                            extracted = []
                        if extracted:
                            break

                    if not extracted and tdoc_dir.exists():
                        extracted = list(tdoc_dir.glob("*.docx")) + list(tdoc_dir.glob("*.docm")) + list(tdoc_dir.glob("*.doc"))

                    if not extracted:
                        QMessageBox.warning(self, "Download Failed", f"Could not acquire document for {payload}.")
                        return

                    snap = self._import_local_session_plan_path(extracted[0], tdoc_hint=tdoc_candidate)

            self.db.save_snapshot(snap)
            self._load_data()
            QMessageBox.information(self, "Import Successful", f"Successfully loaded session plan:\n{snap.snapshot_label}")

        except Exception as e:
            logger.error(f"[SessionPlanWindow] Import failed: {e}", exc_info=True)
            QMessageBox.critical(self, "Import Error", f"Failed to parse session plan:\n{e}")

    def _import_local_session_plan_path(self, source_path: Path, tdoc_hint: str = "SessionPlan") -> SessionPlanSnapshot:
        """Stage and parse a local session plan document or zip archive into a snapshot."""
        p = Path(source_path)
        ext = p.suffix.lower()
        if ext == ".zip":
            import zipfile
            from core.utils.paths import to_extended_path

            extract_dir = self.meeting_dir / "Agenda" / "SessionPlans" / "Extracted" / p.stem
            os.makedirs(to_extended_path(extract_dir), exist_ok=True)
            with zipfile.ZipFile(to_extended_path(p), "r") as zf:
                zf.extractall(to_extended_path(extract_dir))
            doc_candidates = [
                c
                for pat in ("*.docx", "*.docm", "*.doc", "*.htm*", "*.txt", "*.csv")
                for c in extract_dir.glob(pat)
                if not c.name.startswith("~$")
            ]
            if not doc_candidates:
                raise FileNotFoundError(f"No supported session plan file found inside archive {p.name}.")
            p = doc_candidates[0]
            ext = p.suffix.lower()

        staged = SessionPlanParser.stage_session_plan_file(p, self.meeting_dir, tdoc_hint)
        if ext in (".docx", ".docm"):
            return SessionPlanParser.parse_docx(staged, meeting_number=self.meeting_number)
        return SessionPlanParser.parse_text_or_markdown(
            staged.read_text(encoding="utf-8", errors="replace"),
            meeting_number=self.meeting_number,
            cached_file_path=str(staged),
        )

    def _download_session_plan_url(self, raw_url: str, timeout: int = 15) -> Path:
        """Download a direct HTTP/HTTPS session plan URL into the meeting cache via NetworkSession."""
        from urllib.parse import unquote, urlsplit
        from core.network.session import NetworkSession
        from core.utils.paths import to_extended_path

        parsed = urlsplit(raw_url)
        raw_name = unquote(Path(parsed.path).name) if parsed.path else ""

        session = NetworkSession.get_instance()
        NetworkSession.apply_humanness(session, url=raw_url)
        resp = session.get(raw_url, stream=True, timeout=timeout)
        resp.raise_for_status()

        cd = resp.headers.get("Content-Disposition", "")
        if cd:
            m_star = re.search(r"filename\*\s*=\s*(?:UTF-8''|[^']*')?([^;]+)", cd, re.IGNORECASE)
            m_plain = re.search(r'filename\s*=\s*"([^"]+)"|filename\s*=\s*([^;\s]+)', cd, re.IGNORECASE)
            cd_name = ""
            if m_star:
                cd_name = unquote(m_star.group(1).strip().strip('"'))
            elif m_plain:
                cd_name = (m_plain.group(1) or m_plain.group(2) or "").strip()
            if cd_name and (not raw_name or "." not in raw_name):
                raw_name = cd_name

        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", Path(raw_name).name if raw_name else "").strip(" .")
        if not safe_name:
            safe_name = "session_plan_download.docx"

        dl_dir = self.meeting_dir / "Agenda" / "SessionPlans" / "Downloaded"
        os.makedirs(to_extended_path(dl_dir), exist_ok=True)
        dest_path = dl_dir / safe_name
        tmp_path = dl_dir / f".{safe_name}.part"

        try:
            with open(to_extended_path(tmp_path), "wb") as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        fh.write(chunk)
            os.replace(to_extended_path(tmp_path), to_extended_path(dest_path))
        except Exception:
            try:
                ext_tmp = to_extended_path(tmp_path)
                if os.path.exists(ext_tmp):
                    os.remove(ext_tmp)
            except OSError:
                pass
            raise

        return dest_path

    # --- Diffing & Exporting ---

    def _open_diff_dialog(self):
        snaps = self.db.list_snapshots(self.meeting_number)
        if len(snaps) < 2:
            QMessageBox.information(
                self,
                "Comparison Requires 2 Snapshots",
                "You need at least two session plan snapshots to run a comparison.\nImport an earlier or revised session plan first.",
            )
            return

        snap_a = self.db.get_snapshot_by_id(snaps[1]["id"])
        snap_b = self.db.get_snapshot_by_id(snaps[0]["id"])
        diff_res = SessionPlanDiffEngine.compare_snapshots(snap_a, snap_b)
        dlg = SessionPlanDiffDialog(diff_res, self)
        dlg.exec_()
