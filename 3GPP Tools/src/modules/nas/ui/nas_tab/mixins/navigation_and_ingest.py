"""Spec opener menus, reverse-lookup navigation, ingestion dispatch, and shutdown mixin."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QPoint, Qt, QThread, QTimer
from PyQt5.QtWidgets import QAction, QApplication, QMenu

from core.ui.theme import C
from core.ui.toast import notify_info
from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.ui.nas_ingest_controller import (
    clear_version_clicked,
    fetch_from_specs_db_clicked,
    import_cross_ref_spec_requested,
    import_local_file_clicked,
    on_import_error,
    on_import_progress,
    on_import_success,
    start_batch_ingestion,
    wipe_db_clicked,
)
from modules.nas.ui.nas_matrix_controller import show_matrix_context_menu
from modules.spec_search.core.spec_doc_resolver import SpecDocResolver


class NASTabNavigationAndIngestMixin:
    """Implements Open-Spec menus, reverse lookup jump/dispatch, ingest forwarding, and shutdown."""

    def _on_msg_list_context_menu(self, pos: QPoint):
        item = self.msg_list.itemAt(pos)
        if not item:
            return

        name = item.data(Qt.UserRole)
        spec_num = str(item.data(Qt.UserRole + 2) or "").replace("TS", "").strip()
        clause = str(item.data(Qt.UserRole + 3) or "").strip()

        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)

        clause_tag = f" (Clause {clause})" if clause else ""
        if spec_num:
            act_word = QAction(f"📝 Open TS {spec_num} in Word{clause_tag}", self)
            act_word.triggered.connect(
                lambda: self.open_specification(
                    spec_number=spec_num,
                    clause_number=clause,
                    doc_format="word",
                )
            )
            menu.addAction(act_word)

            act_pdf = QAction(f"📕 Open TS {spec_num} in PDF", self)
            act_pdf.triggered.connect(
                lambda: self.open_specification(
                    spec_number=spec_num,
                    clause_number=clause,
                    doc_format="pdf",
                )
            )
            menu.addAction(act_pdf)

            act_folder = QAction("📂 Open Local Folder in Explorer", self)
            act_folder.triggered.connect(
                lambda: self.open_specification(
                    spec_number=spec_num,
                    doc_format="folder",
                )
            )
            menu.addAction(act_folder)

            menu.addSeparator()

            act_specs = QAction(f"📚 Show TS {spec_num} in Specifications Tab", self)
            act_specs.triggered.connect(
                lambda: self.navigate_to_specs_tab_requested.emit(spec_num, "")
            )
            menu.addAction(act_specs)

            act_search = QAction(f"🔎 Search TS {spec_num} in Spec Search Tab", self)
            act_search.triggered.connect(
                lambda: self.search_in_spec_search_requested.emit(spec_num, "")
            )
            menu.addAction(act_search)

            menu.addSeparator()

        act_copy = QAction(f"📋 Copy '{name}'", self)
        act_copy.triggered.connect(lambda: QApplication.clipboard().setText(str(name)))
        menu.addAction(act_copy)

        if clause:
            act_copy_cl = QAction(f"📋 Copy Clause '{clause}'", self)
            act_copy_cl.triggered.connect(lambda: QApplication.clipboard().setText(str(clause)))
            menu.addAction(act_copy_cl)

        menu.exec_(self.msg_list.viewport().mapToGlobal(pos))

    # -------------------------------------------------------------------------
    # Specification Resolution & Launching
    # -------------------------------------------------------------------------

    def _on_open_spec_btn_clicked(self):
        spec_num = getattr(self, "_current_message_spec", None)
        clause = getattr(self, "_current_message_clause", None)
        self.open_specification(spec_number=spec_num, clause_number=clause, doc_format="word")

    def _update_open_spec_button(self):
        spec_num = getattr(self, "_current_message_spec", None)
        clause = getattr(self, "_current_message_clause", None)

        if not spec_num:
            checked = self.version_tree.get_checked_versions_info()
            checked_specs = sorted(set(v.get("spec_number") for v in checked if v.get("spec_number")))
            if len(checked_specs) == 1:
                spec_num = checked_specs[0]

        if not spec_num:
            self.open_spec_btn.setEnabled(False)
            self.open_spec_btn.setText("Open Spec ▾")
            self.open_spec_btn.setToolTip("Select a message or specification to open its document")
            self.open_spec_btn.setMenu(None)
            return

        self.open_spec_btn.setEnabled(True)
        self.open_spec_btn.setText(f"Open TS {spec_num} ▾")
        tip = f"Open 3GPP TS {spec_num}"
        if clause:
            tip += f" at Clause {clause}"
        tip += " in Microsoft Word (Clause copied to clipboard)"
        self.open_spec_btn.setToolTip(tip)

        menu = self._build_open_spec_menu(spec_num, clause)
        self.open_spec_btn.setMenu(menu)

    def _build_open_spec_menu(self, spec_num: str, clause: Optional[str] = None) -> QMenu:
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)

        checked = self.version_tree.get_checked_versions_info()
        spec_vers = [v.get("version") for v in checked if v.get("spec_number") == spec_num and v.get("version")]
        if not spec_vers:
            all_v = self.version_tree.get_all_versions_for_spec(spec_num)
            spec_vers = [v.get("version") for v in all_v if v.get("version")]

        latest_ver = spec_vers[0] if spec_vers else ""
        clause_tag = f" (Clause {clause})" if clause else ""

        act_word = QAction(f"📝 Open TS {spec_num} v{latest_ver} in Word{clause_tag}", self)
        act_word.triggered.connect(
            lambda: self.open_specification(
                spec_number=spec_num,
                version=latest_ver,
                clause_number=clause,
                doc_format="word",
            )
        )
        menu.addAction(act_word)

        if len(spec_vers) > 1:
            rel_menu = menu.addMenu(f"🏷️ Open Release in Word ({len(spec_vers)} active)...")
            for v in spec_vers:
                v_clean = str(v).lstrip("vV")
                parts = v_clean.split(".")
                major_rel = f"Rel-{parts[0]} · " if parts and parts[0].isdigit() else ""
                sub_act = QAction(f"📝 {major_rel}v{v}", self)
                sub_act.triggered.connect(
                    lambda checked=False, target_v=v: self.open_specification(
                        spec_number=spec_num,
                        version=target_v,
                        clause_number=clause,
                        doc_format="word",
                    )
                )
                rel_menu.addAction(sub_act)

        act_pdf = QAction(f"📕 Open TS {spec_num} v{latest_ver} in PDF", self)
        act_pdf.triggered.connect(
            lambda: self.open_specification(
                spec_number=spec_num,
                version=latest_ver,
                clause_number=clause,
                doc_format="pdf",
            )
        )
        menu.addAction(act_pdf)

        act_folder = QAction("📂 Open Local Folder in Explorer", self)
        act_folder.triggered.connect(
            lambda: self.open_specification(
                spec_number=spec_num,
                version=latest_ver,
                doc_format="folder",
            )
        )
        menu.addAction(act_folder)

        menu.addSeparator()

        act_specs = QAction(f"📚 Show TS {spec_num} v{latest_ver} in Specifications Tab", self)
        act_specs.triggered.connect(
            lambda: self.navigate_to_specs_tab_requested.emit(spec_num, latest_ver)
        )
        menu.addAction(act_specs)

        act_search = QAction(f"🔎 Search TS {spec_num} in Spec Search Tab", self)
        act_search.triggered.connect(
            lambda: self.search_in_spec_search_requested.emit(spec_num, latest_ver)
        )
        menu.addAction(act_search)

        return menu

    def open_specification(
        self,
        spec_number: Optional[str] = None,
        version: Optional[str] = None,
        clause_number: Optional[str] = None,
        doc_format: str = "word",
    ) -> bool:
        """
        Dispatches specification document opening (Word, PDF, Explorer folder)
        for the active protocol specification with clipboard clause search assistance.
        """
        clean_spec = re.sub(
            r"^(?:TS|TR)\s*",
            "",
            str(spec_number or getattr(self, "_current_message_spec", "") or "").strip(),
            flags=re.IGNORECASE,
        )

        if not clean_spec:
            checked = self.version_tree.get_checked_versions_info()
            if checked:
                clean_spec = checked[0].get("spec_number", "")

        if not clean_spec:
            notify_info(
                self,
                "No Specification Selected",
                "Please select a message, Information Element, or specification first.",
            )
            return False

        clean_ver = str(version or "").lstrip("vV").strip()
        if not clean_ver:
            active_vers = [
                v["version"]
                for v in self.version_tree.get_checked_versions_info()
                if v.get("spec_number") == clean_spec and v.get("version")
            ]
            if active_vers:
                clean_ver = sorted(active_vers, key=parse_version_tuple, reverse=True)[0]
            else:
                all_v = self.version_tree.get_all_versions_for_spec(clean_spec)
                if all_v:
                    clean_ver = all_v[0].get("version", "")

        clause = clause_number if clause_number is not None else getattr(self, "_current_message_clause", None)

        prev_text = self.open_spec_btn.text()
        self.open_spec_btn.setText("⏳ Opening...")

        def _on_opened(doc_path: Optional[Path] = None):
            self.open_spec_btn.setText("✓ Opened!")
            QTimer.singleShot(1500, lambda: self.open_spec_btn.setText(prev_text))

        success = SpecDocResolver.open_spec(
            spec_number=clean_spec,
            version=clean_ver,
            clause_number=clause,
            doc_format=doc_format,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            parent_widget=self,
            on_opened_callback=_on_opened,
        )

        if not success:
            self.open_spec_btn.setText(prev_text)

        return success

    def _release_reverse_lookup_worker(self, req_id: int):
        """Release a completed reverse-lookup worker safely."""
        worker = self._reverse_lookup_workers.pop(req_id, None)
        if worker is not None:
            worker.deleteLater()

    def _on_reverse_lookup_finished(self, containing_msgs: List[Dict[str, Any]], req_id: int):
        if req_id == self._reverse_lookup_request_id:
            self.inspector.set_containing_messages(containing_msgs)

    def _on_matrix_context_menu(self, pos):
        show_matrix_context_menu(self, pos)

    def _jump_to_message(self, message_name: str):
        self.current_selected_message_name = message_name
        if self._browse_mode != "messages":
            self.msg_search.clear()
            self.ie_search.clear()
            self._set_browse_mode("messages", trigger_populate=True)
            return

        self._save_config()

        if self.msg_search.text().strip() or self.ie_search.text().strip():
            visible_names = [
                self.msg_list.item(i).data(Qt.UserRole)
                for i in range(self.msg_list.count())
                if not self.msg_list.item(i).isHidden()
            ]
            if message_name not in visible_names:
                self.msg_search.clear()
                self.ie_search.clear()
                self._populate_messages()
                return

        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            if item.data(Qt.UserRole) == message_name:
                item.setHidden(False)
                self.msg_list.setCurrentItem(item)
                self.msg_list.scrollToItem(item)
                self._on_message_clicked(item)
                break

    # -------------------------------------------------------------------------
    # Ingestion Actions
    # -------------------------------------------------------------------------

    def _on_fetch_from_specs_db_clicked(self):
        fetch_from_specs_db_clicked(self)

    def _on_import_local_file_clicked(self):
        import_local_file_clicked(self)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        start_batch_ingestion(self, tasks)

    def _on_import_error(self, err_msg: str):
        on_import_error(self, err_msg)

    def _on_import_progress(self, msg: str, val: int):
        on_import_progress(self, msg, val)

    def _on_import_success(self, spec_count: int, msg_count: int):
        on_import_success(self, spec_count, msg_count)

    def _on_clear_version_clicked(self):
        clear_version_clicked(self)

    def _on_wipe_db_clicked(self):
        wipe_db_clicked(self)

    def _on_import_cross_ref_spec_requested(self, spec_number: str, major_version: int):
        import_cross_ref_spec_requested(self, spec_number, major_version)

    def shutdown(self, wait_ms: int = 1500):
        """Cooperatively wait for short-lived NAS query workers during app shutdown."""
        self._msg_search_timer.stop()
        self._ie_search_timer.stop()

        workers = list(self._msg_query_workers.values()) + list(self._reverse_lookup_workers.values())
        ingest_thread = getattr(self, "_ingest_thread", None) or getattr(self, "thread", None)
        if isinstance(ingest_thread, QThread) and ingest_thread.isRunning():
            workers.append(ingest_thread)

        for worker in workers:
            if worker.isRunning():
                worker.requestInterruption()

        for worker in workers:
            if worker.isRunning():
                worker.wait(wait_ms)
