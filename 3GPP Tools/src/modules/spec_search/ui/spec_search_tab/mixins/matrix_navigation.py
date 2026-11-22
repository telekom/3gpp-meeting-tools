"""Evolution Matrix context menu, deep-link jump, focus_spec, and shutdown mixin."""

from __future__ import annotations

import re
from typing import Optional

from PyQt5.QtCore import QThread, Qt
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox, QTableView

from core.ui.toast import notify_info
from modules.spec_search.core.spec_clause_exporter import copy_clause_to_clipboard
from modules.spec_search.core.spec_search_db import parse_version_tuple
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel


class SpecSearchTabMatrixNavMixin:
    """Handles matrix navigation, cell double-click/context menus, deep-links, and shutdown."""

    def _handle_quick_diff(self, spec_num: str, base_ver: str, target_ver: str, clause_num: str, clause_title: str):
        """Opens the clause comparison dialog pre-configured with base and target releases."""
        self._open_clause_diff_dialog(
            spec_num=spec_num,
            version=target_ver,
            clause_num=clause_num,
            clause_title=clause_title,
            base_version=base_ver,
        )

    def _handle_jump_to_version(self, spec_number: str, target_version: str, clause_number: str):
        """Jumps the inspector and matrix selection directly to the target version for this clause."""
        clean_target = target_version.lstrip("v").strip()
        for i in range(self.spec_results_tabs.count()):
            tab_title = self.spec_results_tabs.tabText(i)
            if f"TS {spec_number}" in tab_title:
                self.spec_results_tabs.setCurrentIndex(i)
                table = self.spec_results_tabs.widget(i)
                if isinstance(table, QTableView):
                    model = table.model()
                    if isinstance(model, SpecEvolutionMatrixModel):
                        for r in range(model.rowCount()):
                            if str(model.data(model.index(r, 0), Qt.DisplayRole)).strip() == clause_number:
                                target_col = 0
                                for c_idx, v_col in enumerate(model._version_cols):
                                    if clean_target in v_col:
                                        target_col = c_idx + 2
                                        break
                                target_idx = model.index(r, target_col)
                                table.setCurrentIndex(target_idx)
                                self._on_tab_table_clicked(target_idx, spec_number, table)
                                break
                break

    def _on_table_double_clicked(self, index, spec_num: str, table: QTableView):
        """Double-clicking a matrix cell directly opens the specification in Word."""
        model = table.model()
        if not model:
            return
        row, col = index.row(), index.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])
        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")
        self._handle_open_spec_document(spec_num, version, clause_num, "word")

    def _show_matrix_context_menu(self, pos, table: QTableView, spec_num: str):
        """Right-click context menu on Evolution Matrix cells."""
        idx = table.indexAt(pos)
        if not idx.isValid():
            return
        model = table.model()
        if not model:
            return

        row, col = idx.row(), idx.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        clause_title = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip()

        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])

        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")

        prev_ver_header = model.get_preceding_version(ver_header)
        prev_ver = None
        if prev_ver_header:
            m = re.search(r"v([0-9\.]+)", prev_ver_header)
            prev_ver = m.group(1) if m else prev_ver_header.lstrip("v")

        menu = QMenu(self)

        act_word = menu.addAction(f"📝 Open TS {spec_num} v{version} (Word)")
        act_word.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "word"))

        act_pdf = menu.addAction(f"📕 Open TS {spec_num} v{version} (PDF)")
        act_pdf.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "pdf"))

        act_folder = menu.addAction("📂 Open Local Folder in Explorer")
        act_folder.triggered.connect(lambda: self._handle_open_spec_document(spec_num, version, clause_num, "folder"))

        menu.addSeparator()

        act_specs = menu.addAction(f"🔍 Show TS {spec_num} v{version} in Specifications Tab")
        act_specs.triggered.connect(lambda: self._handle_show_in_specs_tab(spec_num, version))

        act_changes = menu.addAction(f"📜 View Changes && TDocs for TS {spec_num} v{version}...")
        act_changes.triggered.connect(lambda: self._open_version_history_dialog(spec_num, version))

        menu.addSeparator()

        if prev_ver:
            act_compare = menu.addAction(f"🔍 Compare Clause {clause_num} (Diff && AI Analysis vs v{prev_ver})...")
            act_compare.triggered.connect(lambda: self._open_clause_diff_dialog(spec_num, version, clause_num, clause_title, base_version=prev_ver))
        else:
            act_compare = menu.addAction(f"🔍 Compare Clause {clause_num} (Diff && AI Analysis)...")
            act_compare.triggered.connect(lambda: self._open_clause_diff_dialog(spec_num, version, clause_num, clause_title))

        menu.addSeparator()
        clause_rec = self.db.get_clause_content_by_spec_ver(spec_num, version, clause_num)
        clause_text = clause_rec.get("content", "") if clause_rec else ""
        rel_date = clause_rec.get("release_date", "") if clause_rec else ""
        actual_title = clause_rec.get("clause_title", clause_title) if clause_rec else clause_title

        menu.addSeparator()
        act_ext_word = menu.addAction(f"✂️ Extract Clause {clause_num} in Word (.docx)")
        act_ext_word.triggered.connect(
            lambda: self._handle_extract_clause_word(
                spec_num, version, clause_num, actual_title, clause_text, save_as=False
            )
        )
        act_save_word = menu.addAction(f"💾 Save Clause {clause_num} as Word (.docx)...")
        act_save_word.triggered.connect(
            lambda: self._handle_extract_clause_word(
                spec_num, version, clause_num, actual_title, clause_text, save_as=True
            )
        )
        act_export_md = menu.addAction(f"📘 Export Clause {clause_num} to Markdown (.md)...")
        act_export_md.triggered.connect(
            lambda: self._handle_export_clause_markdown(
                spec_num, version, clause_num, actual_title, clause_text
            )
        )

        menu.addSeparator()
        copy_menu = menu.addMenu("📋 Copy Clause")
        act_cp_email = copy_menu.addAction("📋 Copy for Email / Outlook (Rich Text)")
        act_cp_email.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="email"
            )
        )
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown (Quote Block)")
        act_cp_md.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="markdown"
            )
        )
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="text"
            )
        )
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(
            lambda: copy_clause_to_clipboard(
                spec_num, version, clause_num, actual_title, clause_text, rel_date, mode="citation"
            )
        )
        act_copy = menu.addAction("📋 Copy Clause Citation")
        act_copy.triggered.connect(lambda: QApplication.clipboard().setText(f"3GPP TS {spec_num} v{version}, Clause {clause_num} ({clause_title})"))

        menu.exec_(table.viewport().mapToGlobal(pos))

    def focus_spec(self, spec_number: str, version: Optional[str] = None):
        """
        External API to focus on a specific specification and version from other tabs.
        Ensures the spec is selected in the version tree and search bar is focused.
        """
        clean_spec = spec_number.strip()
        clean_ver = str(version).lstrip("vV").strip() if version else None

        # Check if spec is in version tree
        all_specs = [
            self.version_tree.topLevelItem(i).data(0, Qt.UserRole).get("spec_number")
            for i in range(1, self.version_tree.topLevelItemCount())
            if self.version_tree.topLevelItem(i).data(0, Qt.UserRole)
        ]

        if clean_spec in all_specs:
            target_spec_item = None
            target_child_item = None
            for i in range(1, self.version_tree.topLevelItemCount()):
                item = self.version_tree.topLevelItem(i)
                data = item.data(0, Qt.UserRole) or {}
                if data.get("spec_number") == clean_spec:
                    target_spec_item = item
                    item.setExpanded(True)
                    if clean_ver:
                        for c_idx in range(item.childCount()):
                            child = item.child(c_idx)
                            c_data = child.data(0, Qt.UserRole) or {}
                            if str(c_data.get("version", "")).lstrip("vV").strip() == clean_ver:
                                target_child_item = child
                                child.setCheckState(0, Qt.Checked)
                                break
                    if not target_child_item and not clean_ver:
                        item.setCheckState(0, Qt.Checked)
                    break

            if target_child_item:
                self.version_tree.setCurrentItem(target_child_item)
                self.version_tree.scrollToItem(target_child_item)
            elif target_spec_item:
                self.version_tree.setCurrentItem(target_spec_item)
                self.version_tree.scrollToItem(target_spec_item)
                if clean_ver:
                    reply = QMessageBox.question(
                        self,
                        "Version Not Indexed",
                        f"TS {clean_spec} is indexed in Spec Search, but version v{clean_ver} is not yet added.\n\n"
                        f"Would you like to index TS {clean_spec} v{clean_ver} now?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes,
                    )
                    if reply == QMessageBox.Yes and self.specs_db:
                        files = self.specs_db.get_spec_files(clean_spec)
                        matching_file = next(
                            (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                            None
                        )
                        if matching_file:
                            task = {
                                "spec_number": clean_spec,
                                "version": clean_ver,
                                "filename": matching_file.get("filename", ""),
                                "file_url": matching_file.get("url", ""),
                                "release_date": matching_file.get("upload_date", ""),
                            }
                            self._start_batch_ingestion([task])
                            return

            self.selected_version_ids = self.version_tree.get_selected_version_ids()
        else:
            prompt_msg = (
                f"TS {clean_spec} v{clean_ver} is not yet indexed in the Spec Search database.\n\nWould you like to index it now?"
                if clean_ver
                else f"TS {clean_spec} is not yet indexed in the Spec Search database.\n\nWould you like to import it from the Specifications database now?"
            )
            reply = QMessageBox.question(
                self,
                "Spec Not Indexed",
                prompt_msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                if clean_ver and self.specs_db:
                    files = self.specs_db.get_spec_files(clean_spec)
                    matching_file = next(
                        (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                        None
                    )
                    if matching_file:
                        task = {
                            "spec_number": clean_spec,
                            "version": clean_ver,
                            "filename": matching_file.get("filename", ""),
                            "file_url": matching_file.get("url", ""),
                            "release_date": matching_file.get("upload_date", ""),
                        }
                        self._start_batch_ingestion([task])
                        return
                self._on_fetch_clicked()
                return

        self.search_input.setFocus()
        self.search_input.selectAll()
        if self.search_input.text().strip():
            self._execute_search()

    def _handle_clause_reference_clicked(self, spec_number: str, version: str, target_clause: str):
        """Navigates the Clause Inspector directly to a clicked same-spec or cross-spec clause."""
        clean_spec = re.sub(r"^(?:TS|TR)\s*", "", str(spec_number or "").strip(), flags=re.IGNORECASE)
        if not clean_spec or not target_clause:
            return

        versions = self.db.get_versions_for_spec(clean_spec)
        if not versions:
            notify_info(
                self,
                "Specification Not Indexed",
                f"TS {clean_spec} is not indexed in Spec Search yet.",
            )
            return

        target_ver = str(version or "").lstrip("vV").strip()
        avail_vers = [str(v.get("version", "")) for v in versions if v.get("version")]

        if not target_ver or target_ver not in avail_vers:
            checked_for_spec = [
                str(item.get("version", "")).lstrip("vV").strip()
                for item in self.version_tree.get_checked_versions_info()
                if str(item.get("spec_number", "")).strip() == clean_spec
            ]
            checked_for_spec = [v for v in checked_for_spec if v in avail_vers]
            if checked_for_spec:
                checked_for_spec.sort(key=parse_version_tuple, reverse=True)
                target_ver = checked_for_spec[0]
            else:
                target_ver = avail_vers[0]

        candidates = [target_clause.strip()]
        if target_clause.lower().startswith("annex "):
            candidates.append(target_clause[6:].strip())

        clause_data = None
        for cand in candidates:
            clause_data = self.db.get_clause_content_by_spec_ver(clean_spec, target_ver, cand)
            if clause_data:
                break

        if not clause_data:
            notify_info(
                self,
                "Clause Not Found",
                f"Clause {target_clause} was not found in TS {clean_spec} v{target_ver}.",
            )
            return

        prev_ver = None
        if target_ver in avail_vers:
            idx = avail_vers.index(target_ver)
            if idx + 1 < len(avail_vers):
                prev_ver = avail_vers[idx + 1]

        spec_title = self.get_spec_short_name(clean_spec) or self.get_spec_title(clean_spec)
        self.inspector.display_clause(
            clause_number=clause_data.get("clause_number", target_clause),
            clause_title=clause_data.get("clause_title", ""),
            spec_number=clause_data.get("spec_number", clean_spec),
            version=clause_data.get("version", target_ver),
            release_date=clause_data.get("release_date"),
            content=clause_data.get("content", ""),
            search_query="",
            first_added_ver=None,
            prev_ver=prev_ver,
            spec_title=spec_title,
        )

    def shutdown(self, wait_ms: int = 1500):
        """Cooperatively stop timers and wait for active background workers during app shutdown."""
        self._search_timer.stop()
        self._save_config_timer.stop()

        workers = list(self._active_query_workers.values())
        workers.extend(w for w in getattr(self, "_active_doc_workers", []) if isinstance(w, QThread))
        workers.extend(w for w in getattr(self, "_active_word_workers", []) if isinstance(w, QThread))
        import_thread = getattr(self, "_import_thread", None) or getattr(self, "thread", None)
        if isinstance(import_thread, QThread) and import_thread.isRunning():
            workers.append(import_thread)
        if hasattr(self, "_wipe_worker") and self._wipe_worker and self._wipe_worker.isRunning():
            workers.append(self._wipe_worker)

        for worker in workers:
            if worker.isRunning():
                worker.requestInterruption()

        for worker in workers:
            if worker.isRunning():
                worker.wait(wait_ms)
