"""Source management, background heading scan, and tree filter/selection mixin."""

import sys
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QTreeWidgetItem

from modules.word_tools.core.word_excerpt_extractor import WordHeadingScanThread


def _get_facade_attr(name: str, default):
    facade = sys.modules.get("modules.word_tools.ui.cards.excerpt_card")
    if facade is not None and hasattr(facade, name):
        return getattr(facade, name)
    return default


class ExcerptCardScanningTreeMixin:
    """Manages source files, background heading scan lifecycle, and tree filter/selection state."""

    def _scan_excerpt_files(self, files):
        if not files:
            return

        existing = {
            str(Path(path).resolve()).casefold()
            for path in self._excerpt_input_paths
        }
        added = False
        for file_path in files:
            path = str(Path(file_path).resolve())
            key = path.casefold()
            if key in existing:
                continue
            self._excerpt_input_paths.append(path)
            existing.add(key)
            added = True

        if not added:
            return

        self._refresh_excerpt_source_list()
        self._rescan_excerpt_sources()

    def _refresh_excerpt_source_list(self):
        self.excerpt_source_list.clear()
        for path in self._excerpt_input_paths:
            self.excerpt_source_list.addItem(Path(path).name)
        has_sources = bool(self._excerpt_input_paths)
        self.excerpt_clear_sources_btn.setEnabled(has_sources)
        self.excerpt_remove_btn.setEnabled(False)

    def _invalidate_excerpt_scan(self):
        self._excerpt_paths = []
        self._excerpt_headings = []
        self.excerpt_tree.clear()
        self.excerpt_extract_btn.setEnabled(False)
        self._update_preview_filename()

    def _disconnect_running_scan(self):
        self._excerpt_scan_gen += 1
        if self._excerpt_scan_thread is not None:
            old_thread = self._excerpt_scan_thread
            for sig_name in ("scanned", "progress", "failed"):
                try:
                    getattr(old_thread, sig_name).disconnect()
                except Exception:
                    pass
            if old_thread.isRunning():
                old_thread.requestInterruption()
                self._active_scan_threads.append(old_thread)
                old_thread.finished.connect(
                    lambda t=old_thread: self._active_scan_threads.remove(t)
                    if t in self._active_scan_threads
                    else None
                )
            self._excerpt_scan_thread = None

    def _rescan_excerpt_sources(self):
        self._disconnect_running_scan()
        self._invalidate_excerpt_scan()
        if not self._excerpt_input_paths:
            self.excerpt_drop.set_state("ready", "📥 Drop one or more .docx files here to scan their headings")
            self.excerpt_progress.setVisible(False)
            self.excerpt_progress_label.setVisible(False)
            return
        self.excerpt_drop.set_state("ready", f"Scanning {len(self._excerpt_input_paths)} document(s)...")
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setText(f"Preparing to parse {len(self._excerpt_input_paths)} document(s)...")
        self.excerpt_progress_label.setVisible(True)
        mode = self.excerpt_order_combo.currentData()
        gen = self._excerpt_scan_gen
        scan_thread_cls = _get_facade_attr("WordHeadingScanThread", WordHeadingScanThread)
        thread = scan_thread_cls(list(self._excerpt_input_paths), mode)
        self._excerpt_scan_thread = thread
        thread.scanned.connect(
            lambda ordered_paths, headings, g=gen: self._on_excerpt_scanned(ordered_paths, headings)
            if g == self._excerpt_scan_gen
            else None
        )
        thread.progress.connect(
            lambda percent, message, g=gen: self._on_excerpt_scan_progress(percent, message)
            if g == self._excerpt_scan_gen
            else None
        )
        thread.failed.connect(
            lambda message, g=gen: self._on_excerpt_scan_failed(message)
            if g == self._excerpt_scan_gen
            else None
        )
        thread.start()

    def _remove_excerpt_sources(self):
        selected_rows = sorted(
            {self.excerpt_source_list.row(item) for item in self.excerpt_source_list.selectedItems()},
            reverse=True,
        )
        if not selected_rows:
            return
        for row in selected_rows:
            if 0 <= row < len(self._excerpt_input_paths):
                del self._excerpt_input_paths[row]
        self._refresh_excerpt_source_list()
        self._rescan_excerpt_sources()

    def _clear_excerpt_sources(self):
        self._disconnect_running_scan()
        self._excerpt_input_paths = []
        self._refresh_excerpt_source_list()
        self._invalidate_excerpt_scan()
        self.excerpt_drop.set_state("ready", "📥 Drop one or more .docx files here to scan their headings")
        self.excerpt_progress.setVisible(False)
        self.excerpt_progress_label.setVisible(False)

    def _on_excerpt_scan_progress(self, percent, message):
        self.excerpt_progress.setValue(percent)
        self.excerpt_progress_label.setText(message)

    def _on_excerpt_scanned(self, ordered_paths, headings):
        self._excerpt_paths = list(ordered_paths)
        self._excerpt_headings = list(headings)
        self._populate_excerpt_tree()
        self.excerpt_drop.set_state(
            "ready", f"Ready: {len(headings)} headings from {len(ordered_paths)} document(s)"
        )
        self.excerpt_extract_btn.setEnabled(bool(headings))
        self.excerpt_progress.setValue(100)
        self.excerpt_progress_label.setText(f"Scan complete: {len(headings)} headings found.")

    def _on_excerpt_scan_failed(self, message):
        self.excerpt_drop.set_state("error", "Heading scan failed")
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setVisible(True)
        self.excerpt_progress_label.setText("Heading scan failed.")
        QMessageBox.critical(self, "Heading Scan Failed", message)

    def _populate_excerpt_tree(self):
        self.excerpt_tree.blockSignals(True)
        self.excerpt_tree.clear()
        by_source = {}
        for h in self._excerpt_headings:
            by_source.setdefault(h.source_path, []).append(h)

        for source_path in self._excerpt_paths:
            source_item = QTreeWidgetItem([f"📄 {Path(source_path).name}", Path(source_path).name])
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
            self.excerpt_tree.addTopLevelItem(source_item)
            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([f"[H{h.level}] {h.text}", Path(source_path).name])
                item.setData(0, Qt.UserRole, h.id)
                item.setCheckState(0, Qt.Unchecked)
                parent = None
                for level in range(h.level - 1, 0, -1):
                    if level in parents:
                        parent = parents[level]
                        break
                (parent or source_item).addChild(item)
                parents[h.level] = item
                for level in list(parents):
                    if level > h.level:
                        del parents[level]
        self.excerpt_tree.expandToDepth(1)
        self.excerpt_tree.blockSignals(False)
        self._update_preview_filename()

    def _filter_excerpt_tree(self, text):
        needle = text.strip().lower()
        matches = 0

        def visit(item):
            nonlocal matches
            is_heading = bool(item.data(0, Qt.UserRole))
            own_match = not needle or needle in item.text(0).lower()
            if needle and is_heading and own_match:
                matches += 1
            child_match = False
            for i in range(item.childCount()):
                child_match = visit(item.child(i)) or child_match
            visible = own_match or child_match
            item.setHidden(not visible)
            if needle and child_match:
                item.setExpanded(True)
            return visible

        for i in range(self.excerpt_tree.topLevelItemCount()):
            visit(self.excerpt_tree.topLevelItem(i))

        if hasattr(self, "lbl_excerpt_match_count"):
            if needle:
                self.lbl_excerpt_match_count.setText(f"{matches} match(es)")
            else:
                self.lbl_excerpt_match_count.setText("")

    def _select_all_excerpt_headings(self):
        self.excerpt_tree.blockSignals(True)

        def select(item):
            if item.data(0, Qt.UserRole) and not item.isHidden():
                item.setCheckState(0, Qt.Checked)
            for i in range(item.childCount()):
                select(item.child(i))

        for i in range(self.excerpt_tree.topLevelItemCount()):
            select(self.excerpt_tree.topLevelItem(i))
        self.excerpt_tree.blockSignals(False)
        self._update_preview_filename()

    def _select_level1_excerpt_headings(self):
        self.excerpt_tree.blockSignals(True)
        self._clear_excerpt_selection(update_preview=False)
        for i in range(self.excerpt_tree.topLevelItemCount()):
            src_item = self.excerpt_tree.topLevelItem(i)
            for c in range(src_item.childCount()):
                item = src_item.child(c)
                if item.data(0, Qt.UserRole) and not item.isHidden():
                    item.setCheckState(0, Qt.Checked)
        self.excerpt_tree.blockSignals(False)
        self._update_preview_filename()

    def _clear_excerpt_selection(self, update_preview: bool = True):
        self.excerpt_tree.blockSignals(True)

        def clear(item):
            if item.data(0, Qt.UserRole):
                item.setCheckState(0, Qt.Unchecked)
            for i in range(item.childCount()):
                clear(item.child(i))

        for i in range(self.excerpt_tree.topLevelItemCount()):
            clear(self.excerpt_tree.topLevelItem(i))
        self.excerpt_tree.blockSignals(False)
        if update_preview:
            self._update_preview_filename()

    def _on_tree_item_changed(self, item, column):
        if column == 0:
            self._update_preview_filename()
