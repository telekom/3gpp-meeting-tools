# --- File: src/core/ui/main_window/mixins/panels_navigation.py ---
"""
Bottom drawer, RightSidePanel docking/promotion, tab changes, Command Palette,
cross-tab navigation, and Specs -> Spec Search indexing triggers.

Mixin of DragDropUI: relies on attributes set in DragDropUI.__init__ / _setup_ui.
"""
from pathlib import Path
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMessageBox
from core.ui.command_palette import (
    CommandPaletteDialog,
    DynamicPaletteSearcher,
    build_static_command_items,
)
from core.ui.right_side_panel import MODE_ASSISTANT, MODE_SPEC_SEARCH
from core.utils.paths import get_project_root


class PanelsNavigationMixin:
    """Bottom drawer, RightSidePanel docking/promotion, tab changes, Command Palette,"""

    def toggle_bottom_drawer(self):
        """Toggle the bottom Console & Queue drawer cleanly."""
        if hasattr(self, "bottom_drawer"):
            self.bottom_drawer.toggle()

    def _on_bottom_drawer_toggled(self, opened: bool):
        if hasattr(self, "bottom_drawer") and hasattr(self, "splitter"):
            self.bottom_drawer.apply_splitter_sizes(self.splitter)
        if hasattr(self, "status_bar") and hasattr(self, "bottom_drawer"):
            self.status_bar.update_drawer_badge(
                opened,
                self.bottom_drawer.queue_count,
                self.bottom_drawer.unread_errors,
            )

    def _on_drawer_badges_changed(self, queue_count: int, unread_errors: int):
        if hasattr(self, "status_bar") and hasattr(self, "bottom_drawer"):
            self.status_bar.update_drawer_badge(
                self.bottom_drawer.is_open,
                queue_count,
                unread_errors,
            )

    def _on_queue_updated(self, display_items: list):
        self.queue_panel.update_list(display_items)
        if hasattr(self, "bottom_drawer"):
            self.bottom_drawer.set_queue_count(len(display_items))

    def _on_tab_changed(self, index: int):
        """Triggers lazy loading when an inactive tab is selected and syncs NavRail."""
        if hasattr(self, "nav_rail"):
            self.nav_rail.set_active_index(index)
        active_widget = self.tabs.widget(index)
        if active_widget is getattr(self, "assistant_tab", None) and getattr(self, "_assistant_docked", False):
            self._set_assistant_docked(False)
            if (
                hasattr(self, "right_panel")
                and self.right_panel.is_open
                and self.right_panel.mode == MODE_ASSISTANT
            ):
                self.right_panel.close_panel()
        if hasattr(active_widget, 'ensure_loaded'):
            active_widget.ensure_loaded()

    def _set_assistant_docked(self, docked: bool):
        """Move single-instance AssistantTab body between full tab and RightSidePanel."""
        if (
            not hasattr(self, "assistant_tab")
            or not hasattr(self.assistant_tab, "set_docked_mode")
            or not hasattr(self, "right_panel")
        ):
            return
        want = bool(docked)
        if want == getattr(self, "_assistant_docked", False):
            return
        if want:
            self.assistant_tab.set_docked_mode(True)
            self.right_panel.set_slot_widget(MODE_ASSISTANT, self.assistant_tab.body_container)
            self._assistant_docked = True
        else:
            self._assistant_docked = False
            self.assistant_tab.restore_body_into_tab()

    def _sync_assistant_dock_slot(self):
        if not hasattr(self, "right_panel") or not hasattr(self, "_set_assistant_docked"):
            return
        want = bool(self.right_panel.is_open and self.right_panel.mode == MODE_ASSISTANT)
        self._set_assistant_docked(want)

    def toggle_right_panel(self, mode: str):
        """Toggle or switch the contextual RightSidePanel (Assistant / Quick Spec Search)."""
        if hasattr(self, "right_panel"):
            self.right_panel.toggle_mode(mode)

    def open_right_panel(self, mode: str):
        """Open the contextual RightSidePanel on the requested mode."""
        if hasattr(self, "right_panel"):
            self.right_panel.open_mode(mode)

    def close_right_panel(self):
        """Close the contextual RightSidePanel."""
        if hasattr(self, "right_panel"):
            self.right_panel.close_panel()

    def _on_workspace_splitter_moved(self, _pos: int, _index: int):
        if hasattr(self, "right_panel") and hasattr(self, "workspace_splitter"):
            self.right_panel.capture_width_from_splitter(self.workspace_splitter)
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def _on_right_panel_state_changed(self, opened: bool):
        self._sync_assistant_dock_slot()
        if (
            opened
            and hasattr(self, "right_panel")
            and self.right_panel.mode == MODE_SPEC_SEARCH
            and hasattr(self, "quick_search_panel")
        ):
            self.quick_search_panel.refresh_versions()
            QTimer.singleShot(0, self.quick_search_panel.focus_input)
        if hasattr(self, "right_panel") and hasattr(self, "workspace_splitter"):
            self.right_panel.apply_splitter_sizes(self.workspace_splitter)
        if hasattr(self, "status_bar") and hasattr(self, "right_panel"):
            self.status_bar.update_right_panel_badges(opened, self.right_panel.mode)
        if hasattr(self, "nav_rail"):
            self.nav_rail.apply_responsive_width(self.width(), dock_open=opened)
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def _on_right_panel_promote(self, mode: str):
        """Promote docked panel mode into its full workspace tab and close the dock."""
        if mode == MODE_ASSISTANT:
            if hasattr(self, "_set_assistant_docked"):
                self._set_assistant_docked(False)
            if hasattr(self, "assistant_tab"):
                self.tabs.setCurrentWidget(self.assistant_tab)
        elif mode == MODE_SPEC_SEARCH and hasattr(self, "spec_search_tab"):
            self.tabs.setCurrentWidget(self.spec_search_tab)
            if hasattr(self, "quick_search_panel"):
                q = self.quick_search_panel.current_query()
                if q and hasattr(self.spec_search_tab, "search_input"):
                    self.spec_search_tab.search_input.setText(q)
        if hasattr(self, "right_panel"):
            self.right_panel.close_panel()

    def _navigate_from_quick_search_to_spec_search(self, spec_number: str, version: str, clause: str):
        """Switch to full Spec Search tab and focus spec, version, clause, and query."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        if spec_number:
            self.spec_search_tab.focus_spec(spec_number, version)
        if clause and hasattr(self.spec_search_tab, "clause_filter_input"):
            self.spec_search_tab.clause_filter_input.setText(clause)
        if hasattr(self, "quick_search_panel"):
            q = self.quick_search_panel.current_query()
            if q and hasattr(self.spec_search_tab, "search_input"):
                self.spec_search_tab.search_input.setText(q)
        if hasattr(self.spec_search_tab, "_execute_search"):
            self.spec_search_tab._execute_search()

    def _navigate_from_search_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def _navigate_from_specs_to_search(self, spec_number: str, version: str):
        """Switches to Spec Search tab and focuses on the requested spec and version."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab.focus_spec(spec_number, version)

    def _navigate_from_assistant_to_spec_search(self, spec_number: str, version: str, clause: str):
        """Switches to Spec Search tab and loads the requested spec, version, and clause."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        if spec_number:
            self.spec_search_tab.focus_spec(spec_number, version)
        if clause and hasattr(self.spec_search_tab, "clause_filter_input"):
            self.spec_search_tab.clause_filter_input.setText(clause)

    def _navigate_from_assistant_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def open_command_palette(self):
        """Launch the modal Ctrl+K quick switcher & command palette."""
        if not hasattr(self, "_palette_searcher") or self._palette_searcher is None:
            root_db = get_project_root() / "db"
            self._palette_searcher = DynamicPaletteSearcher(
                db_path=root_db / "3gpp_data.db",
                protocol_db_path=root_db / "3gpp_protocol_data.db",
                on_open_spec=lambda s: self._navigate_from_search_to_specs(s, ""),
                on_open_meeting=self._open_meeting_from_palette,
                on_open_work_item=self._open_work_item_from_palette,
                on_open_protocol=self._navigate_from_assistant_to_protocols,
            )
        dlg = CommandPaletteDialog(
            build_static_command_items(self),
            dynamic_searcher=self._palette_searcher.search,
            parent=self,
        )
        dlg.exec_()

    def _open_meeting_from_palette(self, query_or_tdoc: str):
        """Switch to Meetings tab and filter-safely focus the requested meeting/TDoc."""
        if hasattr(self, "meetings_tab"):
            self.tabs.setCurrentWidget(self.meetings_tab)
            if hasattr(self.meetings_tab, "focus_meeting"):
                self.meetings_tab.focus_meeting(query_or_tdoc)

    def _open_work_item_from_palette(self, query: str):
        """Switch to Work Items tab and filter-safely focus the requested WI."""
        if hasattr(self, "work_items_tab"):
            self.tabs.setCurrentWidget(self.work_items_tab)
            if hasattr(self.work_items_tab, "focus_work_item"):
                self.work_items_tab.focus_work_item(query)

    def _navigate_from_assistant_to_protocols(self, protocol: str, message: str):
        """Switches to Protocols tab and selects the requested message."""
        self.tabs.setCurrentWidget(self.nas_tab)
        if hasattr(self.nas_tab, "ensure_loaded"):
            self.nas_tab.ensure_loaded()
        if message and hasattr(self.nas_tab, "_jump_to_message"):
            self.nas_tab._jump_to_message(message)

    def _smart_index_spec_from_specs_tab(self, spec_number: str):
        """Switches to Spec Search tab and opens the targeted Smart Import dialog for this spec."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab._open_smart_import_dialog(initial_spec=spec_number)

    def _index_spec_from_specs_tab(self, spec_number: str, version: str):
        """Directly indexes a single specification version from Specifications tab into Spec Search."""
        self._batch_index_specs_from_specs_tab([(spec_number, version)])

    def _batch_index_specs_from_specs_tab(self, spec_items: list):
        """Directly indexes one or more specification versions from Specifications tab into Spec Search."""
        from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
        from modules.spec_search.core.spec_search_threads import is_change_mark_file

        if not spec_items:
            return

        tasks = []
        for spec_number, version in spec_items:
            clean_spec = str(spec_number).strip()
            clean_ver = str(version).lstrip("vV").strip()

            # 1. Exact lookup via get_spec_files
            files = self.specs_tab.db.get_spec_files(clean_spec)
            matching_file = None
            if clean_ver:
                matching_file = next(
                    (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                    None
                )
            elif files:
                matching_file = files[0]
                clean_ver = str(matching_file.get("version", "")).lstrip("vV").strip()

            filename = ""
            file_url = ""
            release_date = ""

            if matching_file:
                filename = matching_file.get("filename", "")
                file_url = matching_file.get("url", "")
                release_date = matching_file.get("upload_date", "")
            else:
                rows = self.specs_tab.db.search_files(spec_number=clean_spec, release_version=clean_ver)
                if rows:
                    filename = rows[0][4]
                    file_url = rows[0][6]
                    release_date = rows[0][7]
                    if not clean_ver and len(rows[0]) > 5:
                        clean_ver = str(rows[0][5]).lstrip("vV").strip()

            if not filename and not file_url:
                continue

            task = {
                "spec_number": clean_spec,
                "version": clean_ver,
                "filename": filename,
                "file_url": file_url,
                "release_date": release_date,
            }

            # Check if Word document is already extracted locally in download_dir
            spec_dir = Path(self.specs_tab.download_dir) / clean_spec
            stem = filename.replace(".zip", "")
            if spec_dir.exists() and stem:
                local_docs = [
                    f for f in spec_dir.iterdir()
                    if f.is_file()
                    and f.stem.lower().startswith(stem.lower())
                    and f.suffix.lower() in [".docx", ".doc"]
                    and not f.name.startswith("~$")
                    and not is_change_mark_file(f.name)
                ]
                if local_docs:
                    local_docs.sort(key=ProtocolDocxDispatcher._extract_part_index)
                    task["local_docx_paths"] = local_docs

            tasks.append(task)

        if not tasks:
            QMessageBox.warning(
                self,
                "Spec Not Found",
                "Could not find specification metadata or archive records for the requested version(s)."
            )
            return

        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab._start_batch_ingestion(tasks)
