# --- File: src/core/ui/main_window/mixins/puml_conversion.py ---
"""
PlantUML editor cache, templates/snippets, Visio extraction, conversion queue,
Word <-> Markdown handlers, conversion toasts, and output folder/file openers.

Mixin of DragDropUI: relies on attributes set in DragDropUI.__init__ / _setup_ui.
"""
import datetime
import logging
import os
import subprocess
import webbrowser
from pathlib import Path
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QMessageBox,
)
from core.ui.toast import ToastAction
from core.utils.paths import get_project_root
from modules.puml2visio.core.visio_converter import VisioReaderThread
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.utils.utils import encode_plantuml


class PumlConversionMixin:
    """PlantUML editor cache, templates/snippets, Visio extraction, conversion queue,"""

    # --- AUTO-SAVE LOGIC ---
    def _load_cache(self):
        if self.cache_file.exists():
            try:
                text = self.cache_file.read_text(encoding="utf-8")
                if text.strip():
                    self.code_tab.text_input.setPlainText(text)
                    self.log_message("♻️ Restored main window previous session.", logging.INFO)
            except Exception as e:
                self.log_message(f"⚠️ Could not load previous session: {e}", logging.WARNING)

    def save_cache(self):
        try:
            self.cache_file.write_text(self.code_tab.get_text(), encoding="utf-8")
        except Exception:
            pass

    def _on_send_code_to_editor(self, code: str):
        self.tabs.setCurrentWidget(self.code_tab)
        self._set_editor_text(code)
        self.log_message("📝 Loaded extracted PlantUML code into Editor tab.", logging.INFO)

    def _on_insert_snippet_from_visio(self, snippet):
        self.tabs.setCurrentWidget(self.code_tab)
        if hasattr(self.code_tab, "text_input"):
            self.code_tab.text_input.insert_snippet(snippet)
        label = getattr(snippet, 'label', 'snippet')
        self.log_message(f"🧩 Inserted '{label}' into PlantUML editor.", logging.INFO)

    def _set_editor_text(self, text):
        if hasattr(self.code_tab, "set_text"):
            self.code_tab.set_text(text)
        else:
            cursor = self.code_tab.text_input.textCursor()
            cursor.beginEditBlock()
            cursor.select(QTextCursor.Document)
            cursor.insertText(text)
            cursor.endEditBlock()
            cursor.setPosition(0)
            self.code_tab.text_input.setTextCursor(cursor)
        self.live_preview.update_now()

    def insert_template(self, selected_template):
        tpl = PLANTUML_TYPES[selected_template]["template"]
        curr_text = self.code_tab.get_text()
        if curr_text and curr_text.strip() != tpl.strip():
            reply = QMessageBox.question(
                self,
                "Replace Diagram?",
                f"The editor already contains PlantUML code.\n\n"
                f"Do you want to replace it completely with the '{selected_template}' boilerplate?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self._set_editor_text(tpl)
        self.log_message(f"📝 Inserted boilerplate for '{selected_template}' diagram.", logging.INFO)

    def clear_editor(self):
        self._set_editor_text("")
        if hasattr(self.code_tab, "_current_file_path"):
            self.code_tab._current_file_path = None
        self.log_message("🗑️ Editor cleared.", logging.INFO)

    def copy_editor_code(self):
        code_text = self.code_tab.get_text()
        if code_text:
            QApplication.clipboard().setText(code_text)
            self.log_message("📄 Source code copied to clipboard.", logging.INFO)
        else:
            self.log_message("⚠️ Editor is empty. Nothing to copy.", logging.WARNING)

    def open_export_folder(self):
        export_dir = get_project_root().parent / "output"
        export_dir.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(export_dir)
            self.log_message(f"📂 Opened export directory: {export_dir}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open directory: {e}", logging.ERROR)

    def open_template_docs(self, selected_template):
        url = PLANTUML_TYPES[selected_template]["url"]
        try:
            webbrowser.open(url)
            self.log_message(f"🌐 Opened documentation for '{selected_template}'.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open URL: {e}", logging.ERROR)

    def toggle_live_view(self, checked):
        open_browser = self.code_tab.live_view_btn.isChecked()
        self.live_preview.toggle(checked, open_browser=open_browser)


    def check_for_jar_updates(self):
        self.log_message("\n🔄 Initiating manual update check...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Checking online for updates...")
        self.status_bar.showMessage("⏳ Checking for updates...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=True)

    def _on_editor_file_dropped(self, file_path: str):
        """Dispatches files dropped onto the PlantUML editor to the appropriate loader."""
        path = Path(file_path)
        ext = path.suffix.lower()
        if ext in (".vsdx", ".vsdm"):
            self.extract_code_from_visio(file_path)
        else:
            self.load_puml_file(file_path)

    def load_puml_file(self, file_path):
        """Loads a PlantUML file into the editor, updates live preview, and logs details."""
        path = Path(file_path)
        if not path.is_file():
            self.log_message(f"❌ File not found: {file_path}", logging.ERROR)
            return

        try:
            success = self.code_tab.load_file(path)
            if success:
                self.live_preview.update_now()
                from core.ui.drop_zone import format_file_size
                size_str = format_file_size(path.stat().st_size)
                self.log_message(f"📂 Loaded PlantUML diagram from: {path.name} ({size_str})", logging.INFO)
                self.status_bar.showMessage(f"Loaded {path.name}", 4000)
            else:
                self.log_message(f"❌ Could not load PlantUML file: {path.name}", logging.ERROR)
        except Exception as e:
            self.log_message(f"❌ Failed to load PlantUML file '{path.name}': {e}", logging.ERROR)
            QMessageBox.warning(self, "Load Error", f"Failed to load PlantUML file:\n{e}")

    def extract_code_from_visio(self, file_path):
        self.code_tab.text_input.clear()
        self.code_tab.text_input.setPlaceholderText(
            f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait..."
        )
        self.code_tab.text_input.setEnabled(False)
        self.log_message(f"📂 Reading embedded source from: {Path(file_path).name}")
        self.reader_thread = VisioReaderThread(file_path)
        self.reader_thread.text_extracted.connect(self.on_visio_code_read)
        self.reader_thread.error_occurred.connect(self.on_visio_code_error)
        self.reader_thread.start()

    def on_visio_code_read(self, source_code):
        self.code_tab.text_input.setEnabled(True)
        self._set_editor_text(source_code)
        self.log_message("✅ Successfully extracted PlantUML source from Visio file.")

    def on_visio_code_error(self, error_msg):
        self.code_tab.text_input.setEnabled(True)
        self.code_tab.text_input.setPlaceholderText(
            "Paste PlantUML code OR drag & drop a .puml or .vsdx file here to load..."
        )
        self.log_message(f"❌ {error_msg}")

    def show_in_planttext(self):
        raw_text = self.code_tab.get_text()
        if not raw_text: return
        try:
            url = f"https://www.planttext.com/?text={encode_plantuml(raw_text)}"
            webbrowser.open(url)
            self.log_message("🌐 Opened code in planttext.com")
        except Exception as e:
            self.log_message(f"❌ Failed to open: {e}")

    def copy_out_path(self):
        if self.last_out_path:
            QApplication.clipboard().setText(self.last_out_path)
            self.log_message(f"📋 Copied to clipboard: {self.last_out_path}")

    def _save_and_queue_pasted_text(self, target_format):
        raw_text = self.code_tab.get_text()
        if not raw_text: return

        params = {}
        if target_format == "pptx_append":
            from modules.puml2visio.core.powerpoint_converter import get_open_powerpoint_presentations
            from modules.puml2visio.ui.select_presentation_dialog import SelectPresentationDialog

            open_presentations = get_open_powerpoint_presentations()
            if not open_presentations:
                reply = QMessageBox.question(
                    self,
                    "No Open PowerPoint Presentation",
                    "No open Microsoft PowerPoint presentation could be detected.\n\n"
                    "Would you like to export to a new PowerPoint presentation instead?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    target_format = "pptx"
                else:
                    return
            elif len(open_presentations) == 1:
                target_pres = open_presentations[0]
                params = {
                    "target_presentation": target_pres["name"],
                    "insert_at": "end"
                }
            else:
                dialog = SelectPresentationDialog(open_presentations, parent=self)
                if dialog.exec_() != QDialog.Accepted:
                    return
                if getattr(dialog, "is_new_presentation_requested", lambda: False)():
                    target_format = "pptx"
                else:
                    chosen = dialog.get_selected_presentation()
                    if not chosen:
                        return
                    params = {
                        "target_presentation": chosen["name"],
                        "insert_at": chosen.get("insert_at", "end")
                    }

        out_dir = get_project_root().parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = out_dir / f"{base_name}.puml"

        counter = 1
        while (
            puml_path.exists()
            or puml_path.with_suffix(".vsdx").exists()
            or puml_path.with_suffix(".pptx").exists()
            or puml_path.with_suffix(".svg").exists()
        ):
            puml_path = out_dir / f"{base_name}_{counter}.puml"
            counter += 1

        with open(puml_path, "w", encoding="utf-8") as f:
            f.write(raw_text)

        self.queue_manager.add_item(puml_path, target_format, params=params)

    def _on_word_convert_markdown_requested(self, doc: str, *args, **kwargs) -> None:
        """Enqueues Word to Markdown conversion supporting dict or positional arguments."""
        if args and isinstance(args[0], dict):
            opts = args[0]
        elif kwargs:
            opts = kwargs
        elif len(args) >= 1 and isinstance(args[0], str):
            rev = args[0]
            media = args[1] if len(args) > 1 else True
            out_dir = args[2] if len(args) > 2 else None
            hdr = args[3] if len(args) > 3 else True
            opts = {"revisions": rev, "extract_media": media, "output_dir": out_dir, "include_header": hdr}
        else:
            opts = {}

        params = {
            "fmt": "md",
            "revisions": opts.get("revisions_mode") or opts.get("revisions", "markup"),
            "extract_media": opts.get("extract_media", True),
            "output_dir": opts.get("output_dir"),
            "include_header": opts.get("include_header", True),
        }
        self.queue_manager.add_item(Path(doc), "word_to_markdown", params=params)

    def _on_word_convert_markdown_to_word_requested(self, doc: str, *args, **kwargs) -> None:
        """Enqueues Markdown to Word conversion supporting dict or positional arguments."""
        if args and isinstance(args[0], dict):
            opts = args[0]
        elif kwargs:
            opts = kwargs
        else:
            opts = {}

        params = {
            "output_dir": opts.get("output_dir"),
            "companion_media_dir": opts.get("companion_media_dir"),
            "render_diagrams": opts.get("render_diagrams", True),
            "apply_3gpp_styles": opts.get("apply_3gpp_styles", True),
            "auto_detect_media": opts.get("auto_detect_media", True),
            "auto_open": opts.get("auto_open", False),
        }
        self.queue_manager.add_item(Path(doc), "markdown_to_word", params=params)

    def on_conversion_failed(self, error_msg: str):
        if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_failed'):
            self.word_tab.on_conversion_failed(error_msg)

    def on_conversion_success(self, out_path: str):
        if out_path == "SPECS_DB_PASS_ONE":
            self.log_message("🔄 Basic files loaded! You can now search while metadata downloads.", logging.INFO)
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(True)

        elif out_path == "SPECS_DB_PASS_TWO":
            self.log_message("✅ Deep metadata fully updated!", logging.INFO)
            if hasattr(self, 'specs_tab'):
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(False)
                else:
                    self.specs_tab.refresh_table()
            self.show_toast("Specifications Updated", "Deep metadata synchronization completed.", variant="success")

        elif out_path == "OPENED_IN_PPT":
            self.log_message("👁️ PowerPoint is open with your new slide. You can copy it directly.")
            self.show_toast("PowerPoint Updated", "Added diagram slide to open presentation.", variant="success")

        elif out_path == "SPECS_DB_UPDATED":
            self.log_message("🔄 Reloading specifications table...")
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
            self.show_toast("Specifications Reloaded", "Specification catalogue refreshed.", variant="success")

        elif out_path and out_path.startswith("MEETINGS_DB"):
            self.log_message("🔄 Reloading meetings table...", logging.INFO)
            if hasattr(self, 'meetings_tab'):
                # Phase 1 (FTP folders) and Phase 2 (DynaReport scheduled meetings)
                # can both add new meetings and therefore affect available WG choices.
                if out_path in ("MEETINGS_DB_PHASE_1", "MEETINGS_DB_PHASE_2"):
                    self.meetings_tab._populate_filters()
                self.meetings_tab.refresh_table()

                # Phase 2 (DynaReport dates) and Phase 3 (Docs/ TDoc bounds)
                # can both unblock Quick TDoc Jump lookups.
                if out_path in ("MEETINGS_DB_PHASE_2", "MEETINGS_DB_PHASE_3"):
                    self.meetings_tab.refresh_quick_tdoc_lookup()
            if (
                out_path in ("MEETINGS_DB_PHASE_2", "MEETINGS_DB_PHASE_3")
                and hasattr(self, "status_bar")
                and hasattr(self.status_bar, "refresh_world_clock")
            ):
                self.status_bar.refresh_world_clock()

        elif out_path:
            self.last_out_path = out_path
            self.code_tab.set_copy_path_enabled(True, out_path)

            if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_success'):
                self.word_tab.on_conversion_success(out_path)

            if hasattr(self, 'batch_tab') and hasattr(self.batch_tab, 'on_conversion_success'):
                self.batch_tab.on_conversion_success(out_path)

            file_name = Path(out_path).name
            if out_path.lower().endswith(('.svg', '.txt')):
                try:
                    os.startfile(out_path)
                    ext = Path(out_path).suffix[1:].upper()
                    self.log_message(f"👁️ Opened {ext} in default system viewer.")
                except Exception as e:
                    self.log_message(f"⚠️ Could not automatically open file: {e}")
            else:
                self.log_message(f"✅ Successfully saved/downloaded: {file_name}", logging.INFO)

            self.show_toast(
                f"Saved {file_name}",
                str(Path(out_path).parent),
                variant="success",
                actions=[
                    ToastAction("Open", lambda p=out_path: self._open_output_file(p)),
                    ToastAction("Show in Folder", lambda p=out_path: self._reveal_output_folder(p)),
                ],
            )

    def show_toast(self, title: str, message: str = "", variant: str = "info", actions=None, duration_ms: int = 6000):
        """Show a non-blocking toast notification if ToastManager is initialized."""
        if hasattr(self, "toast_manager") and self.toast_manager is not None:
            return self.toast_manager.show_toast(
                title=title,
                message=message,
                variant=variant,
                actions=actions,
                duration_ms=duration_ms,
            )
        return None

    def _open_output_file(self, path_str: str):
        try:
            if hasattr(os, "startfile"):
                os.startfile(path_str)
            else:
                webbrowser.open(Path(path_str).as_uri())
        except Exception as exc:
            self.log_message(f"⚠️ Could not open file '{path_str}': {exc}", logging.WARNING)

    def _reveal_output_folder(self, path_str: str):
        p = Path(path_str)
        try:
            if os.name == "nt" and p.exists():
                subprocess.Popen(["explorer", "/select,", str(p)])
            elif hasattr(os, "startfile"):
                os.startfile(str(p.parent if p.parent.exists() else p))
        except Exception as exc:
            self.log_message(f"⚠️ Could not reveal folder for '{path_str}': {exc}", logging.WARNING)
