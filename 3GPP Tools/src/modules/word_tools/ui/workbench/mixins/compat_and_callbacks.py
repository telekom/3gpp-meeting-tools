"""Drop handlers, backward-compatible delegating properties, and external worker callbacks."""

from __future__ import annotations

from typing import Optional


class WordTabCompatAndCallbacksMixin:
    """Preserves legacy delegating properties/presets and MainWindow conversion lifecycle hooks."""

    def _on_lo_dropped(self, files: list):
        if files:
            self.convert_doc_requested.emit(files[0], "docx")

    def _on_split_dropped(self, files: list):
        if files:
            self.split_doc_requested.emit(files[0])

    @property
    def _last_converted_md(self):
        return getattr(self.card_markdown, "_last_converted_md", None)

    @_last_converted_md.setter
    def _last_converted_md(self, val):
        self.card_markdown._last_converted_md = val
        self.card_markdown._last_md_output_path = val

    @property
    def _excerpt_scan_thread(self):
        return getattr(self.card_excerpt, "_excerpt_scan_thread", None)

    @property
    def _excerpt_extract_thread(self):
        return getattr(self.card_excerpt, "_excerpt_extract_thread", None)

    def _inspect_selected_author_document(self, file_path: str):
        if hasattr(self.card_author_renamer, "_inspect_selected_author_document"):
            return self.card_author_renamer._inspect_selected_author_document(file_path)

    def _apply_preset_3gpp(self):
        if hasattr(self.card_author_renamer, "_apply_preset_3gpp"):
            return self.card_author_renamer._apply_preset_3gpp()

    def _apply_preset_lang_only(self):
        if hasattr(self.card_author_renamer, "_apply_preset_lang_only"):
            return self.card_author_renamer._apply_preset_lang_only()

    def _apply_preset_anonymize(self):
        if hasattr(self.card_author_renamer, "_apply_preset_anonymize"):
            return self.card_author_renamer._apply_preset_anonymize()

    def _select_all_authors(self, select: bool):
        if hasattr(self.card_author_renamer, "_select_all_authors"):
            return self.card_author_renamer._select_all_authors(select)

    def _apply_author_preset(self, author_name: str, initials: Optional[str] = None):
        if hasattr(self.card_author_renamer, "_apply_author_preset"):
            return self.card_author_renamer._apply_author_preset(author_name, initials)

    def _derive_initials(self, name: str) -> str:
        if hasattr(self.card_author_renamer, "_derive_initials"):
            return self.card_author_renamer._derive_initials(name)
        return ""

    # -------------------------------------------------------------------------
    # External Worker Lifecycle Hooks (Invoked by MainWindow)
    # -------------------------------------------------------------------------
    def on_conversion_success(self, out_path: str):
        if hasattr(self.card_markdown, "on_conversion_success"):
            self.card_markdown.on_conversion_success(out_path)
        if hasattr(self.card_convert, "on_conversion_success"):
            self.card_convert.on_conversion_success(out_path)
        if hasattr(self.card_visio, "on_extraction_success"):
            self.card_visio.on_extraction_success(out_path)

    def on_conversion_failed(self, error_msg: str):
        if hasattr(self.card_markdown, "on_conversion_failed"):
            self.card_markdown.on_conversion_failed(error_msg)

    def on_processing_state_changed(self, is_processing: bool):
        if hasattr(self.card_markdown, "on_processing_state_changed"):
            self.card_markdown.on_processing_state_changed(is_processing)
