# Font Scaling Package (`core.ui.font_scaling`)

Acyclic modular package extracted from `core/ui/font_scale.py` (thin backward-compatible re-export shim):

- `state.py`: Canonical typography scale constants, `FontScalePreset`/`FONT_SCALE_PRESETS`, shared mutable scaling globals, `_FontScaleNotifier`, pixel/row-height math helpers, and `QSettings` persistence.
- `reader_zoom.py`: Live `Ctrl+Wheel` / `Ctrl+0` reader/editor zoom (`zoom_text_widget`, `bind_reader_zoom_settings`, `install_reader_zoom`) and `create_help_tips_pill()`.
- `qss_and_hooks.py`: `scale_stylesheet()`, `scale_html_fonts()`, Qt widget/HTML hooks (`install_font_scaling`), `set_unscaled_stylesheet()`, `register_scaled_row_height()`, and `apply_font_scale()`.
