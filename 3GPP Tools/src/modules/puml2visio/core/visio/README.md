# Visio COM Conversion Package (`core/visio`)

Acyclic modular breakdown behind the backward-compatible `modules.puml2visio.core.visio_converter` shim:

- **`svg_sanitize.py`**: Pre-import `<tspan>` attribute cleanup (`sanitize_svg_for_visio`) preventing Visio import errors.
- **`visio_pipeline.py`**: Stages 1–10 Visio COM pipeline (`convert_svg_to_vsdx`), group flattening, background removal, text shape normalization, page resizing, embedded `PlantUML Source` page creation, and ordered COM proxy release before `doc.Close()`/`visio.Quit()`.
- **`threads.py`**: `VisioReaderThread`, `ConverterThread`, and `SvgConverterThread` QThread workers with thread-local `pythoncom` apartment lifecycle.
