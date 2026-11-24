# Spec History Extraction Package (`core/history_extract`)

Acyclic modular breakdown behind the backward-compatible `modules.specifications.core.spec_history_extractor` shim:

- **`xml_utils.py`**: OpenXML namespaces (`NS`), regex patterns (`RE_TDOC_TOKEN`, `RE_VERSION`), cell/row extraction (`parse_cell_text`, `expand_row_cells`), header mapping (`detect_header_columns`), and `extract_cr_cover_sheet_metadata`.
- **`extractor.py`**: `SpecHistoryExtractor` for resolving multi-part spec documents and extracting normalized Change History tables via `lxml`.
- **`threads.py`**: `SpecHistoryExtractThread` and `SpecHistoryBatchScanThread` background Qt workers.
