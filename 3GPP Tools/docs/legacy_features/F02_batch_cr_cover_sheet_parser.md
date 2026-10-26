# F-02: Batch CR Cover Sheet Parser & Tabular Excel Exporter

## 1. Executive Summary & Objective

In 3GPP Working Groups and Plenaries, a single meeting frequently handles between 500 and 2,500 Change Requests (CRs). A Change Request is submitted as a Microsoft Word document (`.docx` or legacy `.doc`) structured with a standardized front cover sheet followed by revision-marked text.

The cover sheet contains critical normative metadata:
* **Target Specification:** (e.g. `TS 23.501`)
* **CR Number & Revision:** (e.g. `CR 1245 rev 2`)
* **Category:** (`F` - Correction, `A` - Corresp. to earlier release, `B` - Addition of feature, `C` - Functional modification, `D` - Editorial)
* **Release:** (e.g. `Rel-18`, `Rel-19`)
* **Work Item Code:** (e.g. `FS_eSBA`, `5G_URLLC`)
* **Reason for Change:** High-level problem statement.
* **Summary of Change:** Technical resolution and architecture impacts.
* **Consequences if not approved:** Operational risks and specification ambiguities.
* **Clauses affected:** (e.g. `5.2.3, 6.4.1, Table 8.2-1`)

Currently, in **3GPP Tools**, delegates can use the AI TDoc Triage dialog to analyze one CR at a time. However, there is **no mechanism to batch-parse all CRs in a meeting or Work Item** into a consolidated, filterable Microsoft Excel spreadsheet.

The objective of **F-02** is to implement a high-speed, headless cover sheet extraction engine and batch exporter that compiles hundreds of CR cover sheets into a structured Excel workbook (`CR_export.xlsx`) with automated styling, column widths, and autofilters.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/parsing/word/pywin32.py` (functions: `parse_cr`, `parse_list_of_crs`, `CrMetadata`, `CrCategory`)
* `3GPP Meeting Helper/gui/tdocs_table.py` (function: `export_crs`)

### Data Structures & Legacy Extraction Logic

#### Legacy Data Model (`parsing/word/pywin32.py`)
```python
class CrCategory(Enum):
    F = 1  # Correction
    A = 2  # Corresponds to a correction in an earlier release
    B = 3  # Addition of feature
    C = 4  # Functional modification of feature
    D = 5  # Editorial modification

class CrMetadata(NamedTuple):
    spec: str
    cr: str
    rev: str
    current_version: str
    new_version: str
    category: CrCategory
    release: str
    work_item: str
    clauses_affected: str
    reason_for_change: str
    summary_of_change: str
    consequences_if_not_approved: str
    path: str
    ai: str
    tdoc: str
```

#### Table Cell Matrix Parsing (`parse_cr`)
The legacy implementation used Word COM (`win32com.client.Dispatch("Word.Application")`) to inspect the first table in the CR document:
1. **Spec Number:** Cell `(1, 1)` -> Text like `"23.501"` (stripped of `"TS"` / `"TR"` prefixes).
2. **CR Number:** Cell `(1, 3)` -> e.g. `"1420"`.
3. **Revision:** Cell `(1, 5)` -> e.g. `"1"` or `"-"`.
4. **Category Checkboxes:** Cell `(5, 1)` contained checkboxes `F | A | B | C | D`. It checked for `'X'` characters inside the respective box cells.
5. **Release Checkboxes:** Checked row 6 for `'X'` in release columns (Rel-15 through Rel-20).
6. **Descriptive Text Blocks:** Scanned succeeding rows for cell labels matching `"Reason for change:"`, `"Summary of change:"`, `"Consequences if not approved:"`, and `"Clauses affected:"`, extracting the adjacent or nested text blocks.

#### Excel Generation & Formatting (`gui/tdocs_table.py`)
```python
crs_df = parse_list_of_crs(file_path_list)
crs_df = crs_df.set_index('TDoc')
crs_df.to_excel(excel_export_filename, sheet_name="CRs", engine='xlsxwriter')

wb = open_excel_document(excel_export_filename)
set_first_row_as_filter(wb)
vertically_center_all_text(wb)
set_wrap_text(wb)
set_column_width('A', wb, 11)  # TDoc
set_column_width('B', wb, 9)   # Spec
set_column_width('C', wb, 9)   # CR#
set_column_width('F', wb, 20)  # WI
set_column_width('I', wb, 50)  # Reason for Change
```

### Bottlenecks of Legacy Implementation
* **Heavy Word COM Dependency:** Required launching Microsoft Word in the background for every single file. Parsing 50 CRs took 2–3 minutes and occasionally hung if Word displayed a security alert or macro prompt.
* **Single-Threaded COM Apartment:** Word COM could not be easily multi-threaded without complex COM marshaling (`pythoncom.CoInitialize()`).
* **Table Coordinate Fragility:** Relied on fixed row/column indices which occasionally failed on non-standard rapporteur CR templates or older meeting files.

---

## 3. Current State in 3GPP Tools

* **AI TDoc Triage:** `src/modules/meetings/core/tdoc_triage_worker.py` extracts text from `.docx` files via `python-docx` for Ollama prompting.
* **Document Resolution:** `src/modules/meetings/core/tdoc_file_handler.py` provides high-speed, parallel pre-download and unzipping of meeting contributions.
* **Missing Capability:** There is no batch cover-sheet parser that aggregates structured CR metadata across multiple files directly into Excel without LLM intervention.

---

## 4. Modern Architecture & Design Specification

### Architectural Pipeline

```mermaid
flowchart TD
    UI["TDocs Window: Select CRs or Filter by WI / Status"] --> Trigger["Menu: Reports > Export CR Cover Sheets to Excel"]
    Trigger --> Worker["CRCoverSheetExportWorker (QThread)"]
    Worker --> Pool["ThreadPoolExecutor (5-8 Workers)"]
    
    subgraph HeadlessParsing["Headless OpenXML / python-docx Engine"]
        Pool --> Unzip["Unzip .zip -> Locate .docx/.doc"]
        Unzip --> FastDocx{"Is .docx or .doc?"}
        FastDocx -- ".docx" --> OpenXML["Direct lxml / python-docx Table Parser (0.05s / file)"]
        FastDocx -- ".doc" --> COMBridge["Fallback: Word COM / LibreOffice doc2docx"]
        OpenXML --> Extractor["Extract Normalized Metadata (Regex & Cell Anchors)"]
        COMBridge --> Extractor
    end

    Extractor --> Aggregator["Pandas DataFrame Assembly"]
    Aggregator --> OpenPyXL["openpyxl Workbook Generator"]
    OpenPyXL --> Styling["Apply 3GPP Executive Theme:
    - Bold Frozen Header
    - Auto-filter on Row 1
    - Soft Status / Category Badges
    - Left-Aligned Text Wrapping on Explanations"]
    Styling --> Output["Save to Meeting/Agenda/CR_export.xlsx & Auto-Launch"]
```

### Extraction Algorithm: Anchor-Based Flexible Parsing

Instead of rigid row/column coordinates (which break across TSG templates), the modern extractor will use **label-anchor detection**:

```python
def extract_cr_metadata_from_docx(doc_path: Path) -> dict:
    doc = Document(doc_path)
    tables = doc.tables
    if not tables:
        return {}

    cover_table = tables[0]
    metadata = {
        "spec": "", "cr_num": "", "rev": "", "cat": "", "release": "",
        "work_item": "", "clauses": "", "reason": "", "summary": "", "consequences": ""
    }

    full_text = []
    for row in cover_table.rows:
        row_text = [cell.text.strip() for cell in row.cells]
        for idx, text in enumerate(row_text):
            t_lower = text.lower()
            if "spec" in t_lower and "ts" in t_lower and idx + 1 < len(row_text):
                metadata["spec"] = sanitize_spec(row_text[idx+1])
            elif "cr num" in t_lower or "cr" == t_lower:
                metadata["cr_num"] = extract_digits(row_text[idx+1])
            elif "rev" in t_lower:
                metadata["rev"] = extract_rev(row_text[idx+1])
            elif "work item" in t_lower:
                metadata["work_item"] = row_text[idx+1]
            elif "reason for change" in t_lower:
                metadata["reason"] = row_text[idx+1]
            elif "summary of change" in t_lower:
                metadata["summary"] = row_text[idx+1]
            elif "consequences if not approved" in t_lower:
                metadata["consequences"] = row_text[idx+1]
            elif "clauses affected" in t_lower:
                metadata["clauses"] = row_text[idx+1]
                
    return metadata
```

### Excel Workbook Design (`openpyxl`)
The output spreadsheet will match the professional formatting used in 3GPP Tools meeting statistics:
* **Header Fill:** Slate Navy `#1E293B` with bold white text `#FFFFFF`.
* **Row Striping:** Alternating `#FFFFFF` and `#F8FAFC`.
* **Category Pill Colors:** 
  * `Cat B` (Feature addition): Soft Emerald `#DEF7EC` / `#03543F`.
  * `Cat F` (Correction): Soft Blue `#E1EFFE` / `#1E429F`.
  * `Cat C` (Functional modification): Soft Amber `#FEF08A` / `#713F12`.
* **Auto-Filter & Freeze Panes:** Frozen top row, enabled AutoFilter on columns A through N.
* **Column Widths & Auto-Fit:** Explicit column widths with text wrap enabled for Reason, Summary, and Consequences.

---

## 5. Implementation Roadmap & Acceptance Criteria

### Step-by-Step Implementation Tasks
1. [ ] Create `src/modules/meetings/core/cr_cover_sheet_parser.py` implementing anchor-based `python-docx` extraction.
2. [ ] Add `CRCoverSheetExportWorker(QThread)` in `src/modules/meetings/core/exporter_thread.py` with multi-threaded pre-scan and progress reporting.
3. [ ] Add `"Export CR Cover Sheets to Excel..."` action in `src/modules/meetings/ui/tdocs_menus.py` under the `📊 Reports ▾` menu and context menu.
4. [ ] Build `openpyxl` styling helper in `src/modules/meetings/core/excel_exporter.py` for professional formatting.

### Acceptance Criteria
* [x] Capable of parsing 100 CR documents in under 5 seconds (headless OpenXML without Word COM).
* [x] Accurate extraction of Spec, CR#, Rev, Category, Release, Work Item, Reason, Summary, and Consequences.
* [x] Supports both single selection, multi-selection, and "All CRs in current filtered view".
* [x] Resulting Excel workbook opens cleanly in Excel, LibreOffice, and Google Sheets with auto-filters and wrapped text.
