# F-17: 3GPP Guru Web Viewer Integration

## 1. Executive Summary & Objective

When reading or cross-referencing 3GPP specifications, delegates frequently need quick clause lookups in a browser without downloading, unzipping, and launching heavy Microsoft Word `.docx` documents. 

[3GPP Guru](https://3gpp.guru) is a widely used third-party web service that hosts pre-rendered, syntax-highlighted HTML conversions of official 3GPP specifications with hyperlinked cross-references (e.g. clicking `TS 23.501 clause 5.2.3` within `TS 23.502` jumps directly to that clause in the browser).

The legacy application included a 1-click integration in the Specifications table (`gui/specs_table.py`) that synthesized the exact URL structure for 3GPP Guru and launched the matching release in the user's default web browser.

In **3GPP Tools**, the Specifications and Spec Search modules provide direct links to the official 3GPP Portal (`portal.3gpp.org`) and 3GPP FTP directories, but do **not** provide direct deep links to 3GPP Guru.

The objective of **F-17** is to restore **"🌐 Open in 3GPP Guru"** deep links across the Specifications tab, Evolution Matrix, and Spec Search inspector.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/gui/specs_table.py` (functions: `get_3gpp_guru_url_for_version_text`, `SpecsTable.on_double_click`)

### Legacy Implementation Details

```python
# gui/specs_table.py
def get_3gpp_guru_url_for_version_text(spec_id: str, version_text: str) -> str:
    """
    Returns the URL for the matching specification in 3GPP Guru.
    Example output: https://3gpp.guru/trts/Rel-19/23502-j00.html
    """
    file_version = version_to_file_version(version_text)  # e.g. "19.0.0" -> "j00"
    version_metadata = file_version_to_version_metadata(file_version)
    spec_name_without_dots = spec_id.replace('.', '')     # e.g. "23.502" -> "23502"

    url = f"https://3gpp.guru/trts/Rel-{version_metadata.release}/{spec_name_without_dots}-{file_version}.html"
    return url
```

#### Conversion Rules for 3GPP Version Formatting
3GPP encodes release letters in filenames:
* Release 8: `800`
* Release 9: `900`
* Release 10: `a00`
* Release 11: `b00`
* Release 18: `i00`
* Release 19: `j00`
* Release 20: `k00`

---

## 3. Current State in 3GPP Tools

* `src/modules/specifications/ui/spec_menu.py` has menu items:
  * `🌐 Open in 3GPP Portal`
  * `📂 Open on 3GPP FTP`
  * `📋 Copy Citation`
* **Missing Feature:** No option to open the specification version directly in 3GPP Guru.

---

## 4. Modern Architecture & Design Specification

### URL Generation Algorithm (`src/core/utils/spec_url_helpers.py`)

```python
def get_3gpp_guru_url(spec_number: str, version: str) -> str:
    """
    Constructs a 3gpp.guru web viewer URL.
    spec_number: "23.501" or "23501"
    version: "v18.4.0" or "18.4.0" or "i40"
    """
    clean_spec = spec_number.replace("TS", "").replace("TR", "").replace(".", "").strip()
    
    # Resolve 3GPP release letter encoding
    # 10->a, 11->b, 12->c, 13->d, 14->e, 15->f, 16->g, 17->h, 18->i, 19->j, 20->k
    file_version = convert_version_to_file_code(version)
    release_num = extract_release_number(version)
    
    return f"https://3gpp.guru/trts/Rel-{release_num}/{clean_spec}-{file_version}.html"
```

### UI Integration Touchpoints
1. **Specifications Tab:** Right-click row or click Column 0 / Column 2 action menus (`⋮` / `Get Word ▾`) -> Add: `"🌐 Open in 3GPP Guru"`.
2. **Spec Search Tab:** In `SpecClauseInspector` control bar and Evolution Matrix right-click menu -> Add: `"🌐 View Spec in 3GPP Guru"`.
3. **Spec Info Dialog (`spec_info_dialog.py`):** Add an interactive button in the dialog footer alongside `3GPP Portal` and `FTP`.

---

## 5. Implementation Roadmap & Acceptance Criteria

### Step-by-Step Implementation Tasks
1. [ ] Implement `get_3gpp_guru_url` in `src/core/utils/spec_url_helpers.py`.
2. [ ] Add `"🌐 Open in 3GPP Guru"` action to `src/modules/specifications/ui/spec_menu.py`.
3. [ ] Add action to `src/modules/spec_search/ui/spec_search_tabs.py`.

### Acceptance Criteria
* [x] Accurately constructs URLs for Releases 8 through 20 (handling letter encodings `a` through `k`).
* [x] Single click launches the browser directly to the specified version.
* [x] Zero network or scraping overhead (pure URL template synthesis).
