# F-18: 3GPP Forge APIs Direct Launcher

## 1. Executive Summary & Objective

Starting with 5G System (5GS) Release 15, 3GPP service-based architecture (SBA) specifications (TS 29.500 series, TS 29.518, TS 29.502, TS 29.571, etc.) define normative RESTful APIs using OpenAPI 3.0 YAML/JSON descriptions.

Rather than maintaining these complex schema definitions solely in Word documents, ETSI and 3GPP host the official, version-controlled machine-readable API definitions on the **3GPP Forge GitLab** portal:
👉 `https://forge.3gpp.org/rep/all/apis`

In the legacy application, `_3GPP_Delegate_Helper.py` provided a prominent launcher button **"3GPP APIs"** on the main dashboard, enabling telecommunications protocol architects and API designers to jump directly to official OpenAPI definitions with one click.

In **3GPP Tools**, while the Protocol Matrix and Evolution tabs cover NAS, ASN.1 (RRC/NGAP), GTP-U, and PFCP, there is no direct link to 3GPP Forge APIs in the main interface.

The objective of **F-18** is to restore quick launcher links to 3GPP Forge OpenAPI repositories within relevant protocol and specification menus.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/_3GPP_Delegate_Helper.py` (button: `3GPP APIs`, line 70)
* `3GPP Meeting Helper/server/common/server_utils.py` (`apis_3gpp_forge_url = 'https://forge.3gpp.org/rep/all/apis'`)

### Legacy Implementation Details
```python
# _3GPP_Delegate_Helper.py
launch_3gpp_apis_url = ttk.Button(
    main_frame,
    width=button_width,
    text='3GPP APIs',
    command=lambda: open_url(apis_3gpp_forge_url)
)
```

---

## 3. Current State in 3GPP Tools

* No link or launcher button exists for 3GPP Forge OpenAPI repositories in the main window or protocol tabs.

---

## 4. Modern Architecture & Design Specification

### Proposed Integration Touchpoints

1. **Protocols Tab (`src/modules/nas/ui/nas_tabs.py`):**
   * In the top control bar or `⚙️ Manage ▾` menu: Add action `"🌐 3GPP Forge OpenAPIs (forge.3gpp.org)"`.
2. **Specifications Tab (`src/modules/specifications/ui/`):**
   * If a selected specification belongs to the TS 29.5xx series (5G Core SBA APIs), dynamically display a button or menu action: `"🌐 View OpenAPI Specs on 3GPP Forge"`, linking directly to `https://forge.3gpp.org/rep/all/apis`.
3. **Application Help / External Resources Menu:**
   * In the main system bar or Help menu: Add a curated list of standards developer portals:
     * `3GPP Portal (portal.3gpp.org)`
     * `3GPP Forge APIs (forge.3gpp.org)`
     * `3GPP Guru Viewer (3gpp.guru)`
     * `3GPP Work Plan & GANTT (3gpp.org/ftp/Information/WORK_PLAN)`

---

## 5. Implementation Roadmap & Acceptance Criteria

### Step-by-Step Implementation Tasks
1. [ ] Add `FORGE_APIS_URL = "https://forge.3gpp.org/rep/all/apis"` to `src/core/config/paths.py`.
2. [ ] Add menu action to `src/modules/nas/ui/` and `src/modules/specifications/ui/`.

### Acceptance Criteria
* [x] 1-click action opens the default web browser directly to 3GPP Forge APIs.
* [x] Seamlessly integrated without cluttering the UI.
