# Word Tools & Contribution Studio Guide (`src/modules/word_tools`)

Developer and AI agent guide for the Word Studio, Document Comparison, CR/pCR Workbench, Bidirectional Markdown/PowerPoint Converters, Document Sanitizer, and MIP Sensitivity Labeling.

---

## 1. Directory Map

```text
src/modules/word_tools/
├── AGENTS.md
├── core/
│   ├── cr_pcr/                                  # Modular CR/pCR generator package (see cr_pcr/README.md)
│   ├── cr_pcr_generator.py                      # Re-export shim -> core/cr_pcr/ package
│   ├── styles/                                  # Modular 3GPP styles registry & packager (see styles/README.md)
│   ├── styles_3gpp.py                           # Re-export shim -> core/styles/ package
│   ├── document_opener.py / word_comparator.py  # Safe Word COM opening, snapshotting & Track Changes diff
│   ├── document_sanitizer.py / author_renamer.py # F-07 proofing language, rev purge & author hub
│   ├── docx_splitter.py / word_excerpt_extractor.py / word_excerpt_fragment_worker.py / word_excerpt_naming.py
│   ├── md_to_docx/                              # Modular Markdown -> Word package (see md_to_docx/README.md)
│   ├── markdown_to_docx.py                      # Re-export shim -> core/md_to_docx/ package
│   ├── converter/                               # Modular Word format converter package (see converter/README.md)
│   ├── presentation/                            # Modular PPTX/PPT/PDF -> Markdown package (see presentation/README.md)
│   ├── presentation_to_markdown.py (shim) / word_converter.py (shim) / libreoffice_converter.py
│   ├── sensitivity/                             # Modular Purview/MIP policy & labeling package (see sensitivity/README.md)
│   ├── sensitivity_label.py                     # Re-export shim -> core/sensitivity/ package
│   ├── word_config.py / word_extractor.py / word_media_repair.py
└── ui/
    ├── word_tabs.py                             # WordExtractorTab (2-tier ribbon + segmented workbenches)
    ├── document_selector_pane.py                # DocumentSelectorPane 2.0 & ModernDropZone
    ├── cr_pcr/                                  # Modular CrPcrDialog package (see cr_pcr/README.md)
    ├── cr_pcr_dialog.py / batch_label_dialog.py / word_settings_dialog.py   # (shims -> cr_pcr/ & word_settings/)
    ├── word_settings/                           # Modular WordSettingsDialog package (see word_settings/README.md)
    └── cards/                                   # compare_card, markdown_card, cr_pcr_card, sanitizer_card, excerpt_card, converter_card
```

---

## 2. UI Shell, Cards & Input Widgets

- **2-Tier Header Ribbon & Segmented Cards (`word_tabs.py`, `ui/cards/`, `word_settings_dialog.py`):**
  - Row 1: live Active Word monitor pill (`🪟 Active Word`), `3gpp_70.dotx` macro badge, LibreOffice status pill, `⚙️ Word Settings ▾` (`WordSettingsDialog` configuring corporate sensitivity labels, default company, templates, naming schemes).
  - Row 2: segmented pills (`Compare Documents`, `Word → Markdown`, `3GPP CR && Templates`, `Sanitizer && Revisions`, `Slice && Excerpt`, `Format Converter`, `Visio && Media Extractor`).
- **`DocumentSelectorPane 2.0` & `ModernDropZone`:** Accepts local files, foreground Word COM docs, and remote/SharePoint URLs. Uses anti-aliased `QPainter` dashed borders, file archetype badges (`.docx`, `.md`, `.puml`, `.pptx`, `.vsdx`), OpenXML sensitivity badges, and morphed Active Document Card with `[🔄 Replace]` / `[✕ Clear]`.
- **Backward-Compatible `WordExtractorTab` API:** Preserve aliases (`btn_swap_slots`, `pane_a`, `pane_b`, `pane_cr`, `pane_visio`, `pane_convert`, `pane_markdown`, `run_markdown_btn`, `md_revisions_combo`, `author_table`) and delegating properties/methods (`_last_converted_md`, `_excerpt_scan_thread`, `_excerpt_extract_thread`, `_inspect_selected_author_document`, `_apply_preset_*`).

## 3. Word COM Automation, Comparison & Active Polling

- **Document Comparison (`word_comparator.py`, `word_tabs.py`):** Uses Word COM Track Changes with 1-click **"⇄ Swap A ↔ B"**, `get_active_word_document`, non-destructive snapshotting (`SaveCopyAs` with `shutil.copy2` disk fallback on OLE error `0x800a1704`, and `doc.SaveAs2(..., AddToRecentFiles=False)` for headless instances), formatting diff toggle, and keep-open options.
- **Foreground Word Status Polling (`_poll_active_word` in `word_tabs.py`, `document_opener.py`):**
  - **Must** pass `allow_unsaved_snapshot=False` (never write disk snapshots during background polling).
  - **Must** be visibility-gated (`if not force and not self.isVisible(): return`) with instant `showEvent` refresh.
  - Guard all live COM property accesses (`word.Documents.Count`, `ActiveDocument`, `doc.Name`, `FullName`, `Saved`) in `try...except` against `0x80010001` modal/busy rejections.

## 4. 3GPP CR/pCR Workbench, Style Preservation & Excerpts

- **CR/pCR Generation (`cr_pcr_dialog.py`, `cr_pcr_generator.py`, `styles_3gpp.py`, `word_excerpt_naming.py`):**
  - Launches from baseline specs or active Word drafts; manages `3gpp_70.dotx` macro template install/uninstall and 3GPP FTP form links (`Template_3GPP_CR.docx`, `Pseudo-CR-Form-v0-1.zip`).
  - Official MCC pseudo-CR formatting: `CRCoverPage` header (12pt bold Arial, non-italic TDoc ID), `<w:pBdr><w:bottom w:sz="6"/>` divider, 3.5cm hanging indent (`<w:ind w:left="1985" w:hanging="1985"/>`), Arial Italic abstract, XML escaping, and 4-sided boxed blue 14pt Arial change separators (`* * * First Change * * * *`, `* * * End of Changes * * * *`).
  - Differentiated export naming (`word_excerpt_naming.py`) with presets, reactive filename preview in `CrPcrDialog`, `↺ Reset`, placeholder omission (`Draft-TDoc`, `N/A`), and persistence in `word_config.json`.
- **Non-Destructive Style Merging (`merge_and_promote_3gpp_styles`, `ensure_3gpp_package_styles`):** Preserves existing table styles, borders (`<w:tblBorders>`, `<w:tcBorders>`), widths, and shading while guaranteeing Word's `TableGrid` style without clobbering `word/styles.xml`, `theme1.xml`, or `fontTable.xml`. Retries reads and atomic `os.replace` up to 25 attempts (with `shutil.copyfile` fallback) against OneDrive/AV file locks.
- **Interactive Heading Excerpt & Slicing (`docx_splitter.py`, `word_excerpt_extractor.py`, `WordExcerptThread`):** Hierarchical `[H1]`/`[H2]`/`[H3]` browser, search match counter, deduplication of ancestor/child clauses, dual `.docx` + `.md` (`{stem}_media/` figures via `convert_document_to_markdown`), and full staging isolation inside `get_temp_root()` before copying to OneDrive targets.

## 5. Bidirectional Markdown ⇄ Word, Decks, Sanitizer & MIP Labeling

- **Word → Markdown:** Headless extraction with revision mode (`Inline Markup`, `Accept All Changes`, `Reject All Changes`), `{doc_name}_media/` figure export, EMF/WMF PNG rasterization, and embedded `.vsdx` linking.
- **PowerPoint / PDF Deck → Markdown (`presentation/` package & `presentation_to_markdown.py` shim):** 4-tier engine: (1) Windows PowerPoint COM (1080p `slide_NN.png` + titles, bullets, `HasTable`, `NotesPage`); (2) Headless OpenXML `.pptx` parser (`xml.etree.ElementTree`); (3) PDF slide engine (`pypdfium2` + `pypdf`); (4) Headless LibreOffice fallback for `.ppt`. Outputs `## Slide N`, image embeds, GFM tables, and `> [!NOTE] Speaker Notes`.
- **Markdown → Word (`md_to_docx/` package & `markdown_to_docx.py` shim):** Headless OpenXML synthesis resolving relative/`{doc_stem}_media/` images (auto-scaled to printable margins), SVG $\to$ PNG via QtSvg, ````plantuml```` blocks via `plantuml.jar`, `3gpp_70.dotx` + `TableGrid` styling, GFM callouts, and MIP sensitivity labels.
- **Meeting Export Completeness & Recovery (`export_manifest.py`, `presentation_to_markdown.py`):** `is_export_complete()` checks non-zero size, `>=120` bytes (unless converter-stamped), absence of `"Document conversion was not available"` / `"__-________"` corrupt prefixes, `"## Slide "` presence for decks, and on-disk slide PNGs. `plan_incremental_session()` queues incomplete exports into `tdocs_to_update`. `_clean_tdoc_prefix` uses `re.sub(r'[^\w\-]', '_', ...)` and purges old corrupted images; `_clean_for_json` serializes `set`/`tuple`/`Path` in `row_data`; `_find_word_doc` prioritizes `.docx`, `.docm`, `.pptx`, `.doc`, `.ppt` over `.pdf`.
- **Sanitizer & Author Hub (F-07, `document_sanitizer.py`, `author_renamer.py`):** Pure Python OpenXML ZIP engine normalizing proofing language across all stories/styles, purging trivial `<w:rPrChange>` language-only balloons while preserving technical revisions, consolidating reviewer names (`'Deutsche Telekom'`, `'Editor'`, `'Rapporteur'`, `'Anonymous'`) + initials + `dc:creator`/`cp:lastModifiedBy`, auditing in `<150ms`, and creating `.bak` backups.
- **MIP / Purview Labeling (F-06, `sensitivity_label.py`, `batch_label_dialog.py`):** Auto-discovers Windows/Office policy (`%LOCALAPPDATA%\Microsoft\Office\CLP\policy*.gz`, enabled `File`-scoped labels, `TenantId` $\to$ `SiteId`), defaults to lowest-order unencrypted label (clean no-op if absent), marks RMS-encrypted labels (`🔒 [Encrypted]`), inspects `docProps/custom.xml` headlessly, supports folder/multi-file `BatchSensitivityLabelDialog`, and applies labels automatically across `document_opener.py`, `word_comparator.py`, `cr_pcr_generator.py`, and `docx_splitter.py`.
